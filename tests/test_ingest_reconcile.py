"""Synthetic-only fixtures; no user/business data is embedded in tests."""
import io
from pathlib import Path
import sys
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from benchmark.ingest import HEADERS, EXTRA_HEADERS, read_analysis, read_collection, merge_records
from benchmark.reconcile import QUANTITY_FIELDS, FEE_FIELDS, reconcile_records

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def workbook(sheets):
    """Minimal synthetic OOXML bytes, generated only as test input."""
    out = io.BytesIO()
    wb = ET.Element('{'+NS+'}workbook')
    shs = ET.SubElement(wb, '{'+NS+'}sheets')
    rels = ET.Element('{'+PKG+'}Relationships')
    with ZipFile(out, 'w') as archive:
        for i, (name, rows, merges) in enumerate(sheets, 1):
            ET.SubElement(shs, '{'+NS+'}sheet', {'name': name, 'sheetId': str(i), '{'+REL+'}id': 'rId'+str(i)})
            ET.SubElement(rels, '{'+PKG+'}Relationship', {'Id': 'rId'+str(i), 'Target': f'worksheets/sheet{i}.xml', 'Type': REL+'/worksheet'})
            root = ET.Element('{'+NS+'}worksheet')
            data = ET.SubElement(root, '{'+NS+'}sheetData')
            for rn, vals in sorted(rows.items()):
                row = ET.SubElement(data, '{'+NS+'}row', {'r': str(rn)})
                for c, value in enumerate(vals, 1):
                    if value is None:
                        continue
                    cell = ET.SubElement(row, '{'+NS+'}c', {'r': col(c)+str(rn)})
                    if isinstance(value, tuple) and value[0] == 'formula':
                        ET.SubElement(cell, '{'+NS+'}f').text = value[1]
                        if len(value) > 2:
                            ET.SubElement(cell, '{'+NS+'}v').text = str(value[2])
                    elif isinstance(value, tuple) and value[0] == 'error':
                        cell.set('t', 'e'); ET.SubElement(cell, '{'+NS+'}v').text = value[1]
                    elif isinstance(value, str):
                        cell.set('t', 'inlineStr'); text = ET.SubElement(cell, '{'+NS+'}is')
                        ET.SubElement(text, '{'+NS+'}t').text = value
                    else:
                        ET.SubElement(cell, '{'+NS+'}v').text = str(value)
            if merges:
                node = ET.SubElement(root, '{'+NS+'}mergeCells', {'count': str(len(merges))})
                for merge in merges:
                    ET.SubElement(node, '{'+NS+'}mergeCell', {'ref': merge})
            archive.writestr(f'xl/worksheets/sheet{i}.xml', ET.tostring(root))
        archive.writestr('xl/workbook.xml', ET.tostring(wb))
        archive.writestr('xl/_rels/workbook.xml.rels', ET.tostring(rels))
    return out.getvalue()


def record(station='测试甲', month=1, source='old.xlsx', date=None):
    v = [0.] * 71
    v[:8] = ['source', f'{month}月份', station, station+'调度', '测试区', '测试类型', 10., date]
    v[66] = '测试公司'
    v[67:71] = [None, None, None, '原备注']
    return {'values': v, 'company': '测试公司', 'month': month, 'year': 2026,
            'source': source, 'sheet': 'Sheet2', 'source_row': 5,
            'record_id': source+'|'+station+'|'+str(date),
            'present_fields': list(range(67)), 'field_states': {str(i): 'present' for i in range(67)}}


class IngestionTests(unittest.TestCase):
    def test_analysis_preserves_71_columns_and_original_provenance(self):
        r = record(); r['values'][68:71] = ['存量', -3.25, '保留的历史核对备注']
        headers = HEADERS + [EXTRA_HEADERS[i] for i in range(67, 71)]
        data = workbook([('Sheet2', {4: headers, 5: r['values']}, [])])
        result = read_analysis(data, '2026_analysis.xlsx')
        self.assertEqual(len(result['records']), 1)
        self.assertEqual(result['records'][0]['values'][:71], r['values'])
        self.assertEqual(len(result['headers']), 67)

    def test_analysis_keeps_partial_and_multiphase_commission_observations(self):
        a, b = record('测试甲'), record('测试乙')
        a['values'][7] = 2010
        b['values'][7] = '一期：2010.1.1\n二期：2012.2.2'
        result = read_analysis(workbook([('Sheet2', {4: HEADERS, 5: a['values'][:67], 6: b['values'][:67]}, [])]), '2026_old.xlsx')
        self.assertEqual(result['records'][0]['values'][7], 2010)
        self.assertEqual(result['records'][1]['values'][7], b['values'][7])

    def test_variable_column_count_and_explicit_capacity_unit(self):
        indexes = list(range(1, 67)); indexes.remove(66)
        headers = [HEADERS[i] for i in indexes]
        headers[indexes.index(6)] = '装机容量（万千瓦）'
        headers.insert(7, '新增加的说明列')
        source = record()['values']; source[6] = 2.5
        row = [source[i] for i in indexes]; row.insert(7, '文字说明')
        result = read_collection(workbook([('测试公司', {4: headers, 5: row}, [])]), '2026_collect.xlsx', HEADERS, 2026)
        self.assertEqual(result['records'][0]['values'][6], 25)
        self.assertEqual(result['records'][0]['values'][20], source[20])
        self.assertEqual(result['records'][0]['normalizations'][0]['factor'], 10)

    def test_merged_identity_only_numeric_not_broadcast_and_blank_break(self):
        indexes = list(range(1, 66))
        headers = [HEADERS[i] for i in indexes]
        a = record('测试甲')['values']; b = record('测试乙')['values']; b[1] = b[6] = b[63] = None
        c = record('测试丙')['values']; c[1] = None
        data = workbook([('测试公司', {4: headers, 5: [a[i] for i in indexes], 6: [b[i] for i in indexes], 8: [c[i] for i in indexes]}, ['A5:A6', 'F5:F6', 'BK5:BK6'])])
        result = read_collection(data, '2026_collect.xlsx', HEADERS, 2026)
        self.assertEqual(len(result['records']), 2)
        self.assertEqual(result['records'][1]['month'], 1)
        self.assertIsNone(result['records'][0]['values'][6])
        self.assertIsNone(result['records'][1]['values'][6])
        self.assertEqual(result['records'][0]['shared_numeric_observations'][0]['raw_value'], 10)
        self.assertIn('merged_numeric_not_broadcast', {x['code'] for x in result['warnings']})
        self.assertIn('missing_key', {x['code'] for x in result['warnings']})

    def test_errors_and_missing_formula_cache_are_not_zero(self):
        r = record(); r['values'][21] = ('error', '#DIV/0!'); r['values'][59] = ('formula', 'U5+AD5')
        result = read_analysis(workbook([('Sheet2', {4: HEADERS, 5: r['values'][:67]}, [])]), '2026_analysis.xlsx')
        rec = result['records'][0]
        self.assertIsNone(rec['values'][21]); self.assertIsNone(rec['values'][59])
        self.assertEqual(rec['field_states']['59'], 'formula_missing_cache')
        self.assertEqual(rec['original_errors']['21'], '#DIV/0!')

    def test_bad_energy_unit_capacity_is_not_guessed(self):
        headers = list(HEADERS); headers[6] = '装机容量（万千瓦时）'
        result = read_analysis(workbook([('Sheet2', {4: headers, 5: record()['values'][:67]}, [])]), '2026_analysis.xlsx')
        self.assertIsNone(result['records'][0]['values'][6])
        self.assertIn('capacity_energy_unit', {x['code'] for x in result['warnings']})

    def test_year_mismatch_rejected(self):
        data = workbook([('测试公司', {4: HEADERS, 5: record()['values'][:67]}, [])])
        with self.assertRaisesRegex(ValueError, '年度|年份|年不一致'):
            read_collection(data, '2025_collect.xlsx', HEADERS, 2026)

    def test_collection_does_not_reimport_helper_business_or_control_rows(self):
        base = record()['values'][:67]
        helper_headers = HEADERS + ['记录用途', '小时源变更']
        controller = [None] * 67
        for j in [1, 2, 3, 4, 5, 7, 66]:
            controller[j] = base[j]
        helper_rows = {6: helper_headers, 7: base + ['实际记录', 0], 8: controller + ['月份控制', 0]}
        data = workbook([('Sheet2', {4: HEADERS, 5: base}, []),
                         ('小时联动源', helper_rows, []), ('重命名的派生表', helper_rows, [])])
        result = read_collection(data, '2026_uploaded_analysis.xlsx', HEADERS, 2026)
        self.assertEqual(len(result['records']), 1)
        self.assertEqual(result['records'][0]['sheet'], 'Sheet2')
        self.assertEqual(sum(x['code'] == 'derived_helper_sheet_ignored' for x in result['warnings']), 2)

    def test_explicit_month_year_survives_filename_without_year(self):
        r = record(); r['values'][1] = '2025年1月'
        result = read_analysis(workbook([('Sheet2', {4: HEADERS, 5: r['values'][:67]}, [])]), 'old.xlsx')
        self.assertEqual(result['records'][0]['year'], 2025)
        self.assertEqual(result['year'], 2025)

    def test_mixed_explicit_month_years_and_filename_conflict_rejected(self):
        a, b = record('测试甲'), record('测试乙'); a['values'][1] = '2025年1月'; b['values'][1] = '2026年1月'
        data = workbook([('Sheet2', {4: HEADERS, 5: a['values'][:67], 6: b['values'][:67]}, [])])
        with self.assertRaisesRegex(ValueError, '年度'):
            read_analysis(data, 'old.xlsx')
        with self.assertRaisesRegex(ValueError, '年不一致'):
            read_analysis(data, '2026_old.xlsx')


class MergeTests(unittest.TestCase):
    def test_fill_zero_keep_old_and_report_two_new_sources(self):
        old = record(); old['values'][20] = None; old['values'][22] = 7
        a = record(source='a.xlsx'); a['values'][20] = 0; a['values'][22] = 8
        b = record(source='b.xlsx'); b['values'][20] = 4; b['values'][22] = 9
        result = merge_records([old], [a, b])
        self.assertEqual(result['records'][0]['values'][20], 0)
        self.assertEqual(result['records'][0]['values'][22], 7)
        conflict = [x for x in result['issues'] if x['code'] == 'new_source_conflict']
        self.assertTrue(conflict)
        self.assertEqual(conflict[0]['first_source']['source'], 'a.xlsx')
        self.assertEqual(conflict[0]['second_source']['source'], 'b.xlsx')
        self.assertIsNone(old['values'][20])  # No mutation of caller's data.

    def test_prefer_new_reports_before_after(self):
        old = record(); new = record(source='new.xlsx'); new['values'][22] = 9
        result = merge_records([old], [new], 'prefer_new')
        self.assertEqual(result['records'][0]['values'][22], 9)
        self.assertEqual(result['changes'][0]['before'], 0)
        self.assertEqual(result['changes'][0]['after'], 9)

    def test_missing_settlement_can_match_unique_dispatch_identity(self):
        old, new = record(), record(source='new.xlsx')
        old['values'][2] = None; old['values'][3] = '测试唯一站'
        new['values'][2] = '测试唯一站'; new['values'][3] = None
        result = merge_records([old], [new])
        self.assertEqual(len(result['records']), 1)
        self.assertEqual(result['records'][0]['values'][2], '测试唯一站')

    def test_same_name_two_phases_need_date_unique_match(self):
        old = [record(date='2020-01-01'), record(date='2021-01-01')]
        new = record(source='new.xlsx', date='2021.1.1'); new['values'][20] = 4
        result = merge_records(old, [new], 'prefer_new')
        self.assertEqual(len(result['records']), 2)
        self.assertEqual(result['records'][0]['values'][20], 0)
        self.assertEqual(result['records'][1]['values'][20], 4)
        ambiguous = record(source='ambiguous.xlsx')
        self.assertEqual(merge_records(old, [ambiguous])['stats']['quarantined_records'], 1)

    def test_capacity_change_does_not_create_new_station(self):
        old = record(); new = record(source='new.xlsx'); new['values'][6] = 20
        result = merge_records([old], [new])
        self.assertEqual(len(result['records']), 1)
        self.assertEqual(result['records'][0]['values'][6], 10)
        self.assertTrue(any(x.get('field_index') == 6 for x in result['issues']))

    def test_financial_identity_collision_quarantined(self):
        old = record('测试甲'); new = record('测试乙', source='new.xlsx')
        for j in [20, 21, 22, 32, 33, 34, 59, 60, 61]:
            old['values'][j] = new['values'][j] = j + 1
        result = merge_records([old], [new])
        self.assertEqual(len(result['records']), 1)
        self.assertEqual(result['issues'][0]['code'], 'identity_financial_collision')

    def test_mixed_years_rejected(self):
        old = record(); new = record(source='new.xlsx'); new['year'] = 2027
        with self.assertRaisesRegex(ValueError, '跨年'):
            merge_records([old], [new])


class ReconciliationTests(unittest.TestCase):
    def test_complete_total_components_allow_price_correction(self):
        r = record(); v = r['values']; v[20] = 10; v[29] = 2; v[56] = -1
        v[8] = 3; v[32] = 4; v[59] = 18
        v[22] = 50; v[34] = 40; v[55] = -3; v[46] = 2; v[61] = 89; v[60] = 99
        result = reconcile_records([r]); out = result['records'][0]['values']
        self.assertEqual(out[59], 18); self.assertEqual(out[61], 89)
        self.assertAlmostEqual(out[60], 89/18)
        self.assertEqual(result['stats']['price_corrected'], 1)

    def test_missing_fee_not_zero_even_if_known_sum_matches(self):
        r = record(); v = r['values']; v[20] = v[59] = 10; v[22] = v[61] = 100; v[60] = 99; v[16] = None
        result = reconcile_records([r])
        self.assertEqual(result['records'][0]['values'][60], 99)
        self.assertFalse(result['records'][0]['reconciliation']['fee_complete'])
        self.assertIn('fee_components_unavailable', {x['code'] for x in result['issues']})

    def test_zero_quantity_adjustment_and_price_undefined(self):
        r = record(); v = r['values']; v[31] = v[61] = -5; v[59] = v[60] = 0
        result = reconcile_records([r])
        self.assertEqual(result['records'][0]['values'][61], -5)
        self.assertEqual(result['records'][0]['values'][60], 0)
        self.assertIn('zero_volume_adjustment', {x['code'] for x in result['issues']})

    def test_positive_assessment_fee_not_flipped_mechanism_excluded(self):
        r = record(); v = r['values']; v[20] = v[59] = 10; v[22] = 100; v[55] = 7; v[61] = 107; v[63] = 999; v[65] = 888; v[60] = None
        result = reconcile_records([r])
        self.assertAlmostEqual(result['records'][0]['values'][60], 10.7)
        self.assertEqual(result['records'][0]['reconciliation']['known_fee_sum'], 107)

    def test_inconsistent_totals_do_not_rewrite_price_or_fees(self):
        r = record(); v = r['values']; v[20] = v[59] = 10; v[22] = 100; v[61] = 50; v[60] = 10
        result = reconcile_records([r])
        self.assertEqual(result['records'][0]['values'][60:62], [10, 50])
        self.assertFalse(result['changes'])


if __name__ == '__main__':
    unittest.main()
