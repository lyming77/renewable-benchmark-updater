"""Read-only XLSX ingestion and conservative, provenance-preserving merging.

This module does not write workbooks, evaluate uploaded formulas, or infer
company-specific corrections from the historical example data.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from copy import deepcopy
from datetime import date, datetime, timedelta
from hashlib import sha256
from io import BytesIO
import math
import posixpath
import re
import unicodedata
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile

HEADERS = [
    '来源位置', '月份', '结算单元名称', '调度名称', '所在地区', '场站类型', '装机容量（MW）', '投产时间',
    '省间日前结算电量', '省间日前结算价格', '省间日前结算电费', '省间中长期结算电量', '省间中长期结算价格', '省间中长期结算电费',
    '省间应急外送结算电量', '省间应急外送结算价格', '省间应急外送结算电费', '非市场结算电量', '非市场结算价格', '非市场电能量电费',
    '实际上网电量', '实时结算均价', '实时电能量电费', '日前出清电量', '日前结算均价', '日前电能量电费',
    '中长期合约电量', '中长期结算均价', '中长期合约电费', '月末调整电量', '月末调整结算均价', '月末调整电费',
    '调试期结算电量', '调试期结算电价', '调试期结算电费', '运行成本补偿费用', '运行成本分摊费用', '供热考核费用', '供热返还费用',
    '市场超额收益回收', '调频辅助服务补偿费用', '调频辅助服务分摊费用', '爬坡辅助服务补偿费用', '爬坡辅助服务考核费用', '爬坡辅助服务分摊费用',
    '绿电结算电量', '绿电环境价值（绿证部分费用）', '绿电环境价值偏差补偿费用', '发电侧中长期偏差收益回收', '优发超出优购曲线匹配偏差费用',
    '可用容量', '容量补偿费用', '煤电容量电费', '煤电容量考核电费', '抽水蓄能容量电费', '两个细则费用', '跨月追退补结算电量',
    '跨月追退补结算电费', '清算费用', '总结算电量', '总结算电价', '总结算电费', '机制电量', '机制电费', '国补电量', '国补电费', '集团',
]
EXTRA_HEADERS = {67: '利用小时数', 68: '存量增量', 69: '退补调整分配费用', 70: '数据核对说明'}
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
IDENTITY_FIELDS = {1, 2, 3, 4, 5, 7, 66, 68}
NUMERIC_FIELDS = {6, *range(8, 66), 67, 69}
PLACEHOLDERS = {'-', '--', '/', '\\', '—', '——', '########'}
COMPANY_ALIASES = {'国电山东新能源': '国电电力'}


def normalize_text(value):
    if value is None:
        return ''
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(value))).casefold()


def _header(value):
    text = re.sub(r'_x([0-9a-fA-F]{4})_', lambda m: chr(int(m[1], 16)), str(value or ''))
    text = normalize_text(text)
    return re.sub(r'\([^)]*\)|[：:，,。]', '', text)


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def is_blank(value):
    return value is None or isinstance(value, str) and not value.strip()


def _number(value):
    if is_number(value):
        return value
    if isinstance(value, str):
        raw = unicodedata.normalize('NFKC', value).strip().replace(',', '')
        if re.fullmatch(r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?', raw):
            number = float(raw)
            return number if math.isfinite(number) else value
    return value


def _col(n):
    result = ''
    while n:
        n, d = divmod(n - 1, 26)
        result = chr(65 + d) + result
    return result


def _colnum(ref):
    n = 0
    for letter in re.match(r'[A-Za-z]+', ref).group():
        n = n * 26 + ord(letter.upper()) - 64
    return n


def _issue(code, message, *, record=None, field=None, severity='warning', **details):
    result = {'code': code, 'severity': severity, 'message': message, **details}
    if record is not None:
        result.update({k: record.get(k) for k in ('source', 'sheet', 'source_row', 'company', 'month', 'year')})
        result['station'] = record['values'][2] or record['values'][3]
    if field is not None:
        result['field_index'] = field
        result['field'] = HEADERS[field] if field < 67 else EXTRA_HEADERS.get(field, str(field))
        if record is not None:
            result['cell'] = record.get('source_cells', {}).get(str(field))
    return result


def _date(value):
    if isinstance(value, (datetime, date)):
        return value.strftime('%Y-%m-%d')
    if not isinstance(value, str):
        return None
    match = re.fullmatch(r'\s*(20\d{2}|19\d{2})[-./年](\d{1,2})[-./月](\d{1,2})(?:日|(?:[ T].*)?)?\s*', value)
    if match:
        try:
            return date(*(int(x) for x in match.groups())).isoformat()
        except ValueError:
            return None
    return None


def _month(value):
    if is_number(value) and float(value).is_integer() and 1 <= value <= 12:
        return int(value)
    parsed_date = _date(value)
    if parsed_date:
        return int(parsed_date[5:7])
    match = re.fullmatch(r'\s*(?:20\d{2}[-/.年])?(\d{1,2})(?:月(?:份)?)?\s*', str(value or ''))
    return int(match[1]) if match and 1 <= int(match[1]) <= 12 else None


def _year_from_text(value):
    match = re.search(r'(?<!\d)(20\d{2})(?!\d)', str(value or ''))
    return int(match[1]) if match else None


def _resolve(owner, target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))


def _xlsx(data, filename):
    if filename.lower().endswith(('.xls', '.csv', '.tsv')):
        raise ValueError('请先将文件另存为 .xlsx 或 .xlsm；不支持旧二进制XLS和文本文件。')
    try:
        archive = ZipFile(BytesIO(data))
    except BadZipFile as exc:
        raise ValueError('文件不是有效的xlsx/xlsm工作簿，或文件已损坏/加密。') from exc
    with archive:
        info = archive.infolist()
        if sum(x.file_size for x in info) > 512 * 1024 * 1024 or any(x.file_size > 128 * 1024 * 1024 for x in info):
            raise ValueError('工作簿解压体积过大，请移除空白格式区域或拆分后重试。')
        names = set(archive.namelist())
        if 'xl/workbook.xml' not in names:
            raise ValueError('工作簿缺少workbook.xml；请用Excel/WPS另存为xlsx。')
        shared = []
        if 'xl/sharedStrings.xml' in names:
            shared = [''.join(t.text or '' for t in node.iter(S + 't')) for node in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        relationships = {node.get('Id'): node.get('Target') for node in ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))}
        epoch = datetime(1904, 1, 1) if workbook.find(S + 'workbookPr') is not None and workbook.find(S + 'workbookPr').get('date1904') in ('1', 'true') else datetime(1899, 12, 30)
        date_styles = set()
        if 'xl/styles.xml' in names:
            styles = ET.fromstring(archive.read('xl/styles.xml'))
            formats = {int(x.get('numFmtId')): x.get('formatCode', '') for x in styles.findall(S + 'numFmts/' + S + 'numFmt')}
            xfs = styles.find(S + 'cellXfs')
            for i, xf in enumerate(xfs if xfs is not None else []):
                fmt_id = int(xf.get('numFmtId', '0'))
                fmt = re.sub(r'"[^"]*"|\[[^]]*\]|\\.', '', formats.get(fmt_id, '')).lower()
                if fmt_id in set(range(14, 23)) | set(range(27, 37)) | set(range(50, 59)) or re.search(r'y+', fmt) and re.search(r'[md]', fmt):
                    date_styles.add(i)
        sheets = []
        for meta in workbook.find(S + 'sheets'):
            target = relationships.get(meta.get(R + 'id'))
            if not target:
                continue
            path = _resolve('xl/workbook.xml', target)
            root = ET.fromstring(archive.read(path))
            rows, formulas, errors = defaultdict(dict), {}, set()
            for cell in root.iter(S + 'c'):
                ref = cell.get('r', '')
                if not re.fullmatch(r'[A-Z]+\d+', ref):
                    continue
                c, r = _colnum(ref) - 1, int(re.search(r'\d+', ref).group())
                typ, node = cell.get('t'), cell.find(S + 'v')
                formula = cell.find(S + 'f')
                if formula is not None:
                    formulas[(r, c)] = '=' + (formula.text or '')
                value = node.text if node is not None else None
                if typ == 'inlineStr':
                    value = ''.join(t.text or '' for t in cell.iter(S + 't'))
                elif typ == 's' and value is not None:
                    value = shared[int(value)]
                elif typ == 'b' and value is not None:
                    value = value == '1'
                elif typ not in ('str', 'e', 'd') and value is not None:
                    value = _number(value)
                    if is_number(value) and int(cell.get('s', '0')) in date_styles:
                        try:
                            value = (epoch + timedelta(days=value)).isoformat()
                        except (ValueError, OverflowError):
                            pass
                if typ == 'e':
                    errors.add((r, c))
                if value is not None:
                    rows[r][c] = value
            merges = []
            for node in root.findall(S + 'mergeCells/' + S + 'mergeCell'):
                ends = node.get('ref', '').split(':')
                if len(ends) == 2:
                    c1, c2 = _colnum(ends[0]) - 1, _colnum(ends[1]) - 1
                    r1, r2 = (int(re.search(r'\d+', e).group()) for e in ends)
                    merges.append((r1, c1, r2, c2))
            sheets.append({'name': meta.get('name'), 'rows': dict(rows), 'formulas': formulas, 'errors': errors, 'merges': merges})
        return sheets


def _aliases(headers):
    lookup = {_header(h): i for i, h in enumerate(HEADERS)}
    lookup.update({_header(h): i for i, h in enumerate(headers[:67])})
    lookup.update({_header(h): i for i, h in EXTRA_HEADERS.items()})
    names = {
        0: ['源位置', '来源'], 1: ['月', '统计月份'], 2: ['结算单位名称', '结算单元', '结算名称'],
        3: ['调度名', '场站名称', '电站名称'], 4: ['地区', '区域'], 5: ['类型', '电源类型'],
        6: ['装机容量', '场站容量', '容量'], 7: ['投产日期', '并网时间', '并网日期'], 20: ['实际上网结算电量', '实时结算电量', '实时电量'],
        21: ['实时结算电价', '实时电价'], 22: ['实时结算电费', '实时电费'],
        29: ['月末调整结算电量'], 31: ['月末调整结算电费'], 46: ['绿电环境价值', '绿证费用', '绿电环境价值费用'],
        55: ['两个细则考核费用', '两个细则电费'], 56: ['跨月追退补电量', '跨月结算电量'], 57: ['跨月追退补电费', '跨月结算电费'],
        59: ['总结算量', '总电量', '合计结算电量'], 60: ['总结算均价', '结算平均电价'], 61: ['总结算费', '总电费', '合计结算电费'],
        66: ['公司', '公司名称', '集团名称', '所属集团'], 69: ['退补调整分配电费'], 70: ['备注', '核对说明'],
    }
    for i, vals in names.items():
        lookup.update({_header(x): i for x in vals})
    return lookup


def _header_values(sheet, row, width):
    values = [sheet['rows'].get(row, {}).get(c) for c in range(width)]
    for r1, c1, r2, c2 in sheet['merges']:
        if r1 <= row <= r2 and r1 <= row and r1 <= 30:
            anchor = sheet['rows'].get(r1, {}).get(c1)
            for c in range(c1, min(c2 + 1, width)):
                if values[c] is None:
                    values[c] = anchor
    return values


def _mapping(sheet, end, lookup):
    width = min(512, 1 + max((max(row, default=0) for r, row in sheet['rows'].items() if max(1, end - 1) <= r <= end), default=0))
    upper = _header_values(sheet, max(1, end - 1), width)
    lower = _header_values(sheet, end, width)
    mapping, raw_headers, duplicate = {}, {}, []
    for c in range(width):
        for candidate in (lower[c], upper[c], str(upper[c] or '') + str(lower[c] or '')):
            h = _header(candidate)
            if h in lookup:
                j = lookup[h]
                if j in mapping.values():
                    duplicate.append((c, j))
                else:
                    mapping[c] = j
                    raw_headers[c] = str(candidate)
                break
    return mapping, raw_headers, duplicate


def _find_header(sheet, headers):
    lookup = _aliases(headers)
    best = None
    for row in range(1, 31):
        mapping, raw, dup = _mapping(sheet, row, lookup)
        fields = set(mapping.values())
        score = len(fields) + 5 * (1 in fields) + 5 * bool(fields & {2, 3})
        if (1 in fields and fields & {2, 3} and len(fields) >= 8) and (best is None or score > best[0]):
            best = (score, row, mapping, raw, dup)
    return best


def _capacity_unit(text):
    h = normalize_text(text)
    if re.search(r'万千瓦时|万kwh|mwh|兆瓦时', h):
        return None, 'energy_unit'
    if '万千瓦' in h or '万kw' in h:
        return 10, '万千瓦'
    if '兆瓦' in h or re.search(r'mw(?!h)', h):
        return 1, 'MW'
    if '千瓦' in h or re.search(r'kw(?!h)', h):
        return .001, 'kW'
    return 1, 'unconfirmed'


def _read_sheet(sheet, filename, headers, year, *, collection):
    found = _find_header(sheet, headers)
    if not found:
        return None
    _, end, mapping, raw_headers, duplicate = found
    warnings, records = [], []
    for c, j in duplicate:
        warnings.append(_issue('duplicate_header', '两个源列映射到同一字段，未使用重复列；请确认表头。', source=filename, sheet=sheet['name'], source_row=end, field=j, duplicate_column=_col(c + 1)))
    anchor = {}
    for r1, c1, r2, c2 in sheet['merges']:
        if c1 == c2 and r2 > end:
            for r in range(max(end + 1, r1 + 1), r2 + 1):
                anchor[(r, c1)] = (r1, c1)
    current_month = None
    current_year = year
    lookup = _aliases(headers)
    last = max(sheet['rows'], default=end)
    for rn in range(end + 1, last + 1):
        rawrow = sheet['rows'].get(rn, {})
        if not any(not is_blank(v) for v in rawrow.values()):
            current_month = None
            current_year = year
            continue
        # Repeated headers are evidence for a mid-sheet column-layout change.
        newmap, newheads, _ = _mapping(sheet, rn, lookup)
        own_header_fields = {lookup[_header(v)] for v in rawrow.values() if _header(v) in lookup}
        looks_like_header = len(own_header_fields) >= 8 or len(own_header_fields) >= 4 and 1 in own_header_fields and bool(own_header_fields & {2, 3})
        if looks_like_header and len(newmap) >= 8 and {1}.issubset(newmap.values()) and set(newmap.values()) & {2, 3}:
            if newmap != mapping:
                warnings.append(_issue('mid_sheet_layout_change', '检测到中途表头/列结构变化，后续按新标题映射。', source=filename, sheet=sheet['name'], source_row=rn))
            mapping, raw_headers = newmap, newheads
            current_month = None
            current_year = year
            continue
        values = [None] * 71
        capacity_cell_unit = None
        cells, states, originals, merged_numeric = {}, {str(j): 'absent' for j in range(71)}, {}, []
        for c, j in mapping.items():
            value = rawrow.get(c)
            source_r = rn
            if value is None and (rn, c) in anchor:
                ar, ac = anchor[(rn, c)]
                if j in IDENTITY_FIELDS:
                    value, source_r = sheet['rows'].get(ar, {}).get(ac), ar
                elif j in NUMERIC_FIELDS:
                    merged_numeric.append(j)
            formula = sheet['formulas'].get((rn, c))
            state = 'blank' if is_blank(value) else 'present'
            if formula and value is None:
                state = 'formula_missing_cache'
            if (rn, c) in sheet['errors'] or j in NUMERIC_FIELDS and isinstance(value, str) and (value.lstrip().startswith('#') or value.strip() in PLACEHOLDERS):
                originals[str(j)] = value
                value, state = None, 'source_error'
            if j == 6 and isinstance(value, str):
                unit_match = re.fullmatch(r'\s*([+-]?[\d,.]+)\s*(万千瓦时|万千瓦|万kwh|万kw|mwh|mw|kw|千瓦|兆瓦)\s*', unicodedata.normalize('NFKC', value), re.I)
                if unit_match:
                    value, capacity_cell_unit = _number(unit_match[1]), unit_match[2]
            if j in NUMERIC_FIELDS:
                value = _number(value)
                if not is_blank(value) and not is_number(value):
                    originals[str(j)] = value
                    value, state = None, 'invalid_numeric'
            if j == 7 and value is not None:
                if normalize_text(value) in ('存量', '增量'):
                    values[68] = value
                    states['68'] = 'present'
                    originals['7'] = value
                    value, state = None, 'mislabelled_stock_increment'
                else:
                    parsed = _date(value)
                    if parsed:
                        # Preserve the old analysis cell's exact observation;
                        # compare date semantics during merge instead.
                        if collection:
                            value = parsed
                    elif not is_blank(value):
                        originals['7'] = value
                        if collection and re.fullmatch(r'[A-Za-z]', str(value).strip()):
                            value, state = None, 'invalid_date'
                        else:
                            # A year-only entry or multi-phase date description
                            # is a real source observation, not a missing cell.
                            state = 'unstructured_date'
            values[j], states[str(j)] = value, state
            cells[str(j)] = f'{_col(c + 1)}{source_r}'
        station = values[2] or values[3]
        if is_blank(station) or re.fullmatch(r'(?:合计|总计|小计|说明|备注)(?:[：:].*)?', str(station).strip()):
            continue
        explicit_year = _year_from_text(values[1])
        if explicit_year is not None and year is not None and explicit_year != year:
            raise ValueError(f'{filename} / {sheet["name"]}!第{rn}行月份属于{explicit_year}年，与所选{year}年不一致；请分年度处理。')
        explicit_month = _month(values[1])
        if explicit_month:
            current_month = explicit_month
            current_year = explicit_year or year or current_year
        elif not is_blank(values[1]):
            current_month = None
        co = str(values[66]).strip() if not is_blank(values[66]) else sheet['name'] if collection and not re.fullmatch(r'(?:sheet\d*|汇总|明细|分析表|收集表)', sheet['name'], re.I) else None
        if co:
            co = COMPANY_ALIASES.get(co, co)
        rec = {'values': values, 'source': filename, 'sheet': sheet['name'], 'source_row': rn,
               'source_cells': cells, 'company': co, 'month': current_month, 'year': current_year,
               'present_fields': sorted(set(mapping.values())), 'field_states': states,
               'original_errors': originals, 'merged_numeric_fields': merged_numeric,
               'provenance': [{'source': filename, 'sheet': sheet['name'], 'row': rn}],
               'record_id': sha256(f'{filename}|{sheet["name"]}|{rn}'.encode()).hexdigest()[:24]}
        if collection:
            capacity_column = next((c for c, j in mapping.items() if j == 6), None)
            for r1, c1, r2, c2 in sheet['merges']:
                if c1 == c2 == capacity_column and end < r1 < r2 and r1 <= rn <= r2:
                    shared_value = sheet['rows'].get(r1, {}).get(c1)
                    rec.setdefault('shared_numeric_observations', []).append({'field_index': 6, 'raw_value': shared_value,
                        'range': f'{_col(c1+1)}{r1}:{_col(c2+1)}{r2}', 'anchor': f'{_col(c1+1)}{r1}'})
                    values[6], states['6'] = None, 'ambiguous_merged_capacity'
                    warnings.append(_issue('merged_capacity_scope_ambiguous', '容量跨多条场站行合并，无法确认各分期容量；单站容量留空，原合并值和范围另存，不能覆盖旧表已确认单期容量。',
                                           record=rec, field=6, original=shared_value, source_range=rec['shared_numeric_observations'][-1]['range']))
                    break
        if not co or not current_month:
            warnings.append(_issue('missing_key', '缺少可确认的集团或月份，记录暂不并入；请补充来源表。', record=rec, original_company=values[66], original_month=values[1]))
            continue
        if collection or is_blank(values[0]):
            values[0] = f'{filename} / {sheet["name"]}!{rn}'
        values[1], values[66] = f'{current_month}月份', co
        if values[6] is not None:
            c = next((c for c, j in mapping.items() if j == 6), None)
            unittext = capacity_cell_unit or raw_headers.get(c, '')
            factor, unit = _capacity_unit(unittext)
            if factor is None:
                originals['6'] = values[6]
                values[6], states['6'] = None, 'ambiguous_unit'
                warnings.append(_issue('capacity_energy_unit', '装机容量标题使用电量单位，无法可靠换算为MW；保留原数并待确认。', record=rec, field=6, original=originals['6'], original_unit=unittext))
            else:
                if factor != 1:
                    old = values[6]
                    values[6] *= factor
                    rec.setdefault('normalizations', []).append({'field_index': 6, 'before': old, 'after': values[6], 'reason': f'{unit}换算为MW', 'factor': factor})
                if unit == 'unconfirmed':
                    rec['capacity_unit_unconfirmed'] = True
        for j, state in states.items():
            if state in ('source_error', 'invalid_numeric', 'formula_missing_cache', 'invalid_date'):
                warnings.append(_issue(state, '原单元格错误、占位或缺少公式缓存；整理值留空，未按0处理。', record=rec, field=int(j), original=originals.get(j, sheet['formulas'].get((rn, next((c for c, k in mapping.items() if k == int(j)), -1))))))
            elif state == 'unstructured_date':
                warnings.append(_issue(state, '投产信息是年份、分期描述或非标准日期，原文保留；未猜测完整日期，不能据此自动拆分同名分期。', record=rec, field=int(j), original=originals.get(j)))
        if merged_numeric:
            warnings.append(_issue('merged_numeric_not_broadcast', '合并的数值指标仅保留在原锚点，未复制到子行；容量/合并费用需独立核实。', record=rec, field_indexes=merged_numeric))
        if all(is_blank(values[j]) for j in range(8, 66)):
            warnings.append(_issue('empty_transaction_record', '有公司/月/场站身份，但全部交易量价费为空；保留该身份行并明确缺报，不当作真实0交易。', record=rec))
        records.append(rec)
    return {'records': records, 'warnings': warnings, 'sheet_name': sheet['name'], 'header_row': end,
            'mapping': {str(c + 1): j for c, j in mapping.items()}, 'source_headers': raw_headers}


def read_analysis(data: bytes, filename: str) -> dict:
    """Read the old canonical analysis sheet without changing its package."""
    sheets = _xlsx(data, filename)
    candidates = [(s, _find_header(s, HEADERS)) for s in sheets]
    candidates = [(s, h) for s, h in candidates if h]
    if not candidates:
        raise ValueError('没有找到含月份、场站名称和交易字段的分析明细表。')
    preferred = next((s for s, h in candidates if s['name'] == 'Sheet2'), None)
    sheet = preferred or max(candidates, key=lambda x: x[1][0])[0]
    found = _find_header(sheet, HEADERS)
    _, end, mapping, raw, _ = found
    headers = list(HEADERS)
    for c, j in mapping.items():
        if j < 67:
            headers[j] = raw[c]
    headers[0] = '来源位置'
    year = _year_from_text(filename)
    result = _read_sheet(sheet, filename, headers, year, collection=False)
    row_years = {r['year'] for r in result['records'] if r.get('year') is not None}
    if len(row_years) > 1:
        raise ValueError('旧分析表的月份单元格包含多个年度，不能静默归一到同一年；请按年度分别处理。')
    year = year or next(iter(row_years), None)
    result.update({'headers': headers, 'filename': filename, 'year': year, 'input_sha256': sha256(data).hexdigest()})
    if year is None:
        result['warnings'].append(_issue('year_unconfirmed', '文件名未明确年份，请在界面确认年度后再合并，避免跨年同月误合并。', source=filename))
    return result


def read_collection(data: bytes, filename: str, headers: list[str], year: int) -> dict:
    """Read company worksheets or a canonical long table by header semantics."""
    if not 2000 <= int(year) <= 2100:
        raise ValueError('年度必须在2000至2100之间。')
    named_year = _year_from_text(filename)
    if named_year is not None and named_year != int(year):
        raise ValueError(f'上传文件名标注{named_year}年，与所选{year}年不一致；请分年度处理。')
    records, warnings, layouts = [], [], []
    for sheet in _xlsx(data, filename):
        if sheet['name'] in ('小时联动源', '期间场站汇总') or any(
            {_header(v) for v in row.values()} >= {_header('记录用途'), _header('小时源变更')}
            for r, row in sheet['rows'].items() if r <= 30
        ):
            warnings.append(_issue('derived_helper_sheet_ignored', '该页为派生辅助数据，不能再次作为业务明细导入。', source=filename, sheet=sheet['name']))
            continue
        parsed = _read_sheet(sheet, filename, headers, int(year), collection=True)
        if parsed:
            records.extend(parsed['records']); warnings.extend(parsed['warnings'])
            layouts.append({k: parsed[k] for k in ('sheet_name', 'header_row', 'mapping')})
    if not records:
        raise ValueError('没有读取到有效公司/月/场站明细；请检查表头、集团和月份。')
    return {'headers': list(headers[:67]), 'records': records, 'warnings': warnings, 'layouts': layouts,
            'filename': filename, 'year': int(year), 'input_sha256': sha256(data).hexdigest()}


def _values(record):
    values = list(record.get('values', record.get('v', [])))
    return values + [None] * max(0, 71 - len(values))


def _prepare(record):
    r = deepcopy(record)
    r['values'] = _values(r)
    r['company'] = r.get('company') or r['values'][66]
    r['month'] = r.get('month') or _month(r['values'][1])
    r.setdefault('source', '')
    r.setdefault('year', None)
    return r


def _equivalent(a, b, field):
    if is_blank(a) and is_blank(b):
        return True
    if field == 7 and _date(a) and _date(b):
        return _date(a) == _date(b)
    if is_number(a) and is_number(b):
        return math.isclose(a, b, rel_tol=1e-12, abs_tol=1e-8)
    return a == b


def _source(record):
    return {'source': record.get('source'), 'sheet': record.get('sheet'), 'row': record.get('source_row'), 'record_id': record.get('record_id')}


def _financial_signature(record):
    values = record['values']
    nonzero = [(j, float(values[j])) for j in range(8, 66) if is_number(values[j]) and values[j] != 0]
    return tuple(nonzero) if len(nonzero) >= 6 else None


def merge_records(old_records, new_records, conflict_policy='keep_old') -> dict:
    """Fill missing observations; retain every disputed value and its origin."""
    if conflict_policy not in ('keep_old', 'prefer_new'):
        raise ValueError('conflict_policy必须为keep_old或prefer_new。')
    records = [_prepare(r) for r in old_records]
    incoming = [_prepare(r) for r in new_records]
    known_years = {r.get('year') for r in records + incoming if r.get('year') is not None}
    if len(known_years) > 1:
        raise ValueError('本次包含多个年度，不能跨年合并同一分析工作簿；请按年度分别上传。')
    changes, issues, quarantined = [], [], []
    new_field_origins = {}
    stats = Counter({'old_records': len(records), 'incoming_records': len(incoming)})

    def base_key(r):
        return (normalize_text(r['company']), r.get('year'), r.get('month'), normalize_text(r['values'][2] or r['values'][3]))

    station_index, dispatch_index = defaultdict(set), defaultdict(set)
    financial_index = defaultdict(set)
    def index_record(index):
        r = records[index]
        group = (normalize_text(r['company']), r.get('year'), r.get('month'))
        station, dispatch = normalize_text(r['values'][2]), normalize_text(r['values'][3])
        if station:station_index[group + (station,)].add(index)
        if dispatch:dispatch_index[group + (dispatch,)].add(index)
        sig = _financial_signature(r)
        if sig is not None:financial_index[group + (sig,)].add(index)
    for i in range(len(records)):
        index_record(i)

    # A same-file repeated station with distinct valid dates is positive phase
    # evidence. A date changing between two single rows is only a conflict.
    phases = defaultdict(set)
    groups = defaultdict(list)
    for kind, rr in [('old', records), ('new', incoming)]:
        for r in rr:
            groups[(kind, base_key(r), r.get('source'), r.get('sheet'))].append(r)
    for (_, k, _, _), rr in groups.items():
        ds = [_date(r['values'][7]) for r in rr]
        if len(rr) > 1 and all(ds) and len(ds) == len(set(ds)):
            phases[k].update(ds)

    for new in incoming:
        if not new['company'] or not new['month'] or not (new['values'][2] or new['values'][3]):
            issues.append(_issue('missing_merge_key', '缺少公司/月/场站唯一标识，暂不合并。', record=new, severity='error'))
            quarantined.append(new); continue
        if new.get('year') is None:
            issues.append(_issue('year_unconfirmed', '新记录年份未确认，暂不合并。', record=new, severity='error'))
            quarantined.append(new); continue
        candidates = []
        group = (normalize_text(new['company']), new.get('year'), new['month'])
        ns_key, nd_key = normalize_text(new['values'][2]), normalize_text(new['values'][3])
        pool = (station_index[group + (ns_key,)] | dispatch_index[group + (nd_key,)] |
                dispatch_index[group + (ns_key,)] | station_index[group + (nd_key,)])
        for i in sorted(pool):
            old = records[i]
            os, ns = normalize_text(old['values'][2]), normalize_text(new['values'][2])
            od, nd = normalize_text(old['values'][3]), normalize_text(new['values'][3])
            if os and ns and os == ns or (not os or not ns) and bool(({os, od} - {''}) & ({ns, nd} - {''})):
                candidates.append(i)
        if len(candidates) > 1 or phases.get(base_key(new)):
            nd = _date(new['values'][7])
            dated = [i for i in candidates if nd and _date(records[i]['values'][7]) == nd]
            if dated:
                candidates = dated
            elif nd and phases.get(base_key(new)) and candidates:
                candidates = []  # Independently declared separate phase.
        if len(candidates) > 1:
            issues.append(_issue('ambiguous_station_match', '同公司同月同名存在多条，投产分期信息不足以唯一匹配；暂不合并。', record=new, candidates=[_source(records[i]) for i in candidates], severity='error'))
            quarantined.append(new); continue
        if not candidates:
            sig = _financial_signature(new)
            suspicious = [records[i] for i in financial_index[group + (sig,)] if sig is not None and _financial_signature(records[i]) == sig]
            if suspicious:
                issues.append(_issue('identity_financial_collision', '名称/身份不同但多项非零交易值完全相同，可能是版本身份冲突，暂不新增以免重复交易。', record=new, candidates=[_source(r) for r in suspicious], severity='error'))
                quarantined.append(new); continue
            if any(r['company'] == new['company'] and r['month'] == new['month'] and r.get('year') is None and normalize_text(r['values'][2] or r['values'][3]) == normalize_text(new['values'][2] or new['values'][3]) for r in records):
                issues.append(_issue('old_year_unconfirmed', '旧表同站同月记录未确认年份，暂不新增或跨年合并。', record=new, severity='error'))
                quarantined.append(new); continue
            records.append(new)
            index = len(records) - 1
            index_record(index)
            for j, value in enumerate(new['values']):
                if not is_blank(value):new_field_origins[(index, j)] = (value, _source(new))
            changes.append({'kind': 'added_record', 'after': deepcopy(new), 'new_source': _source(new)})
            stats['added_records'] += 1
            continue
        index = candidates[0]; old = records[index]; matched_change = False
        stats['matched_records'] += 1
        for j in range(2, max(len(old['values']), len(new['values']))):
            if j in (66,):
                continue
            a, b = old['values'][j], new['values'][j]
            if is_blank(b):
                continue
            previous_new = new_field_origins.get((index, j))
            if previous_new and not _equivalent(previous_new[0], b, j):
                issues.append(_issue('new_source_conflict', '多个新文件对同一记录同一字段报告不同值；已同时列明来源。', record=new, field=j,
                                     first_value=previous_new[0], second_value=b, first_source=previous_new[1], second_source=_source(new), policy=conflict_policy))
            else:
                new_field_origins.setdefault((index, j), (b, _source(new)))
            if _equivalent(a, b, j):
                continue
            if not is_blank(a):
                issues.append(_issue('value_conflict', '新旧非空值不一致；按所选策略处理，未隐去原值。', record=new, field=j,
                                     old_value=a, new_value=b, old_source=old.get('field_provenance', {}).get(str(j), _source(old)), new_source=_source(new), policy=conflict_policy))
                stats['conflicts'] += 1
                if j == 6 and is_number(a) and is_number(b) and b and (math.isclose(a / b, 10, rel_tol=1e-6) or math.isclose(b / a, 10, rel_tol=1e-6)):
                    issues.append(_issue('suspected_capacity_unit_conflict', '容量差10倍，疑似MW/万千瓦混用，不能仅凭倍数自动纠正。', record=new, field=6, old_value=a, new_value=b))
                if conflict_policy == 'keep_old':
                    continue
            old['values'][j] = b
            kind = 'filled_empty' if is_blank(a) else 'overwritten_conflict'
            changes.append({'kind': kind, 'record_id': old.get('record_id'), 'company': old['company'], 'month': old['month'], 'year': old.get('year'),
                            'station': old['values'][2] or old['values'][3], 'field_index': j, 'field': HEADERS[j] if j < 67 else EXTRA_HEADERS.get(j, str(j)),
                            'before': a, 'after': b, 'old_source': _source(old), 'new_source': _source(new),
                            'source_cell': new.get('source_cells', {}).get(str(j))})
            stats[kind] += 1; matched_change = True
            old.setdefault('field_states', {})[str(j)] = new.get('field_states', {}).get(str(j), 'present')
            old.setdefault('source_cells', {})[str(j)] = new.get('source_cells', {}).get(str(j))
            old.setdefault('field_provenance', {})[str(j)] = _source(new)
        old['present_fields'] = sorted(set(old.get('present_fields', [])) | set(new.get('present_fields', [])))
        old.setdefault('provenance', [_source(old)])
        if _source(new) not in old['provenance']:
            old['provenance'].append(_source(new))
        if matched_change:
            stats['changed_records'] += 1
            index_record(index)
    stats['output_records'] = len(records); stats['quarantined_records'] = len(quarantined)
    return {'records': records, 'changes': changes, 'issues': issues, 'stats': dict(stats), 'quarantined_records': quarantined}
