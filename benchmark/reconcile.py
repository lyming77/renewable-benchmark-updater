"""Evidence-constrained reconciliation; never rebuild fees from Q*P alone."""
from __future__ import annotations

from collections import Counter
from copy import deepcopy
from decimal import Decimal, InvalidOperation
import calendar
import math

from .ingest import HEADERS, EXTRA_HEADERS, is_blank, is_number

BASE_QUANTITY_FIELDS = (20, 29, 56)
INDEPENDENT_QUANTITY_FIELDS = (8, 11, 14, 17, 32)
QUANTITY_FIELDS = (8, 11, 14, 17, 20, 29, 32, 56)
FEE_FIELDS = (10, 13, 16, 19, 22, 25, 28, 31, 34, *range(35, 45), 46, 47, 48, 49, 51, 52, 53, 54, 55, 57, 58)
TRIPLETS = ((8, 9, 10), (11, 12, 13), (14, 15, 16), (17, 18, 19), (20, 21, 22),
            (23, 24, 25), (26, 27, 28), (29, 30, 31), (32, 33, 34))


def _sum(values, fields):
    return float(sum((Decimal(str(values[i])) for i in fields if i < len(values) and is_number(values[i])), Decimal(0)))


def _step(value):
    if not is_number(value) or value == 0:
        return 0.
    try:
        return float(Decimal(1).scaleb(min(0, Decimal(str(value)).normalize().as_tuple().exponent)))
    except InvalidOperation:
        return 0.


def _rounding_bound(values, fields, reported):
    return .5 * (sum(_step(values[i]) for i in fields if i < len(values)) + _step(reported)) + 1e-6


def _field_name(index):
    return HEADERS[index] if index < 67 else EXTRA_HEADERS.get(index, str(index))


def _state(record, index):
    states = record.get('field_states', {})
    return states.get(str(index), states.get(index, 'unknown'))


def _missing(record, fields):
    """A genuine omitted source block is structural; an empty declared cell is not zero."""
    values = record['values']
    optional = set(record.get('not_applicable_fields', []))
    structural, unresolved = [], []
    for index in fields:
        if index < len(values) and is_number(values[index]):
            continue
        if index in optional or str(index) in optional or _state(record, index) in ('not_applicable', 'structural_absent'):
            structural.append(index)
        else:
            unresolved.append(index)
    return structural, unresolved


def reconcile_records(records) -> dict:
    """Keep all business observations; change BI only after independent totals checks.

    Plain blanks remain unresolved unless explicitly classified not applicable.
    Missing/error quantities and fees are never filled with zero. All assessment
    and allocation signs are retained. Mechanism/subsidy amounts are outside BJ.
    """
    result = deepcopy(records)
    issues, changes = [], []
    counts = Counter()
    for record in result:
        values = list(record.get('values', record.get('v', [])))
        values += [None] * max(0, 71 - len(values))
        record['values'] = values
        local_issues = []

        def report(code, message, *, severity='warning', fields=(), **details):
            issue = {'code': code, 'severity': severity, 'message': message,
                     'record_id': record.get('record_id'), 'company': record.get('company', values[66]),
                     'month': record.get('month'), 'year': record.get('year'), 'station': values[2] or values[3],
                     'source': record.get('source'), 'sheet': record.get('sheet'), 'source_row': record.get('source_row'),
                     'field_indexes': list(fields), 'fields': [_field_name(i) for i in fields], **details}
            local_issues.append(issue)
            return issue

        year, month = record.get('year'), record.get('month')
        if is_number(values[6]) and values[6] > 0 and is_number(values[20]) and isinstance(year, int) and isinstance(month, int) and 1 <= month <= 12:
            maximum = values[6] * calendar.monthrange(year, month)[1] * 24
            if values[20] > maximum * 1.01:
                report('capacity_energy_physical_bound', '实际上网电量超过该月容量×日历小时，需核对容量单位、电量单位或场站身份；未自动按10倍换算。', fields=(6, 20), capacity_mw=values[6], quantity_mwh=values[20], monthly_capacity_hours_mwh=maximum)

        # Zero quantity does not imply a missing price: zero-volume adjustment
        # fees remain legitimate observations with an undefined average price.
        for q, p, f in (*TRIPLETS, (59, 60, 61)):
            qty, price, fee = values[q], values[p], values[f]
            if is_number(qty) and qty == 0 and is_number(fee):
                if not is_number(price) or price == 0:
                    record.setdefault('field_states', {})[str(p)] = 'not_applicable'
                    if fee != 0:
                        report('zero_volume_adjustment', '该项电量为0但有已报调整费用，均价无定义；保留费用，不按错误金额处理。', severity='info', fields=(q, p, f), quantity=qty, fee=fee)
                continue
            if all(is_number(x) for x in (qty, price, fee)) and q != 59:
                product = qty * price
                allowance = max(.01, abs(qty) * _step(price) / 2 + abs(price) * _step(qty) / 2 + _step(fee) / 2)
                if abs(fee - product) > allowance + 1e-6:
                    report('component_price_fee_mismatch', '分项电量×电价与已报电费不同，不能据此判断哪一项错误；原量价费全部保留。', fields=(q, p, f), quantity=qty, price=price, fee=fee, product=product, difference=fee - product, tolerance=allowance)

        # A source may explicitly omit entire market blocks. Do not silently
        # promote ordinary blanks into zeros or claim complete decomposition.
        q_structural, q_missing = _missing(record, QUANTITY_FIELDS)
        fee_fields = list(FEE_FIELDS)
        if is_number(values[69]) or 69 in record.get('present_fields', []) or str(69) in record.get('field_states', {}) and _state(record, 69) not in ('absent', 'unknown'):
            fee_fields.append(69)
        f_structural, f_missing = _missing(record, fee_fields)
        qsum, fsum = _sum(values, QUANTITY_FIELDS), _sum(values, fee_fields)
        qtotal, price, ftotal = values[59], values[60], values[61]
        qdiff = qtotal - qsum if is_number(qtotal) else None
        fdiff = ftotal - fsum if is_number(ftotal) else None
        qbound = _rounding_bound(values, QUANTITY_FIELDS, qtotal)
        fbound = _rounding_bound(values, fee_fields, ftotal)
        q_exact = qdiff is not None and abs(qdiff) <= 1e-6
        f_exact = fdiff is not None and abs(fdiff) <= .0100001
        q_complete, f_complete = not q_missing, not f_missing
        if q_missing:
            report('quantity_components_unavailable', '物理电量组成项存在未确认空白或源错误；只列已知和，未将空白当0证明总量。', fields=q_missing, known_quantity_sum=qsum, reported_quantity=qtotal)
        if f_missing:
            report('fee_components_unavailable', '费用组成项存在未确认空白或源错误；只列已知和，未将空白当0证明总费。', fields=f_missing, known_fee_sum=fsum, reported_fee=ftotal)
        if qdiff is None:
            report('summary_quantity_unavailable', '总结算电量缺失，未通过均价或费用倒算补入。', fields=(59,))
        elif not q_exact:
            report('summary_quantity_mismatch' if abs(qdiff) > qbound else 'quantity_within_reported_precision',
                   '总结算电量与已列独立物理量之和存在差额；月末/跨月量可能已含在实时量内，保留总量待核口径。' if abs(qdiff) > qbound else '总量差额在现报数字末位舍入上界内；不据此恢复更细真实电量。',
                   severity='warning' if abs(qdiff) > qbound else 'info', fields=(59, *QUANTITY_FIELDS),
                   reported=qtotal, known_sum=qsum, difference=qdiff, rounding_bound=qbound)
        if fdiff is None:
            report('summary_fee_unavailable', '总结算电费缺失，未用电量×电价补成报送金额。', fields=(61,))
        elif not f_exact:
            report('summary_fee_mismatch' if abs(fdiff) > fbound else 'fee_within_reported_precision',
                   '总结算电费与按原正负号相加的已列分项存在差额；未修改分项或总费。' if abs(fdiff) > fbound else '费用差额在现报数字末位舍入上界内；不据此恢复更细真实费用。',
                   severity='warning' if abs(fdiff) > fbound else 'info', fields=(61, *fee_fields),
                   reported=ftotal, known_sum=fsum, difference=fdiff, rounding_bound=fbound)
        reliable = q_complete and f_complete and q_exact and f_exact and is_number(qtotal) and qtotal > 0 and is_number(ftotal)
        expected_price = ftotal / qtotal if is_number(ftotal) and is_number(qtotal) and qtotal != 0 else None
        corrected = False
        if reliable and expected_price is not None:
            allowance = max(.005, _step(price) / 2) if is_number(price) else 0
            if not is_number(price) or abs(price - expected_price) > allowance + 1e-8:
                change = {'kind': 'verified_summary_price', 'record_id': record.get('record_id'), 'company': record.get('company', values[66]),
                          'month': record.get('month'), 'station': values[2] or values[3], 'field_index': 60, 'field': HEADERS[60],
                          'before': price, 'after': expected_price, 'source': record.get('source'),
                          'reason': '独立物理量组成与原总量一致、完整已列费用和与原总费逐分一致后，按原总费/同口径总量修均价；总量总费未改。'}
                changes.append(change); values[60] = expected_price; corrected = True
                record.setdefault('field_states', {})['60'] = 'computed_verified'
                record.setdefault('computed_fields', {})['60'] = {'formula': 'BJ/BH', 'quantity_fields': list(QUANTITY_FIELDS), 'fee_fields': fee_fields}
        elif expected_price is not None and is_number(price) and abs(price - expected_price) > max(.005, _step(price) / 2) + 1e-8:
            report('summary_price_not_certified', '原均价与总费/总量不同，但总量、总费或组成完整性尚未同时验证，均价暂不修改。', fields=(59, 60, 61), reported_price=price, ratio_of_reported_totals=expected_price)
        record['reconciliation'] = {'known_quantity_sum': qsum, 'known_fee_sum': fsum, 'quantity_difference': qdiff, 'fee_difference': fdiff,
                                    'quantity_fields': list(QUANTITY_FIELDS), 'fee_fields': fee_fields, 'quantity_missing': q_missing, 'fee_missing': f_missing,
                                    'structural_quantity_fields': q_structural, 'structural_fee_fields': f_structural,
                                    'quantity_complete': q_complete, 'fee_complete': f_complete, 'totals_verified': reliable,
                                    'price_corrected': corrected, 'issue_count': len(local_issues),
                                    'quantity_scope': '省间日前+省间长协+应急外送+非市场+实际上网+月末调整+调试期+跨月；省内日前/合约、绿电、机制、国补量不重复相加。',
                                    'fee_scope': '分项费用按原符号相加，含环境价值及单列退补调整；不加总结算费自身、机制电费、国补电费；不因考核费为正数自动取反。'}
        record.setdefault('notes', [])
        if corrected:
            record['notes'].append(changes[-1]['reason'])
        issues.extend(local_issues)
        counts['records'] += 1; counts['totals_verified'] += int(reliable); counts['price_corrected'] += int(corrected)
        counts['pending_records'] += int(any(i['severity'] in ('warning', 'error') for i in local_issues))
    return {'records': result, 'issues': issues, 'changes': changes,
            'stats': {**dict(counts), 'issues': len(issues), 'issue_codes': dict(Counter(i['code'] for i in issues))}}
