"""Complete-sample utilization hours for a shared, ordinary WPS pivot.

This module contains no workbook or business data and has no third-party
dependencies.  The first 67 columns follow the benchmark's A:BO schema.  U is
actual on-grid MWh, G is installed MW.  Transaction rows are never removed to
make the hours population complete.

The generated source uses one bucket per stable coverage pattern, and one
bucket per station whose filter dimensions change.  Calendar control rows make
the selected months observable even when a company has no report in a month.
"""

from __future__ import annotations

import calendar
import hashlib
import json
import math
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime
from numbers import Real
from typing import Any, Iterable


ORIGINAL_FIELD_COUNT = 67
ORIGINAL_CALCULATED_FIELD_COUNT = 16
WPS_FIELD_LIMIT = 255
SOURCE_ROW_START = 5
DIMENSION_COLUMNS = {
    "co": 66,
    "type": 5,
    "region": 4,
    "settlement": 2,
    "dispatch": 3,
    "commission": 7,
}
_FILTER_ALIASES = {
    "company": "co", "companies": "co", "集团": "co", "公司": "co",
    "station_type": "type", "类型": "type", "地区": "region",
    "settlement_name": "settlement", "结算名称": "settlement",
    "dispatch_name": "dispatch", "调度名称": "dispatch",
    "commission_date": "commission", "投产日期": "commission",
}


class HoursModelError(ValueError):
    """Input cannot support an unambiguous complete-station hours model."""


class HoursFieldLimitError(HoursModelError):
    """A generated native WPS cache would exceed its field budget."""


def _is_number(value: Any) -> bool:
    return isinstance(value, Real) and not isinstance(value, bool) and math.isfinite(value)


def _typed_key(value: Any) -> tuple[str, Any]:
    """Preserve filter labels and value types; normalize nothing in source rows."""
    if isinstance(value, datetime):
        return ("datetime", value.isoformat())
    if isinstance(value, date):
        return ("date", value.isoformat())
    try:
        hash(value)
    except TypeError as exc:
        raise HoursModelError(f"筛选字段必须是单个单元格值，收到 {type(value).__name__}") from exc
    return (type(value).__name__, value)


def _name(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(value)))


def _commission_key(value: Any) -> str | None:
    """Only an explicit valid date may separate simultaneous same-name phases."""
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    if not isinstance(value, str):
        return None
    text = unicodedata.normalize("NFKC", value).strip()
    match = re.match(r"^(\d{4})[年./-](\d{1,2})[月./-](\d{1,2})(?:日|\b|T)", text)
    if not match:
        return None
    try:
        return date(*(int(part) for part in match.groups())).isoformat()
    except ValueError:
        return None


def _month(value: Any, year: int) -> int:
    if isinstance(value, (date, datetime)):
        if value.year != year:
            raise HoursModelError(f"月份 {value!s} 不属于统计年度 {year}；请分年度生成。")
        return value.month
    if _is_number(value) and int(value) == value and 1 <= value <= 12:
        return int(value)
    text = unicodedata.normalize("NFKC", str(value or "")).strip()
    match = re.fullmatch(r"(?:(\d{4})(?:年|[-/]))?(\d{1,2})(?:月份|月)?", text)
    if not match or not 1 <= int(match[2]) <= 12:
        raise HoursModelError(f"无法识别月份 {value!r}；应为1—12月的单月标签。")
    if match[1] and int(match[1]) != year:
        raise HoursModelError(f"月份 {value!r} 不属于统计年度 {year}；请分年度生成。")
    return int(match[2])


def _prepare(records: list[dict], year: int) -> tuple[list[dict], list[dict], list[dict]]:
    if isinstance(year, bool) or not isinstance(year, int) or not 1900 <= year <= 9999:
        raise HoursModelError("统计年度必须是1900—9999之间的整数。")
    if not records:
        raise HoursModelError("没有可用于建立利用小时模型的月度记录。")
    actual: list[dict] = []
    groups: dict[tuple, list[dict]] = defaultdict(list)
    issues: list[dict] = []
    for index, item in enumerate(records):
        if not isinstance(item, dict) or not isinstance(item.get("values"), (list, tuple)):
            raise HoursModelError(f"第{index + 1}条记录缺少 values 单元格列表。")
        if len(item["values"]) < ORIGINAL_FIELD_COUNT:
            raise HoursModelError(f"第{index + 1}条记录不足67列，不能按A:BO识别。")
        values = list(item["values"][:ORIGINAL_FIELD_COUNT])
        source_row = index + SOURCE_ROW_START
        company = values[66]
        if company is None or not str(company).strip():
            raise HoursModelError(f"Sheet2第{source_row}行缺少集团，不能建立跨月场站身份。")
        station_name = _name(values[2]) or _name(values[3])
        if not station_name:
            raise HoursModelError(f"Sheet2第{source_row}行结算名称、调度名称均为空。")
        month = _month(values[1], year)
        dims = {key: values[col] for key, col in DIMENSION_COLUMNS.items()}
        for value in dims.values():
            _typed_key(value)
        row = {
            "values": values,
            "source": str(item.get("source") or ""),
            "source_row": source_row,
            "month": month,
            "days": calendar.monthrange(year, month)[1],
            "dimensions": dims,
            "quantity_mwh": values[20],
            "capacity_mw": values[6],
            "valid_numeric": _is_number(values[20]) and _is_number(values[6]) and values[6] > 0,
            "commission_key": _commission_key(values[7]),
        }
        actual.append(row)
        groups[(_typed_key(company), station_name)].append(row)

    units: list[dict] = []
    for (company_key, normalized_name), group in sorted(groups.items(), key=lambda pair: repr(pair[0])):
        company = group[0]["values"][66]
        collision = max(Counter(row["month"] for row in group).values()) > 1
        pieces = {None: group}
        basis = "集团+规范空格和全半角后的结算名称；缺失时使用同名调度名称，不按容量识别。"
        dates = {row["commission_key"] for row in group}
        if collision:
            counts = Counter((row["month"], row["commission_key"]) for row in group)
            if None not in dates and len(dates) > 1 and max(counts.values()) == 1:
                pieces = defaultdict(list)
                for row in group:
                    pieces[row["commission_key"]].append(row)
                basis = "同名同月多行，按明确且能使每月唯一的投产日期拆分独立期批，不按容量拆分。"
                issues.append({"kind": "same_name_phases_split", "company": company,
                               "name": normalized_name, "dates": sorted(dates),
                               "source_rows": [row["source_row"] for row in group]})
            else:
                rows = [row["source_row"] for row in group]
                raise HoursModelError(
                    f"场站身份有歧义：集团{company!s}、名称{normalized_name!s}在同月有多条记录"
                    f"（Sheet2行{','.join(map(str, rows[:12]))}）。投产日期不足以唯一拆分。"
                    "请先确认独立分期/重复报告并修正名称或明确投产日期；不会按容量猜测合并。"
                )
        elif len(dates - {None}) > 1:
            issues.append({"kind": "commission_date_changed", "company": company,
                           "name": normalized_name, "dates": sorted(dates - {None}),
                           "source_rows": [row["source_row"] for row in group],
                           "note": "未出现同月同名冲突，沿用结算名称ID；应核实日期更正或未同时出现的同名分期。"})
        for split_date, members in sorted(pieces.items(), key=lambda pair: str(pair[0])):
            identity = [list(company_key), normalized_name, split_date]
            identity_text = json.dumps(identity, ensure_ascii=False, default=str, separators=(",", ":"))
            unit_id = "U" + hashlib.sha256(identity_text.encode("utf-8")).hexdigest()[:16]
            members = sorted(members, key=lambda row: (row["month"], row["source_row"]))
            dimension_keys = {
                tuple(_typed_key(row["dimensions"][key]) for key in DIMENSION_COLUMNS)
                for row in members
            }
            changes = {}
            for key in DIMENSION_COLUMNS:
                seen = {}
                for row in members:
                    seen.setdefault(_typed_key(row["dimensions"][key]), row["dimensions"][key])
                if len(seen) > 1:
                    changes[key] = list(seen.values())
            display = next((str(row["values"][2]) for row in members if _name(row["values"][2])), str(members[0]["values"][3]))
            unit = {
                "unit_id": unit_id, "identity_key": identity,
                "identity_basis": basis, "company": company,
                "station_name": display,
                "display_name": display + (f"〔投产{split_date}〕" if split_date else ""),
                "commission_split_date": split_date,
                "kind": "stable" if len(dimension_keys) == 1 else "special",
                "valid_mask": sum(1 << (row["month"] - 1) for row in members if row["valid_numeric"]),
                "valid_months": [row["month"] for row in members if row["valid_numeric"]],
                "source_rows": [row["source_row"] for row in members],
                "monthly": members, "dimension_changes": changes,
            }
            for row in members:
                row["unit_id"] = unit_id
            units.append(unit)
    return actual, units, issues


def _filter_sets(filters: dict | None, headers: list[str] | None = None) -> dict:
    result = {}
    names = {headers[col]: key for key, col in DIMENSION_COLUMNS.items()} if headers else {}
    for field, selected in (filters or {}).items():
        key = names.get(field, _FILTER_ALIASES.get(field, field))
        if key not in DIMENSION_COLUMNS:
            raise HoursModelError(f"利用小时不支持筛选字段 {field!r}；请使用公司、类型、地区、结算名、调度名或投产日期。")
        if isinstance(selected, (list, tuple, set, frozenset)):
            values = selected
        else:
            values = [selected]
        result[key] = {_typed_key(value) for value in values}
    return result


def _matches(row: dict, filters: dict) -> bool:
    return all(_typed_key(row["dimensions"][key]) in accepted for key, accepted in filters.items())


def _evaluate_prepared(actual: list[dict], units: list[dict], selected: list[int], filters: dict, year: int) -> dict:
    selected_set = set(selected)
    days = sum(calendar.monthrange(year, month)[1] for month in selected)
    eligible = []
    energy = capacity_days = 0.0
    capacity_changes = []
    for unit in units:
        included = [row for row in unit["monthly"] if row["month"] in selected_set and _matches(row, filters)]
        if not selected or len(included) != len(selected) or any(not row["valid_numeric"] for row in included):
            continue
        if {row["month"] for row in included} != selected_set:
            continue
        eligible.append(unit["unit_id"])
        energy += sum(row["quantity_mwh"] for row in included)
        capacity_days += sum(row["capacity_mw"] * row["days"] for row in included)
        if len({row["capacity_mw"] for row in included}) > 1:
            capacity_changes.append(unit["unit_id"])
    transaction_rows = [row for row in actual if row["month"] in selected_set and _matches(row, filters)]
    matching_control = any(_matches(row, filters) for row in actual)
    effective_days = days if matching_control else 0
    result = {
        "hours": energy * days / capacity_days if capacity_days > 0 and days else None,
        "station_count": len(eligible), "energy_mwh": energy,
        "capacity_mw": capacity_days / days if days and eligible else None,
        "capacity_days": capacity_days, "calendar_days": effective_days,
        "selection_calendar_days": days, "selected_months": selected,
        "month_count": len(selected) if matching_control else 0,
        "eligible_unit_ids": eligible, "all_transaction_row_count": len(transaction_rows),
        "all_transaction_energy_mwh": sum(row["quantity_mwh"] for row in transaction_rows if _is_number(row["quantity_mwh"])),
        "capacity_change_unit_ids": capacity_changes,
        "capacity_basis": "月报容量按所选月份日历天数近似加权；月内变更时点未知。",
        "missing_entire_months": sorted(selected_set - {row["month"] for row in actual}),
    }
    result.update(HH_D=effective_days, HH_N=result["month_count"], HH_E=energy,
                  HH_CD=capacity_days, HH_Count=len(eligible),
                  HH_C=capacity_days / days if days else 0)
    return result


def evaluate_hours(records: list[dict], selected_months: Iterable[Any], filters: dict | None = None, year: int = 2026) -> dict:
    """Evaluate the same complete population as the native shared pivot.

    ``filters`` maps a supported dimension key (co/company, type, region,
    settlement, dispatch, commission) to one value or a list of selected values.
    Omit a dimension for all values; ``None`` selects the actual blank label.
    Missing reports never become zero reports.  No complete population returns
    hours=None, while a complete, truly zero-energy population returns 0 hours.
    """
    actual, units, issues = _prepare(records, year)
    selected = sorted({_month(month, year) for month in selected_months})
    result = _evaluate_prepared(actual, units, selected, _filter_sets(filters), year)
    result["identity_issues"] = issues
    return result


def build_hours_model(records: list[dict], headers: list[str], year: int = 2026) -> dict:
    """Build cached helper rows and ordinary IF-only pivot calculated fields.

    Original input cells A:BO are copied without alteration.  Every rebuild
    recalculates station identities, valid-month masks and classification
    combinations.  The caller must preserve the original 16 calculated fields
    and replace only its utilization-hours formula with ``hours_formula``.
    Native source filters are limited to the six dimensions in this module and
    the original month field.  Source/quantity/helper filters break the calendar
    control proof and must not be presented as supported hours filters.
    """
    if len(headers) < ORIGINAL_FIELD_COUNT:
        raise HoursModelError("表头不足67列，无法构建A:BO小时联动源。")
    original_headers = list(headers[:ORIGINAL_FIELD_COUNT])
    if any(not isinstance(value, str) or not value.strip() for value in original_headers):
        raise HoursModelError("前67列表头必须是非空文本。")
    if len(set(original_headers)) != ORIGINAL_FIELD_COUNT:
        raise HoursModelError("前67列表头存在重名，普通透视表需要唯一字段名。")
    actual, units, issues = _prepare(records, year)
    month_numbers = sorted({row["month"] for row in actual})
    months = []
    for month in month_numbers:
        labels = {_typed_key(row["values"][1]) for row in actual if row["month"] == month}
        if len(labels) != 1:
            raise HoursModelError(f"{month}月存在多个原月份标签（如1月/1月份）；请先统一月份列，避免透视多选重复月份。")
        label = next(row["values"][1] for row in actual if row["month"] == month)
        months.append({"month": month, "days": calendar.monthrange(year, month)[1],
                       "hours": calendar.monthrange(year, month)[1] * 24,
                       "flag_field": f"F{month:02d}", "label": label,
                       "source_month_label": label})
    calendar_by_month = {item["month"]: item for item in months}
    stable = [unit for unit in units if unit["kind"] == "stable"]
    special = [unit for unit in units if unit["kind"] == "special"]
    patterns = sorted({unit["valid_mask"] for unit in stable})
    stable_buckets = []
    special_buckets = []
    for index, mask in enumerate(patterns, 1):
        bucket_id = f"P{index:02d}"
        members = [unit for unit in stable if unit["valid_mask"] == mask]
        bucket = {"bucket_id": bucket_id, "mask": mask,
                  "valid_months": [month for month in month_numbers if mask & (1 << (month - 1))],
                  "missing_months": [month for month in month_numbers if not mask & (1 << (month - 1))],
                  "unit_ids": [unit["unit_id"] for unit in members], "unit_count": len(members),
                  "e_field": f"E_{bucket_id}", "cd_field": f"CD_{bucket_id}", "count_field": f"N_{bucket_id}"}
        stable_buckets.append(bucket)
        for unit in members:
            unit["bucket_id"] = bucket_id
    for index, unit in enumerate(special, 1):
        bucket_id = f"U{index:02d}"
        unit["bucket_id"] = bucket_id
        special_buckets.append({"bucket_id": bucket_id, "unit_id": unit["unit_id"],
                                "station_name": unit["display_name"], "mask": unit["valid_mask"],
                                "valid_months": unit["valid_months"], "source_rows": unit["source_rows"],
                                "e_field": f"E_{bucket_id}", "cd_field": f"CD_{bucket_id}",
                                "good_days_field": f"GD_{bucket_id}", "dimension_changes": unit["dimension_changes"]})
    fields = original_headers + [item["flag_field"] for item in months]
    for bucket in stable_buckets:
        fields.extend(bucket[key] for key in ("e_field", "cd_field", "count_field"))
    for bucket in special_buckets:
        fields.extend(bucket[key] for key in ("e_field", "cd_field", "good_days_field"))
    fields.extend(["记录用途", "单元ID", "小时源变更"])
    if len(set(fields)) != len(fields):
        raise HoursModelError("原字段与小时辅助字段重名，请先重命名冲突的原字段。")
    expected_total = len(fields) + ORIGINAL_CALCULATED_FIELD_COUNT + 6
    if expected_total > WPS_FIELD_LIMIT:
        raise HoursFieldLimitError(
            f"本批数据需要{expected_total}个透视字段，超过WPS普通透视表{WPS_FIELD_LIMIT}字段预算"
            f"（原字段67、月份{len(months)}、稳定覆盖桶{len(stable_buckets)}×3、"
            f"分类变动场站{len(special_buckets)}×3、元信息3、原计算字段16、新计算字段6）。"
            "请核实可统一的分类别名或拆分年度分析；不会删站、删月份或改用不完整样本生成文件。"
        )
    field_index = {name: index for index, name in enumerate(fields)}
    unit_by_id = {unit["unit_id"]: unit for unit in units}
    buckets = {bucket["bucket_id"]: bucket for bucket in stable_buckets + special_buckets}
    combinations = {}
    for row in actual:
        key = tuple(_typed_key(row["dimensions"][dim]) for dim in DIMENSION_COLUMNS)
        if key not in combinations:
            combinations[key] = {"combination_id": f"C{len(combinations) + 1:04d}",
                                 "dimensions": row["dimensions"].copy(), "source_rows": []}
        combinations[key]["source_rows"].append(row["source_row"])
        row["dimension_combination_id"] = combinations[key]["combination_id"]
    rows = []
    row_metadata = []
    for row in actual:
        unit = unit_by_id[row["unit_id"]]
        bucket = buckets[unit["bucket_id"]]
        values = row["values"].copy() + [0] * (len(fields) - ORIGINAL_FIELD_COUNT)
        if row["valid_numeric"]:
            values[field_index[bucket["e_field"]]] = row["quantity_mwh"]
            values[field_index[bucket["cd_field"]]] = row["capacity_mw"] * row["days"]
            count_field = bucket["count_field"] if unit["kind"] == "stable" else bucket["good_days_field"]
            values[field_index[count_field]] = 1 if unit["kind"] == "stable" else row["days"]
        values[field_index["记录用途"]] = "原始月记录"
        values[field_index["单元ID"]] = unit["unit_id"]
        rows.append(values)
        row_metadata.append({key: row[key] for key in ("source_row", "source", "month", "unit_id", "valid_numeric", "dimension_combination_id")}
                            | {"kind": "actual", "bucket_id": unit["bucket_id"]})
    for combination in combinations.values():
        for item in months:
            values = [None] * ORIGINAL_FIELD_COUNT + [0] * (len(fields) - ORIGINAL_FIELD_COUNT)
            values[1] = item["label"]
            for key, col in DIMENSION_COLUMNS.items():
                values[col] = combination["dimensions"][key]
            values[field_index[item["flag_field"]]] = 1
            values[field_index["记录用途"]] = "月份控制"
            values[field_index["单元ID"]] = None
            rows.append(values)
            row_metadata.append({"kind": "sentinel", "source_row": None, "source": "", "unit_id": None,
                                 "bucket_id": None, "month": item["month"], "dimension_combination_id": combination["combination_id"]})

    quote = lambda name: "'" + name.replace("'", "''") + "'"
    # Native WPS calculated fields accept IF(field comparison), but dynamic
    # AND(field, ...) has produced #N/A in the verified target version.  All
    # flags are nonnegative, so SUM(missing-month flags)=0 is the same test.
    stable_gates = {bucket["bucket_id"]: (
        "+".join(quote(calendar_by_month[month]["flag_field"]) for month in bucket["missing_months"]) + "=0"
        if bucket["missing_months"] else None) for bucket in stable_buckets}
    # The enclosing HH_D>0 guard makes a separate GD>0 test redundant.
    special_gates = {bucket["bucket_id"]: f"{bucket['good_days_field']}=HH_D" for bucket in special_buckets}
    energy_terms, cd_terms, count_terms = [], [], []
    for bucket in stable_buckets:
        gate = stable_gates[bucket["bucket_id"]]
        for key, target in (("e_field", energy_terms), ("cd_field", cd_terms), ("count_field", count_terms)):
            target.append(f"IF({gate},{bucket[key]},0)" if gate else bucket[key])
    for bucket in special_buckets:
        gate = special_gates[bucket["bucket_id"]]
        energy_terms.append(f"IF({gate},{bucket['e_field']},0)")
        cd_terms.append(f"IF({gate},{bucket['cd_field']},0)")
    stable_count = f"({'+'.join(count_terms) or '0'})/HH_N"
    special_count = "+".join(f"IF({special_gates[bucket['bucket_id']]},1,0)" for bucket in special_buckets)
    count_expression = stable_count + ("+" + special_count if special_count else "")
    def guarded(expression):
        return f"IF({quote('小时源变更')}=0,IF(HH_D>0,{expression},0),0)"

    formulas = [
        {"name": "HH_D", "formula": f"IF({quote('小时源变更')}=0," + "+".join(f"IF({quote(item['flag_field'])}>0,{item['days']},0)" for item in months) + ",0)", "meaning": "筛选所选月份日历天数，每月只计算一次。"},
        {"name": "HH_N", "formula": f"IF({quote('小时源变更')}=0," + "+".join(f"IF({quote(item['flag_field'])}>0,1,0)" for item in months) + ",0)", "meaning": "筛选所选月份数。"},
        {"name": "HH_E", "formula": guarded('+'.join(energy_terms) or '0'), "meaning": "所选各月都有有效数据且符合分类筛选的同一批站电量MWh。"},
        {"name": "HH_CD", "formula": guarded('+'.join(cd_terms) or '0'), "meaning": "同一批完整站容量MW乘各月日历天数的合计。"},
        {"name": "HH_Count", "formula": guarded(count_expression), "meaning": "所选月份及分类筛选内的完整站数。"},
        {"name": "HH_C", "formula": guarded('HH_CD/HH_D'), "meaning": "同一批完整站按月日历天数加权的平均容量MW之和。"},
    ]
    hours_formula = f"IF({quote('小时源变更')}=0,IF(HH_D>0,IF(HH_CD>0,HH_E*HH_D/HH_CD,0/0),0/0),0/0)"
    companies = list({_typed_key(row["values"][66]): row["values"][66] for row in actual}.values())
    companies.sort(key=str)
    summary_all = _evaluate_prepared(actual, units, month_numbers, {}, year)
    summary_by_company = {
        company: _evaluate_prepared(actual, units, month_numbers, {"co": {_typed_key(company)}}, year)
        for company in companies
    }
    return {
        "headers": fields, "fields": fields, "field_index": field_index, "rows": rows,
        "row_metadata": row_metadata, "formulas": formulas, "hours_formula": hours_formula,
        "source_sheet": "小时联动源", "header_row": 6,
        "units": units, "companies": companies, "months": months,
        "stable_buckets": stable_buckets, "special_buckets": special_buckets,
        "dimension_fields": [{"key": key, "source_col_index": col, "source_field_name": original_headers[col]} for key, col in DIMENSION_COLUMNS.items()],
        "dimension_combinations": list(combinations.values()),
        "summary_all": summary_all, "summary_by_company": summary_by_company,
        "metadata": {
            "year": year, "actual_row_count": len(actual), "sentinel_row_count": len(combinations) * len(months),
            "source_row_count": len(rows), "unit_count": len(units), "stable_unit_count": len(stable),
            "special_unit_count": len(special), "stable_pattern_count": len(patterns),
            "dimension_combination_count": len(combinations), "physical_field_count": len(fields),
            "original_calculated_field_count": ORIGINAL_CALCULATED_FIELD_COUNT,
            "new_calculated_field_count": len(formulas), "total_cache_field_count": expected_total,
            "wps_field_limit": WPS_FIELD_LIMIT, "original_sheet": "Sheet2", "original_header_row": 4,
            "original_source_row_start": SOURCE_ROW_START, "original_field_count": ORIGINAL_FIELD_COUNT,
            "quantity_column": "U", "quantity_unit": "MWh", "capacity_column": "G", "capacity_unit": "MW",
            "supported_filters": list(DIMENSION_COLUMNS), "identity_issues": issues,
            "invalid_source_rows": [row["source_row"] for row in actual if not row["valid_numeric"]],
            "numeric_complete_policy": "所选每月恰1条，U为有限数值且G为有限正数；真实0电量有效；缺月不补0。",
            "original_transaction_scope": "保留全部原始交易记录，完整资格仅用于利用小时及小时辅助指标。",
            "capacity_basis": "月报容量按所选月日历天数近似加权；不按容量建场站ID，不推测月内变更日。",
            "rebuild_on": ["新增或删除记录", "月份变化", "身份或六个筛选维度变化", "U或G有效资格变化"],
            "refresh_contract": "云端每次上传重建全部辅助池；导出后如手动改身份/月/分类/有效资格，须重新上传生成。",
            "unsupported_filters": ["源文件", "单元ID", "记录用途", "交易数值字段", "未进入控制行的其他维度"],
            "no_sample_policy": "无完整样本时hours为None，native公式返回0/0供透视显示错误文本；真实0电量返回0。",
            "formula_functions": ["IF"], "all_actual_original_values_preserved": True,
        },
    }
