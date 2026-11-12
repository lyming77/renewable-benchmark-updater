"""Human-readable change and reconciliation records for the UI and workbook."""
from __future__ import annotations

import json


_LABELS = {
    "quantity": "电量", "price": "电价", "fee": "电费", "product": "电量×电价",
    "reported": "报送总值", "known_sum": "已知组成项合计", "difference": "报送值减组成合计",
    "rounding_bound": "现报精度舍入上界", "tolerance": "容差",
    "known_quantity_sum": "已知电量合计", "reported_quantity": "报送总电量",
    "known_fee_sum": "已知费用合计", "reported_fee": "报送总电费",
    "reported_price": "报送均价", "ratio_of_reported_totals": "报送总费÷报送总量",
    "first_value": "新来源一数值", "second_value": "新来源二数值",
    "first_source": "新来源一", "second_source": "新来源二",
    "candidates": "候选匹配记录", "original": "原报内容", "unit": "来源单位",
}
_KINDS = {"added_record": "新增记录", "filled_empty": "补充空值",
          "overwritten_conflict": "按新值更新", "verified_summary_price": "核实后修正均价"}


def _text(value):
    if value is None:
        return ""
    if isinstance(value, (dict, list, tuple)):
        return json.dumps(value, ensure_ascii=False, default=str)
    return str(value)


def audit_row(item: dict, kind: str, headers: list[str] | None = None) -> dict:
    from .ingest import HEADERS, EXTRA_HEADERS

    source_record = item.get("after") if isinstance(item.get("after"), dict) else {}
    values = source_record.get("values", [])
    field = item.get("field") or item.get("fields") or ""
    if isinstance(field, int):
        field = (headers or HEADERS)[field] if field < 67 else EXTRA_HEADERS.get(field, str(field))
    if isinstance(field, list):
        field = "、".join(map(str, field))
    before = item.get("before", item.get("old_value", item.get("reported", item.get("reported_price"))))
    after = item.get("after", item.get("new_value", item.get("known_sum", item.get("ratio_of_reported_totals"))))
    if isinstance(after, dict):
        after = "新增完整记录，详见Sheet2"
    level = {"error": "需处理", "warning": "待核对", "info": "说明"}.get(item.get("severity"), "变更" if kind == "变更" else "待核对")
    details = "；".join(f"{label}={_text(item[key])}" for key, label in _LABELS.items() if key in item)
    return {
        "类型": _KINDS.get(item.get("kind"), kind), "级别": level,
        "集团": item.get("company", source_record.get("company", "")),
        "年份": item.get("year", source_record.get("year", "")),
        "月份": item.get("month", source_record.get("month", "")),
        "场站": item.get("station") or (values[2] or values[3] if len(values) > 3 else ""),
        "字段": field, "原值": before, "新值或计算值": after, "差额": item.get("difference"),
        "说明": item.get("message") or item.get("reason") or _KINDS.get(item.get("kind"), ""),
        "原来源": _text(item.get("old_source", item.get("source", source_record.get("source")))),
        "新来源": _text(item.get("new_source", "")),
        "来源单元格": _text(item.get("source_cell", item.get("cell", ""))),
        "核对明细": details,
    }
