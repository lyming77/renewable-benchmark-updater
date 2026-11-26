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
          "overwritten_conflict": "按新值更新", "verified_summary_price": "核实后修正均价",
          "normalized_format": "统一日期格式", "capacity_unit_corrected": "核实后换算容量"}


RECONCILIATION_HEADERS = [
    "Sheet2行", "月份", "集团", "场站",
    "实时电量（MWh）", "月末调整量（MWh）", "跨月调整量（MWh）", "其他已列物理量（MWh）",
    "报送总量（MWh）", "已列分项量和（MWh）", "量差额（MWh）", "总量核验",
    "报送总费（元）", "已列分项费和（元）", "费差额（元）", "总费核验",
    "核验前均价（元/MWh）", "报送总费÷总量（元/MWh）", "价差额（元/MWh）", "表内均价（元/MWh）",
    "原价交叉验证及处理", "未确认电量项", "未确认费用项", "明细来源",
]


def reconciliation_rows(records):
    from .ingest import EXTRA_HEADERS, HEADERS, is_number
    from .reconcile import INDEPENDENT_QUANTITY_FIELDS

    def names(fields):
        return "、".join(HEADERS[i] if i < len(HEADERS) else EXTRA_HEADERS[i] for i in fields)

    for row, record in enumerate(records, 5):
        review = record.get("reconciliation")
        if review is None:
            continue
        v = record["values"]
        independent = [v[i] for i in INDEPENDENT_QUANTITY_FIELDS if is_number(v[i])]
        original, implied = review["reported_price"], review["expected_price"]
        price_delta = original - implied if is_number(original) and implied is not None else None
        price_notes = [review["price_evidence"], review["quantity_scope_note"]]
        if review["price_corrected"]:
            price_notes.append("已按已验证的总费÷总量重算均价")
        yield [
            row, record["month"], record["company"], v[2] or v[3],
            v[20], v[29], v[56], round(sum(independent), 8) if independent else None,
            v[59], review["known_quantity_sum"], _audit_difference(review["quantity_difference"]), review["quantity_status"],
            v[61], review["known_fee_sum"], _audit_difference(review["fee_difference"]), review["fee_status"],
            original, round(implied, 8) if implied is not None else None, _audit_difference(price_delta), v[60],
            "；".join(filter(None, price_notes)), names(review["quantity_missing"]), names(review["fee_missing"]), v[0],
        ]


def _audit_difference(value):
    return round(value, 8) if value is not None else None


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
