"""Session-local orchestration. Uploaded business data is never written to disk."""
from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from math import isfinite
from typing import Any, Callable, Sequence
from zipfile import BadZipFile, ZipFile

MAX_UPLOAD_BYTES = 50 * 1024 * 1024
MAX_TOTAL_BYTES = 120 * 1024 * 1024
MAX_EXPANDED_BYTES = 320 * 1024 * 1024


@dataclass
class ProcessingResult:
    workbook: bytes
    records: list[dict]
    headers: list[str]
    issues: list[dict]
    changes: list[dict]
    stats: dict[str, Any]
    hours_metadata: dict[str, Any]


def validate_upload(data: bytes, filename: str) -> None:
    if not filename.lower().endswith(".xlsx"):
        raise ValueError(f"{filename}：请上传 .xlsx 格式文件。")
    if not data:
        raise ValueError(f"{filename}：文件为空。")
    if len(data) > MAX_UPLOAD_BYTES:
        raise ValueError(f"{filename}：文件超过 50 MB，请分批上传。")
    try:
        with ZipFile(BytesIO(data)) as package:
            entries = package.infolist()
            if len(entries) > 10000 or sum(x.file_size for x in entries) > MAX_EXPANDED_BYTES:
                raise ValueError(f"{filename}：工作簿解压后的数据过大，请缩小文件。")
            names = set(package.namelist())
            if "xl/workbook.xml" not in names or "[Content_Types].xml" not in names:
                raise ValueError(f"{filename}：不是可识别的 Excel 工作簿。")
            if any("vbaProject" in name for name in names):
                raise ValueError(f"{filename}：请另存为不含宏的 .xlsx 文件后再上传。")
            if any(x.flag_bits & 1 for x in entries):
                raise ValueError(f"{filename}：请先取消文件加密。")
    except BadZipFile as exc:
        raise ValueError(f"{filename}：无法读取，请确认文件完整且未加密。") from exc


def _items(result: dict, key: str) -> list[dict]:
    output = []
    for item in result.get(key, []) or []:
        output.append(dict(item) if isinstance(item, dict) else {"detail": str(item)})
    return output


def _assign_year(records: list[dict], year: int, filename: str) -> None:
    for row in records:
        actual = row.get("year")
        if actual not in (None, "", year, str(year)):
            raise ValueError(f"{filename}：发现 {actual} 年数据，与选择的 {year} 年不一致。请按年度分别生成。")
        row["year"] = year


def update_workbook(
    previous: bytes,
    collections: Sequence[tuple[str, bytes]],
    *,
    previous_name: str = "上一版分析表.xlsx",
    year: int = 2026,
    conflict_policy: str = "keep_old",
    progress: Callable[[str], None] | None = None,
) -> ProcessingResult:
    from .export import generate_workbook, validate_template
    from .hours import build_hours_model
    from .ingest import merge_records, read_analysis, read_collection
    from .reconcile import reconcile_records
    from .reporting import audit_row

    if not 2000 <= year <= 2100:
        raise ValueError("统计年份应在 2000—2100 之间。")
    if conflict_policy not in {"keep_old", "prefer_new"}:
        raise ValueError("无法识别的冲突处理方式。")
    if not collections or len(collections) > 15:
        raise ValueError("请上传 1—15 份最新收集表。")
    if len(previous) + sum(len(data) for _, data in collections) > MAX_TOTAL_BYTES:
        raise ValueError("上传文件合计超过 120 MB，请分批更新。")
    report = progress or (lambda message: None)
    validate_upload(previous, previous_name)
    for name, data in collections:
        validate_upload(data, name)

    report("检查原透视表模板")
    validate_template(previous)
    report("读取上一版分析表")
    old = read_analysis(previous, previous_name)
    headers, old_records = old["headers"][:67], old["records"]
    if len(headers) != 67 or not old_records:
        raise ValueError("上一版分析表没有完整的 67 列对标明细，请上传此前生成的分析表。")
    _assign_year(old_records, year, previous_name)
    issues = _items(old, "warnings") + _items(old, "issues")
    incoming = []
    for name, data in collections:
        report(f"识别字段与容量单位：{name}")
        parsed = read_collection(data, name, headers, year)
        _assign_year(parsed["records"], year, name)
        incoming.extend(parsed["records"])
        issues.extend(_items(parsed, "warnings"))
        issues.extend(_items(parsed, "issues"))
    if not incoming:
        raise ValueError("最新收集表中没有识别到可合并的场站月度记录，请检查月份和场站表头。")

    report("匹配场站并交叉补全")
    merged = merge_records(old_records, incoming, conflict_policy=conflict_policy)
    changes = _items(merged, "changes")
    issues.extend(_items(merged, "issues"))
    report("核对总量、总费及价格的组成关系")
    checked = reconcile_records(merged["records"])
    changes.extend(_items(checked, "changes"))
    issues.extend(_items(checked, "issues"))
    records = checked["records"]
    report("重建任意选月利用小时与原透视表数据源")
    model = build_hours_model(records, headers, year=year)
    audit = ([audit_row(row, "变更", headers) for row in changes]
             + [audit_row(row, "待核对", headers) for row in issues])
    report("生成新版分析工作簿")
    output = generate_workbook(previous, records, model, audit, year=year)
    return ProcessingResult(
        workbook=output,
        records=records,
        headers=headers,
        issues=issues,
        changes=changes,
        stats={"old_records": len(old_records), "incoming_records": len(incoming),
               "output_records": len(records), "change_count": len(changes),
               "issue_count": len(issues), "merge": merged.get("stats", {}),
               "reconcile": checked.get("stats", {}), "year": year},
        hours_metadata=model.get("metadata", {}),
    )


def _number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and isfinite(value)


def preview_summary(records: list[dict], months: list[int], companies: list[str], types: list[str], year: int) -> list[dict]:
    """One filter context for trade metrics and complete-sample hours."""
    from .hours import evaluate_hours

    if not months or not companies or not types:
        return []
    selected = [r for r in records if r.get("month") in months
                and r["values"][66] in companies and r["values"][5] in types]
    results = []
    for company in companies + ["全部所选集团"]:
        subset = [r for r in selected if company == "全部所选集团" or r["values"][66] == company]
        if not subset:
            continue
        def total(index):
            return sum(r["values"][index] for r in subset if _number(r["values"][index]))
        q = total(59)
        def ratio(amount):
            return amount / q if q else None
        # Identity and phase evidence must be established from the complete
        # dataset before filtering. Nonselected months may prove two same-name
        # phases are distinct, even if each appears once in the selected months.
        group_filter = companies if company == "全部所选集团" else [company]
        h = evaluate_hours(records, months, filters={"co": group_filter, "type": types}, year=year)
        combined = total(22) + total(28) + total(25) + total(10) + total(13)
        results.append({
            "集团": company, "总结算电量（MWh）": q,
            "实时电价贡献（元/MWh）": ratio(total(22)),
            "中长期贡献（元/MWh）": ratio(total(28)),
            "日前贡献（元/MWh）": ratio(total(25)),
            "省间贡献（元/MWh）": ratio(total(10) + total(13)),
            "交易电价（元/MWh）": ratio(combined),
            "累计利用小时（h）": h.get("hours"),
            "完整样本站数": h.get("station_count", 0),
        })
    return results
