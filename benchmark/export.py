"""Rebuild an uploaded analysis workbook while retaining its native controls.

The cloud deployment intentionally uses only Python's ZIP/XML facilities. It
does not import the local authoring scripts, a desktop app, or any business file.
"""
from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from datetime import datetime, timezone
import math
import posixpath
import re
from xml.etree import ElementTree as ET

from .native import (
    CT, NS, PKGREL, REL, X14, ExportError, Package, arithmetic, cell, col,
    colnum, key, local, logical, parse, pivot_rebuild, rebuild_cache,
    rebuild_slicer, relpath, resolve, scalar, shared_strings, tag, value, global_formats_only,
)

__all__ = ["generate_workbook", "validate_template", "ExportError"]


def _cache_parts(pkg):
    wb = pkg.root("xl/workbook.xml"); relations = pkg.relation_map("xl/workbook.xml")
    collection = wb.find(tag("pivotCaches"))
    if collection is None:
        raise ExportError("模板没有原生数据透视表缓存，请上传已有分析xlsx。")
    return {int(c.get("cacheId")): resolve("xl/workbook.xml", relations[c.get(f"{{{REL}}}id")].get("Target")) for c in collection}


def _inspect(pkg):
    sheets = pkg.sheets()
    if not all(name in sheets for name in ("Sheet2", "Sheet5")):
        raise ExportError("模板必须包含Sheet2明细和Sheet5综合透视表。")
    if pkg.root(sheets["Sheet2"]).find(tag("tableParts")) is not None:
        raise ExportError("Sheet2已改为结构化表格，当前版本不支持自动迁移该表对象。")
    caches = _cache_parts(pkg); native = []
    sheet = pkg.root(sheets["Sheet5"])
    rels = pkg.relation_map(sheets["Sheet5"])
    # Native Excel/WPS worksheets locate pivots through worksheet relationships;
    # a pivotTableParts collection is not required (and WPS omits it).
    pivot_relations = [r for r in rels.values() if r.get("Type") == REL + "/pivotTable"]
    if len(pivot_relations) != 2:
        raise ExportError("Sheet5必须保留原主、下方两张原生透视表。")
    for relation in pivot_relations:
        part = resolve(sheets["Sheet5"], relation.get("Target")); root = pkg.root(part)
        cache_id = int(root.get("cacheId")); cache = caches.get(cache_id)
        if cache is None:
            raise ExportError("透视表指向不存在的缓存。")
        definition = pkg.root(cache); fields = list(definition.find(tag("cacheFields")))
        rows = root.find(tag("rowFields")); data = root.find(tag("dataFields"))
        if rows is None or len(rows) != 1 or data is None:
            raise ExportError("模板透视布局已变化，仅支持原单层集团/调度名称行布局。")
        if any(d.get("subtotal", "sum") != "sum" or d.get("showDataAs", "normal") != "normal" for d in data):
            raise ExportError("模板透视包含平均值或自定义值显示方式，当前版本只支持原求和/计算字段布局。")
        cols = root.find(tag("colFields"))
        if cols is not None and any(c.get("x") != "-2" for c in cols):
            raise ExportError("模板透视增加了额外列轴，当前版本仅支持原值字段列轴。")
        row_name = logical(fields[int(rows[0].get("x"))].get("name"))
        native.append({"part": part, "root": root, "cache_id": cache_id, "cache": cache, "fields": fields, "name": root.get("name"), "row_field": row_name, "data_names": [fields[int(d.get("fld"))].get("name") for d in data]})
    main = next((p for p in native if p["row_field"] == "集团"), None)
    lower = next((p for p in native if p["row_field"] == "调度名称"), None)
    if not main or not lower:
        raise ExportError("未找到原集团主透视及调度名称电量透视。")
    original = lower["fields"]
    if len(original) != 83 or any(f.get("databaseField", "1") == "0" for f in original[:67]) or any(f.get("databaseField") != "0" for f in original[67:]):
        raise ExportError("下方透视缓存不符合原67个数据字段加16个计算字段结构。")
    required = {1: "月份", 3: "调度名称", 5: "场站类型", 6: "装机容量\n（MW）", 20: "实际上网电量", 59: "总结算电量", 61: "总结算电费", 66: "集团"}
    if any(logical(original[i].get("name")) != name for i, name in required.items()):
        raise ExportError("上传模板的原始字段顺序不一致，无法安全套用。")
    supported_values = {f.get("name") for f in original} | {"HH_D", "HH_N", "HH_E", "HH_CD", "HH_Count", "HH_C"}
    if "计算-利用小时数" not in main["data_names"] or not set(main["data_names"]).issubset(supported_values) or lower["data_names"] != ["实际上网电量"]:
        raise ExportError("模板的主/下透视值字段已变化，请使用原分析模板。")
    for info in native:
        root = info["root"]
        if any(root.find(tag(n)) is not None for n in ("pageFields", "filters", "conditionalFormats")) or not global_formats_only(root):
            raise ExportError("模板透视包含暂不支持自动迁移的页筛选、值筛选或格式规则。")
    slicers = []
    for part in pkg.parts:
        if not re.fullmatch(r"xl/slicerCaches/[^/]+\.xml", part):
            continue
        root = pkg.root(part); tables = root.find(f"{{{X14}}}pivotTables")
        if tables is None:
            continue
        links = list(tables)
        linked = [p for p in native if any(t.get("name") == p["name"] for t in links)]
        if not linked:
            continue
        if len(links) != 1 or len(linked) != 1 or root.get("sourceName") not in ("月份", "场站类型", "集团") or root.find(f".//{{{X14}}}tabular") is None:
            raise ExportError("原切片器连接已变化，仅支持各透视独立的月份、类型和集团三控件。")
        slicers.append((part, root.get("sourceName"), linked[0]["name"]))
    for info in native:
        if sorted(s for _, s, n in slicers if n == info["name"]) != sorted(("月份", "场站类型", "集团")):
            raise ExportError("模板缺少两组完整的原生月份/类型/集团切片器。")
        definition = pkg.root(info["cache"])
        ext = definition.find(f".//{{{X14}}}pivotCacheDefinition")
        if ext is None or not ext.get("pivotCacheId"):
            raise ExportError("模板缺少切片器缓存标识，请上传原生分析xlsx。")
    return sheets, main, lower, slicers


def validate_template(template_bytes: bytes) -> dict:
    """Raise a user-facing ExportError before running a costly import."""
    try:
        pkg = Package(template_bytes); sheets, main, lower, slicers = _inspect(pkg)
        return {"compatible": True, "sheets": list(sheets), "main_pivot": main["name"], "lower_pivot": lower["name"], "slicer_count": len(slicers), "original_field_count": 67, "has_hours_source": "小时联动源" in sheets}
    except ExportError:
        raise
    except (KeyError, IndexError, TypeError, ValueError, AttributeError) as exc:
        raise ExportError("模板原生透视结构不完整或已改变，请上传此前生成的完整分析xlsx。") from exc


def _next_part(pkg, prefix):
    names = pkg.parts.keys() | pkg.edits.keys()
    numbers = [int(m.group(1)) for p in names if (m := re.fullmatch(re.escape(prefix) + r"(\d+)\.xml", p))]
    return prefix + str(max(numbers + [0]) + 1) + ".xml"


def _next_id(relations):
    return "rId" + str(max([int(m.group(1)) for r in relations if (m := re.fullmatch(r"rId(\d+)", r.get("Id", "")))] + [0]) + 1)


def _content_type(pkg, part, kind):
    root = pkg.root("[Content_Types].xml")
    ET.SubElement(root, f"{{{CT}}}Override", {"PartName": "/" + part, "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml." + kind + "+xml"})
    pkg.write("[Content_Types].xml", root)


def _add_sheet(pkg, name):
    wb = pkg.root("xl/workbook.xml"); rels = pkg.relationships("xl/workbook.xml")
    sheets = wb.find(tag("sheets")); part = _next_part(pkg, "xl/worksheets/sheet")
    rid = _next_id(rels)
    ET.SubElement(rels, f"{{{PKGREL}}}Relationship", {"Id": rid, "Type": REL + "/worksheet", "Target": posixpath.relpath(part, "xl")})
    ET.SubElement(sheets, tag("sheet"), {"name": name, "sheetId": str(max(int(s.get("sheetId")) for s in sheets) + 1), f"{{{REL}}}id": rid})
    pkg.write("xl/workbook.xml", wb); pkg.write(relpath("xl/workbook.xml"), rels)
    _content_type(pkg, part, "worksheet")
    return part


def _new_main_cache(pkg, main, lower):
    if main["cache"] != lower["cache"]:
        return main["cache"], main["cache_id"]
    wb = pkg.root("xl/workbook.xml"); rels = pkg.relationships("xl/workbook.xml")
    caches = wb.find(tag("pivotCaches")); cache_id = max(int(c.get("cacheId")) for c in caches) + 1
    part = _next_part(pkg, "xl/pivotCache/pivotCacheDefinition")
    records = _next_part(pkg, "xl/pivotCache/pivotCacheRecords")
    definition = pkg.root(lower["cache"])
    extids = [int(pkg.root(p).find(f".//{{{X14}}}pivotCacheDefinition").get("pivotCacheId")) for p in _cache_parts(pkg).values()]
    definition.find(f".//{{{X14}}}pivotCacheDefinition").set("pivotCacheId", str(max(extids) + 1))
    definition.set(f"{{{REL}}}id", "rId1")
    pkg.write(part, definition); pkg.write(records, ET.Element(tag("pivotCacheRecords"), {"count": "0"}))
    cr = ET.Element(f"{{{PKGREL}}}Relationships")
    ET.SubElement(cr, f"{{{PKGREL}}}Relationship", {"Id": "rId1", "Type": REL + "/pivotCacheRecords", "Target": posixpath.basename(records)})
    pkg.write(relpath(part), cr)
    rid = _next_id(rels)
    ET.SubElement(rels, f"{{{PKGREL}}}Relationship", {"Id": rid, "Type": REL + "/pivotCacheDefinition", "Target": posixpath.relpath(part, "xl")})
    ET.SubElement(caches, tag("pivotCache"), {"cacheId": str(cache_id), f"{{{REL}}}id": rid})
    pkg.write("xl/workbook.xml", wb); pkg.write(relpath("xl/workbook.xml"), rels)
    pr = pkg.relationships(main["part"])
    for relation in pr:
        if relation.get("Type") == REL + "/pivotCacheDefinition":
            relation.set("Target", posixpath.relpath(part, posixpath.dirname(main["part"])))
    pkg.write(relpath(main["part"]), pr)
    _content_type(pkg, part, "pivotCacheDefinition"); _content_type(pkg, records, "pivotCacheRecords")
    return part, cache_id


def _replace_data(root, data):
    old = root.find(tag("sheetData"))
    if old is None:
        root.append(data)
    else:
        index = list(root).index(old); root.remove(old); root.insert(index, data)


def _styles(root, row_number):
    row = root.find(f"{tag('sheetData')}/{tag('row')}[@r='{row_number}']")
    return {} if row is None else {colnum(c.get("r")): c.get("s") for c in row}


def _write_details(pkg, part, records):
    root = pkg.root(part); data = ET.Element(tag("sheetData")); styles = _styles(root, 5)
    old = root.find(tag("sheetData"))
    for row in old:
        if int(row.get("r")) <= 4:
            data.append(deepcopy(row))
    merges = root.find(tag("mergeCells"))
    if merges is not None and any(max(map(int, re.findall(r"\d+", m.get("ref")))) > 4 for m in merges):
        raise ExportError("Sheet2存在跨数据行的合并单元格，无法安全重建。")
    for number, record in enumerate(records, 5):
        values = list(record["values"][:71]); values += [None] * (71 - len(values))
        notes = record.get("notes", [])
        if notes:
            existing = str(values[70]) if values[70] is not None else ""
            items = [str(n) for n in notes] if isinstance(notes, (list, tuple)) else [str(notes)]
            additions = [n for n in dict.fromkeys(items) if n and n not in existing]
            values[70] = "；".join(s for s in [existing] + additions if s)
        energy, capacity = values[20], values[6]
        hours = energy / capacity if isinstance(energy, (int, float)) and not isinstance(energy, bool) and isinstance(capacity, (int, float)) and capacity > 0 else None
        values[67] = hours
        row = ET.SubElement(data, tag("row"), {"r": str(number)})
        for i, item in enumerate(values, 1):
            if item is not None or i == 68:
                row.append(cell(f"{col(i)}{number}", item, styles.get(i), f'IF(G{number}>0,U{number}/G{number},"")' if i == 68 else None))
    _replace_data(root, data); last = len(records) + 4
    dimension = root.find(tag("dimension"))
    if dimension is not None:
        dimension.set("ref", f"A1:BS{last}")
    auto = root.find(tag("autoFilter"))
    if auto is not None:
        auto.set("ref", f"A4:BS{last}")
        for child in list(auto):
            auto.remove(child)
    # Former conditional/data-validation ranges may refer to a previous row count.
    for child in root:
        if local(child.tag) in ("conditionalFormatting", "dataValidations"):
            for node in ([child] if local(child.tag) == "conditionalFormatting" else child):
                if "sqref" in node.attrib:
                    node.set("sqref", re.sub(r"([A-Z]+)\d+:(\$?[A-Z]+)\d+", lambda m: m.group(1) + "5:" + m.group(2) + str(last), node.get("sqref")))
    pkg.write(part, root)


def _plain_sheet(headers, rows, title, notice, number_style=None):
    root = ET.Element(tag("worksheet")); views = ET.SubElement(root, tag("sheetViews"))
    view = ET.SubElement(views, tag("sheetView"), {"workbookViewId": "0"})
    ET.SubElement(view, tag("pane"), {"ySplit": "6", "topLeftCell": "A7", "activePane": "bottomLeft", "state": "frozen"})
    columns = ET.SubElement(root, tag("cols"))
    ET.SubElement(columns, tag("col"), {"min": "1", "max": str(len(headers)), "width": "18", "customWidth": "1"})
    data = ET.SubElement(root, tag("sheetData"))
    for n, text in ((1, title), (2, notice), (3, "新增或修改明细后，请重新上传云端生成；仅在本文件内编辑不会重建完整样本及分类辅助源。")):
        ET.SubElement(data, tag("row"), {"r": str(n)}).append(cell(f"A{n}", text))
    row = ET.SubElement(data, tag("row"), {"r": "6"})
    for i, header in enumerate(headers, 1):
        row.append(cell(f"{col(i)}6", header))
    for n, values in enumerate(rows, 7):
        row = ET.SubElement(data, tag("row"), {"r": str(n)})
        for i, item in enumerate(values, 1):
            if item is not None:
                row.append(cell(f"{col(i)}{n}", item, number_style if isinstance(item, (int, float)) else None))
    ET.SubElement(root, tag("autoFilter"), {"ref": f"A6:{col(len(headers))}{len(rows)+6}"})
    return root


def _patch_pivot_cells(root, start, old_end, groups, data_names, metric, shared):
    data = root.find(tag("sheetData")); styles = _styles(root, start + 1); header_styles = _styles(root, start)
    last = start + len(groups) + 1; width = len(data_names) + 1
    for row in data:
        n = int(row.get("r"))
        if start <= n <= max(last, old_end):
            for c in list(row):
                if colnum(c.get("r")) <= width:
                    row.remove(c)
    rows = {int(r.get("r")): r for r in data}
    def put(n, values, styling):
        row = rows.get(n)
        if row is None:
            row = ET.SubElement(data, tag("row"), {"r": str(n)}); rows[n] = row
        for i, item in enumerate(values, 1):
            row.append(cell(f"{col(i)}{n}", item, styling.get(i)))
        row[:] = sorted(row, key=lambda c: colnum(c.get("r")))
    put(start, ["集团" if start == 3 else "调度名称"] + data_names, header_styles)
    for n, group in enumerate(groups, start + 1):
        put(n, ["(空白)" if group is None else group] + metric(group), styles)
    put(last, ["总计"] + metric("__all__"), styles)
    data[:] = sorted(data, key=lambda r: int(r.get("r")))


def generate_workbook(template_bytes: bytes, records: list[dict], hours_model: dict, audit_rows: list[dict], year: int = 2026) -> bytes:
    """Return a rebuilt analysis XLSX, or reject an incompatible template."""
    try:
        return _generate(template_bytes, records, hours_model, audit_rows, year)
    except ExportError:
        raise
    except (KeyError, IndexError, TypeError, ValueError, AttributeError) as exc:
        raise ExportError(f"输入数据或模板结构不兼容，未生成文件：{exc}") from exc


def _generate(template_bytes, records, hours_model, audit_rows, year):
    if not records or not isinstance(year, int) or not 2000 <= year <= 2100:
        raise ExportError("请提供有效年份和至少一条明细记录。")
    for record in records:
        if not isinstance(record.get("values"), (list, tuple)) or len(record["values"]) < 67:
            raise ExportError("每条明细必须含原A:BO共67个值。")
        for item in record["values"]:
            scalar(item)
    pkg = Package(template_bytes); sheets, main, lower, slicers = _inspect(pkg)
    original_fields = [f.get("name") for f in lower["fields"][:67]]
    old_calcs = [{"name": f.get("name"), "formula": f.get("formula"), "numFmtId": f.get("numFmtId", "0")} for f in lower["fields"][67:]]
    headers = hours_model.get("fields", hours_model.get("headers")); rows = hours_model.get("rows")
    formulas = hours_model.get("formulas", [])
    if not isinstance(headers, list) or not isinstance(rows, list) or len(rows) < len(records) or [logical(h) for h in headers[:67]] != [logical(h) for h in original_fields]:
        raise ExportError("小时模型缺少完整源或原67字段不一致。")
    if len(set(map(logical, headers))) != len(headers) or len(headers) + len(formulas) + 16 > 255:
        raise ExportError("小时源字段重复或超过WPS255字段兼容上限，请减少分类变动/样本分桶。")
    for row, record in zip(rows, records):
        if len(row) != len(headers) or [key(v) for v in row[:67]] != [key(v) for v in record["values"][:67]]:
            raise ExportError("小时辅助源与整理后的明细顺序/数值不一致。")
    if any(re.search(r"\bAND\s*\(", f["formula"], re.I) for f in formulas):
        raise ExportError("小时模型含WPS透视不兼容的动态AND公式。")
    hformula = hours_model.get("hours_formula")
    if isinstance(hformula, dict):
        hformula = hformula.get("formula")
    if not isinstance(hformula, str) or re.search(r"\bAND\s*\(", hformula, re.I):
        raise ExportError("小时模型缺少WPS兼容的原生利用小时公式。")
    calculations = deepcopy(old_calcs)
    for c in calculations:
        if c["name"] == "计算-利用小时数": c["formula"] = hformula
    all_calcs = list(formulas) + calculations
    if len({c["name"] for c in all_calcs}) != len(all_calcs) or not {"HH_D", "HH_E", "HH_CD", "HH_Count"}.issubset({c["name"] for c in all_calcs}):
        raise ExportError("小时计算字段定义重复或不完整。")
    source_name = hours_model.get("source_sheet", "小时联动源"); header_row = hours_model.get("header_row", 6)
    if source_name != "小时联动源" or header_row != 6:
        raise ExportError("当前版本要求小时联动源表、表头第6行。")
    main_cache, main_id = _new_main_cache(pkg, main, lower)
    source_part = sheets.get(source_name) or _add_sheet(pkg, source_name)
    pkg.write(source_part, _plain_sheet(headers, rows, f"{year}年小时联动源", "本次上传数据已重建。辅助数值是生成时快照，月份及分类切片器继续原生联动。"))
    _write_details(pkg, sheets["Sheet2"], records)
    physical = original_fields + headers[67:]
    main_values, main_lookups = rebuild_cache(pkg, main_cache, physical, rows, all_calcs, source_name, f"A6:{col(len(headers))}{len(rows)+6}")
    actual_rows = [list(r["values"][:67]) for r in records]
    lower_values, lower_lookups = rebuild_cache(pkg, lower["cache"], original_fields, actual_rows, old_calcs, "Sheet2", f"A4:BO{len(records)+4}")
    groups = list(dict.fromkeys(r[66] for r in actual_rows)); dispatches = list(dict.fromkeys(r[3] for r in actual_rows))
    if any(g in (None, "") for g in groups):
        raise ExportError("明细含空集团，无法生成集团对标。")
    lower_start = int(re.search(r"\d+", lower["root"].find(tag("location")).get("ref").split(":")[0]).group())
    if 3 + len(groups) + 2 >= lower_start:
        raise ExportError("集团数量导致主透视与下方透视重叠。")
    main_names = list(main["data_names"])
    if "HH_Count" not in main_names: main_names.append("HH_Count")
    all_names = physical + [c["name"] for c in all_calcs]
    pivot_rebuild(pkg, main["part"], all_names, main_values, main_lookups, "集团", groups, main_names, main_id)
    lower_names = original_fields + [c["name"] for c in old_calcs]
    pivot_rebuild(pkg, lower["part"], lower_names, lower_values, lower_lookups, "调度名称", dispatches, lower["data_names"], lower["cache_id"])
    for part, source, target in slicers:
        is_main = target == main["name"]
        cpart = main_cache if is_main else lower["cache"]
        ext_id = pkg.root(cpart).find(f".//{{{X14}}}pivotCacheDefinition").get("pivotCacheId")
        i = original_fields.index(source)
        rebuild_slicer(pkg, part, (main_values if is_main else lower_values)[i], (main_lookups if is_main else lower_lookups)[i], ext_id)
    sums = defaultdict(lambda: defaultdict(float)); dispatch_sums = defaultdict(float)
    for row in actual_rows:
        dispatch_sums[row[3]] += row[20] if isinstance(row[20], (int, float)) and not isinstance(row[20], bool) else 0
        for group in (row[66], "__all__"):
            for name, number in zip(original_fields, row):
                sums[group][name] += number if isinstance(number, (int, float)) and not isinstance(number, bool) else 0
    summary = hours_model.get("summary_by_company", {})
    def main_metrics(group):
        h = hours_model.get("summary_all", {}) if group == "__all__" else summary.get(group, {})
        if not h: raise ExportError(f"小时模型缺少集团汇总缓存：{group}")
        metrics = []
        for name in main_names:
            if name == "计算-利用小时数": metrics.append(h.get("hours"))
            elif name == "HH_Count": metrics.append(h.get("station_count", h.get("count")))
            elif name in h: metrics.append(h[name])
            elif name in original_fields: metrics.append(sums[group][name])
            else: metrics.append(arithmetic(next(c["formula"] for c in old_calcs if c["name"] == name), sums[group], {c["name"]: c["formula"] for c in old_calcs}))
        return metrics
    view = pkg.root(sheets["Sheet5"]); shared = shared_strings(pkg)
    end = lambda info: int(re.search(r"\d+$", info["root"].find(tag("location")).get("ref")).group())
    _patch_pivot_cells(view, 3, end(main), groups, [d.get("name") for d in pkg.root(main["part"]).find(tag("dataFields"))], main_metrics, shared)
    _patch_pivot_cells(view, lower_start, end(lower), dispatches, ["求和项:实际上网电量"], lambda group: [sum(dispatch_sums.values()) if group == "__all__" else dispatch_sums[group]], shared)
    data = view.find(tag("sheetData")); rowmap = {int(r.get("r")): r for r in data}
    for n, text in ((1, f"{year}年新能源交易对标（本次上传重建）"), (2, "月份/类型/集团可多选。新增或修改明细后，请重新上传云端生成；旧审计表为上传版本的历史快照。")):
        row = rowmap.get(n)
        if row is None: row = ET.SubElement(data, tag("row"), {"r": str(n)})
        for c in list(row):
            if c.get("r") == f"A{n}": row.remove(c)
        row.insert(0, cell(f"A{n}", text))
    data[:] = sorted(data, key=lambda r: int(r.get("r")))
    dimension = view.find(tag("dimension"))
    if dimension is not None: dimension.set("ref", f"A1:U{max(lower_start+len(dispatches)+1, 16)}")
    pkg.write(sheets["Sheet5"], view)
    wb = pkg.root("xl/workbook.xml")
    for defined in wb.findall(f"{tag('definedNames')}/{tag('definedName')}"):
        if defined.get("name") == "_xlnm._FilterDatabase" and "Sheet2" in (defined.text or ""):
            defined.text = f"'Sheet2'!$A$4:$BS${len(records)+4}"
    calc = wb.find(tag("calcPr"))
    if calc is not None: calc.set("fullCalcOnLoad", "1"); calc.set("forceFullCalc", "1"); calc.set("calcMode", "auto")
    views = wb.find(tag("bookViews"))
    if views is not None and len(views): views[0].set("activeTab", str(list(pkg.sheets()).index("Sheet5")))
    pkg.write("xl/workbook.xml", wb)
    existing = pkg.sheets(); audit_name = "本次更新记录"; suffix = 2
    while audit_name in existing:
        audit_name = f"本次更新记录{suffix}"; suffix += 1
    audit_part = _add_sheet(pkg, audit_name)
    audit_headers = list(dict.fromkeys(k for r in audit_rows for k in r)) or ["说明"]
    audit_values = [[_audit_value(r.get(k)) for k in audit_headers] for r in audit_rows]
    if not audit_values: audit_values = [["本次未报告逐项差异或问题。"]]
    audit = _plain_sheet(audit_headers, audit_values, "本次更新记录", f"本轮年份{year}；明细{len(records)}条。原Sheet1及汇总勾稽截至上传旧版本，本页记录本轮变更及问题。")
    _style_update_sheet(pkg, audit, audit_headers, audit_values)
    pkg.write(audit_part, audit)
    # An old calcChain may reference rows removed during a rebuild. Remove it
    # and its relationship/override; applications rebuild calculation order.
    _remove_calc_chain(pkg)
    return pkg.finish()


def _audit_value(item):
    if isinstance(item, (dict, list, tuple)):
        import json
        return json.dumps(item, ensure_ascii=False, default=str)
    return scalar(item)


def _style_update_sheet(pkg, sheet, headers, values):
    """Append styles; touch only the newly added update log's layout."""
    import unicodedata
    root = pkg.root("xl/styles.xml")
    order = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"]
    def ensure(name):
        node = root.find(tag(name))
        if node is None:
            node = ET.SubElement(root, tag(name))
            root[:] = sorted(root, key=lambda n: order.index(local(n.tag)) if local(n.tag) in order else 999)
        return node
    fonts = ensure("fonts")
    if not len(fonts):
        font = ET.SubElement(fonts, tag("font")); ET.SubElement(font, tag("sz"), {"val": "11"}); ET.SubElement(font, tag("name"), {"val": "Microsoft YaHei"})
    header_font = deepcopy(fonts[0])
    for item in list(header_font):
        if local(item.tag) in ("b", "color"): header_font.remove(item)
    ET.SubElement(header_font, tag("b")); ET.SubElement(header_font, tag("color"), {"rgb": "FFFFFFFF"})
    font_id = len(fonts); fonts.append(header_font); fonts.set("count", str(len(fonts)))
    fills = ensure("fills")
    if not len(fills):
        for kind in ("none", "gray125"): ET.SubElement(ET.SubElement(fills, tag("fill")), tag("patternFill"), {"patternType": kind})
    fill_id = len(fills); fill = ET.SubElement(ET.SubElement(fills, tag("fill")), tag("patternFill"), {"patternType": "solid"}); ET.SubElement(fill, tag("fgColor"), {"rgb": "FF17365D"}); ET.SubElement(fill, tag("bgColor"), {"indexed": "64"}); fills.set("count", str(len(fills)))
    borders = ensure("borders")
    if not len(borders): ET.SubElement(borders, tag("border"))
    borders.set("count", str(len(borders)))
    base_styles = ensure("cellStyleXfs")
    if not len(base_styles): ET.SubElement(base_styles, tag("xf"), {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0"})
    base_styles.set("count", str(len(base_styles)))
    xfs = ensure("cellXfs")
    if not len(xfs): ET.SubElement(xfs, tag("xf"), {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0"})
    style_ids = []
    for is_header in (False, True):
        style_ids.append(str(len(xfs)))
        xf = ET.SubElement(xfs, tag("xf"), {"numFmtId": "0", "fontId": str(font_id if is_header else 0), "fillId": str(fill_id if is_header else 0), "borderId": "0", "xfId": "0", "applyAlignment": "1", "applyFont": "1", "applyFill": "1"})
        ET.SubElement(xf, tag("alignment"), {"vertical": "top", "wrapText": "1"})
    xfs.set("count", str(len(xfs))); pkg.write("xl/styles.xml", root)
    widths = []
    for header in headers:
        name = str(header)
        widths.append(65 if "核对明细" in name else 56 if "说明" in name else 24 if "单元格" in name else 45 if "来源" in name else 24 if name in ("场站", "字段", "原值", "新值或计算值") else 11 if name in ("年份", "月份", "类型", "级别") else 18)
    columns = sheet.find(tag("cols")); columns.clear()
    for i, width in enumerate(widths, 1): ET.SubElement(columns, tag("col"), {"min": str(i), "max": str(i), "width": str(width), "customWidth": "1"})
    def display_units(text):
        return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in text)
    for row in sheet.find(tag("sheetData")):
        n = int(row.get("r"))
        for item in row: item.set("s", style_ids[1 if n in (1, 6) else 0])
        if n <= 3:
            row.set("ht", "32"); row.set("customHeight", "1")
        elif n == 6:
            row.set("ht", "34"); row.set("customHeight", "1")
        else:
            data = values[n - 7]
            lines = max([sum(max(1, math.ceil(display_units(line) / max(width - 2, 1))) for line in str(v if v is not None else "").split("\n")) for v, width in zip(data, widths)] + [1])
            row.set("ht", str(min(409, max(24, 15 * lines + 6)))); row.set("customHeight", "1")
    merges = ET.Element(tag("mergeCells"), {"count": "3"})
    for n in (1, 2, 3): ET.SubElement(merges, tag("mergeCell"), {"ref": f"A{n}:{col(max(2, len(headers)))}{n}"})
    # autoFilter precedes mergeCells in the worksheet content model.
    sheet.append(merges)


def _remove_calc_chain(pkg):
    # Keep the original ZIP part itself for exact package preservation; remove
    # its binding so it cannot control the new worksheet formula calculation.
    relations = pkg.relationships("xl/workbook.xml")
    for relation in list(relations):
        if relation.get("Type") == REL + "/calcChain": relations.remove(relation)
    pkg.write(relpath("xl/workbook.xml"), relations)
