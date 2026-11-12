"""Small OOXML primitives for preserving uploaded native PivotTables.

No desktop installation, business fixture, filesystem path or Excel writer is
used at runtime. Existing package parts are copied unless explicitly replaced.
"""
from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from datetime import date, datetime
from io import BytesIO
import ast
import math
import posixpath
import re
from threading import RLock
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_XML_LOCK = RLock()


class ExportError(ValueError):
    """An uploaded template or generated model is incompatible."""


def tag(name: str) -> str:
    return f"{{{NS}}}{name}"


def local(name: str) -> str:
    return name.rsplit("}", 1)[-1]


def parse(raw: bytes) -> ET.Element:
    if b"<!DOCTYPE" in raw.upper() or b"<!ENTITY" in raw.upper():
        raise ExportError("工作簿包含不支持的XML实体声明，请上传普通xlsx文件。")
    try:
        return ET.fromstring(raw)
    except ET.ParseError as exc:
        raise ExportError("工作簿XML损坏，无法安全更新。") from exc


def xml(root: ET.Element, source: bytes = b"") -> bytes:
    """Keep declarations also referenced by string-valued mc:Ignorable."""
    declarations = dict((p.decode(), u.decode()) for p, u in re.findall(
        rb'xmlns(?::([A-Za-z_][\w.-]*))?=["\']([^"\']+)["\']', source))
    namespace = root.tag.split("}")[0][1:] if root.tag.startswith("{") else ""
    with _XML_LOCK:
        for prefix, uri in declarations.items():
            if prefix and not re.fullmatch(r"ns\d+", prefix):
                ET.register_namespace(prefix, uri)
        ET.register_namespace("r", REL)
        if namespace:
            ET.register_namespace("", namespace)
        data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    start = data.index(b"<", data.index(b"?>") + 2)
    end = data.index(b">", start)
    opening = data[start:end]
    additions = []
    for prefix, uri in declarations.items():
        attr = ("xmlns" + (":" + prefix if prefix else "")).encode()
        if not re.search(rb"\b" + re.escape(attr) + rb"=", opening):
            additions.append(b" " + attr + b'="' + uri.encode() + b'"')
    return data[:end] + b"".join(additions) + data[end:]


def col(number: int) -> str:
    result = ""
    while number:
        number, rem = divmod(number - 1, 26)
        result = chr(65 + rem) + result
    return result


def colnum(ref: str) -> int:
    value = 0
    for ch in re.match(r"[A-Z]+", ref).group():
        value = value * 26 + ord(ch) - 64
    return value


def logical(text: str) -> str:
    return re.sub(r"_x([0-9a-fA-F]{4})_", lambda m: chr(int(m.group(1), 16)), text)


def resolve(part: str, target: str) -> str:
    value = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
    if value.startswith("../") or ":" in value:
        raise ExportError("工作簿含不支持的外部部件路径。")
    return value


def relpath(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


class Package:
    def __init__(self, raw: bytes):
        try:
            self.zip = ZipFile(BytesIO(raw))
            entries = self.zip.infolist()
            if len(entries) > 3000 or sum(e.file_size for e in entries) > 350_000_000:
                raise ExportError("模板过大，请使用不超过350MB解压大小的普通分析工作簿。")
            if len({e.filename for e in entries}) != len(entries):
                raise ExportError("工作簿包含重复ZIP部件，无法安全更新。")
            self.parts = {e.filename: self.zip.read(e.filename) for e in entries}
        except (BadZipFile, KeyError, RuntimeError) as exc:
            raise ExportError("无法读取上传文件，请上传未加密的xlsx工作簿。") from exc
        if "xl/workbook.xml" not in self.parts or "[Content_Types].xml" not in self.parts:
            raise ExportError("上传文件不是支持的xlsx工作簿。")
        if any("vbaProject" in name or name.startswith("_xmlsignatures/") for name in self.parts):
            raise ExportError("暂不更新带宏或数字签名的模板，请另存为普通xlsx后上传。")
        self.edits: dict[str, bytes] = {}

    def read(self, name: str) -> bytes:
        if name in self.edits:
            return self.edits[name]
        try:
            return self.parts[name]
        except KeyError as exc:
            raise ExportError(f"模板缺少必要部件：{name}") from exc

    def root(self, name: str) -> ET.Element:
        return parse(self.read(name))

    def write(self, name: str, root: ET.Element):
        self.edits[name] = xml(root, self.parts.get(name, b""))

    def relationships(self, part: str):
        name = relpath(part)
        if name not in self.parts and name not in self.edits:
            return ET.Element(f"{{{PKGREL}}}Relationships")
        return self.root(name)

    def relation_map(self, part: str):
        return {r.get("Id"): r for r in self.relationships(part)}

    def sheets(self):
        relations = self.relation_map("xl/workbook.xml")
        result = {}
        for sheet in self.root("xl/workbook.xml").find(tag("sheets")):
            relation = relations.get(sheet.get(f"{{{REL}}}id"))
            if relation is None or relation.get("TargetMode") == "External":
                raise ExportError("模板工作表关系不完整。")
            result[sheet.get("name")] = resolve("xl/workbook.xml", relation.get("Target"))
        return result

    def finish(self) -> bytes:
        # Parse changed XML and check every internal relationship target.
        allparts = {**self.parts, **self.edits}
        for name, data in self.edits.items():
            if name.endswith((".xml", ".rels")):
                if len(data) < 1_000_000:
                    parse(data)
                else:
                    if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
                        raise ExportError("生成结果含不支持的XML实体声明。")
                    try:
                        for _, element in ET.iterparse(BytesIO(data), events=("end",)):
                            element.clear()
                    except ET.ParseError as exc:
                        raise ExportError("生成的XML部件不完整，已停止导出。") from exc
        for name, data in allparts.items():
            if not name.endswith(".rels"):
                continue
            owner = name.replace("/_rels/", "/")[:-5] if name != "_rels/.rels" else ""
            for relation in parse(data):
                if relation.get("TargetMode") == "External":
                    continue
                target = resolve(owner, relation.get("Target"))
                if target not in allparts:
                    raise ExportError(f"模板存在断开的内部关系：{target}")
        buffer = BytesIO()
        with ZipFile(buffer, "w", ZIP_DEFLATED) as result:
            for item in self.zip.infolist():
                result.writestr(deepcopy(item), self.edits.get(item.filename, self.parts[item.filename]))
            for name in self.edits.keys() - self.parts.keys():
                result.writestr(name, self.edits[name])
        return buffer.getvalue()


def shared_strings(pkg: Package):
    if "xl/sharedStrings.xml" not in pkg.parts:
        return []
    return ["".join(x.itertext()) for x in pkg.root("xl/sharedStrings.xml")]


def value(cell, shared):
    if cell is None:
        return None
    typ = cell.get("t"); v = cell.find(tag("v"))
    if typ == "inlineStr":
        return "".join(cell.find(tag("is")).itertext())
    if v is None or v.text is None:
        return None
    if typ == "s":
        return shared[int(v.text)]
    if typ in ("str", "d", "e"):
        return v.text
    if typ == "b":
        return v.text == "1"
    return float(v.text)


def scalar(item):
    if item is None or item == "":
        return None
    if isinstance(item, (datetime, date)):
        return item.isoformat()
    if isinstance(item, bool):
        return item
    if isinstance(item, (int, float)):
        if not math.isfinite(item):
            raise ExportError("数据包含非有限数值，不能导出。")
        return item
    if not isinstance(item, str):
        raise ExportError(f"数据包含不支持的单元格类型：{type(item).__name__}")
    # XML 1.0 controls are not legal. Refuse rather than silently alter a value.
    if re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", item):
        raise ExportError("数据含不可写入Excel的控制字符。")
    return item


def cell(ref: str, item, style: str | None = None, formula: str | None = None):
    attrs = {"r": ref}
    if style is not None:
        attrs["s"] = str(style)
    result = ET.Element(tag("c"), attrs)
    if formula:
        ET.SubElement(result, tag("f")).text = formula.lstrip("=")
    item = scalar(item)
    if item is None:
        return result
    if isinstance(item, bool):
        result.set("t", "b"); ET.SubElement(result, tag("v")).text = "1" if item else "0"
    elif isinstance(item, (int, float)):
        ET.SubElement(result, tag("v")).text = format(item, ".16g")
    else:
        result.set("t", "inlineStr")
        text = ET.SubElement(ET.SubElement(result, tag("is")), tag("t"))
        if item != item.strip() or "\n" in item:
            text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text.text = item
    return result


def key(item):
    item = scalar(item)
    if item is None:
        return ("m", None)
    if isinstance(item, bool):
        return ("b", "1" if item else "0")
    if isinstance(item, (int, float)):
        return ("n", format(item, ".16g"))
    return ("s", item)


def cache_fields(headers, rows, calculations):
    """Return definition fields, compact indexed records and item lookups."""
    values = [[] for _ in headers]; lookups = [{} for _ in headers]
    for row in rows:
        if len(row) != len(headers):
            raise ExportError("辅助源行宽与表头不一致。")
        for i, item in enumerate(row):
            k = key(item)
            if k not in lookups[i]:
                lookups[i][k] = len(values[i]); values[i].append(k)
    fields = ET.Element(tag("cacheFields"), {"count": str(len(headers) + len(calculations))})
    for name, items in zip(headers, values):
        field = ET.SubElement(fields, tag("cacheField"), {"name": name, "numFmtId": "0"})
        kinds = {k for k, v in items}
        attrs = {"count": str(len(items)), "containsString": str(int("s" in kinds)), "containsBlank": str(int("m" in kinds)), "containsNumber": str(int("n" in kinds)), "containsDate": "0", "containsNonDate": "1"}
        numbers = [float(v) for k, v in items if k == "n"]
        if numbers:
            attrs.update(minValue=format(min(numbers), ".16g"), maxValue=format(max(numbers), ".16g"))
        si = ET.SubElement(field, tag("sharedItems"), attrs)
        for kind, item in items:
            ET.SubElement(si, tag(kind), {} if kind == "m" else {"v": item})
    for calculation in calculations:
        ET.SubElement(fields, tag("cacheField"), {"name": calculation["name"], "formula": calculation["formula"].lstrip("="), "databaseField": "0", "numFmtId": str(calculation.get("numFmtId", 0))})
    # A wide cache has millions of <x> nodes. Stream its already-validated
    # integer indexes instead of allocating an Element object per value.
    records = BytesIO()
    records.write((f'<?xml version="1.0" encoding="utf-8"?><pivotCacheRecords xmlns="{NS}" count="{len(rows)}">').encode())
    for row in rows:
        indexes = (lookups[i][key(item)] for i, item in enumerate(row))
        records.write(('<r>' + ''.join(f'<x v="{index}"/>' for index in indexes) + '</r>').encode())
    records.write(b'</pivotCacheRecords>')
    return fields, records.getvalue(), values, lookups


def rebuild_cache(pkg, part, headers, rows, calculations, source_sheet, source_ref):
    root = pkg.root(part)
    source = root.find(tag("cacheSource"))
    if source is None or source.get("type") != "worksheet":
        raise ExportError("只支持以本工作簿工作表为源的普通透视表。")
    source.clear(); source.set("type", "worksheet")
    ET.SubElement(source, tag("worksheetSource"), {"sheet": source_sheet, "ref": source_ref})
    fields, records, values, lookups = cache_fields(headers, rows, calculations)
    old = root.find(tag("cacheFields")); index = list(root).index(old)
    root.remove(old); root.insert(index, fields)
    root.set("recordCount", str(len(rows))); root.set("refreshOnLoad", "1"); root.set("enableRefresh", "1"); root.set("saveData", "1")
    for attr in ("refreshedDate", "refreshedBy"):
        root.attrib.pop(attr, None)
    records_rel = pkg.relation_map(part).get(root.get(f"{{{REL}}}id"))
    if records_rel is None:
        raise ExportError("模板透视缓存缺少记录关系。")
    records_part = resolve(part, records_rel.get("Target"))
    pkg.write(part, root); pkg.edits[records_part] = records
    return values, lookups


def pivot_rebuild(pkg, part, names, values, lookups, group_field, groups, data_names, cache_id):
    """Keep object identity/style while replacing all field and item indexes."""
    root = pkg.root(part); old = root.find(tag("pivotFields"))
    root.set("cacheId", str(cache_id)); root.set("showError", "1"); root.set("errorCaption", "无完整样本")
    fields = ET.Element(tag("pivotFields"), {"count": str(len(names))})
    ids = {logical(name): i for i, name in enumerate(names)}
    group_index = ids[logical(group_field)]
    data = root.find(tag("dataFields"))
    captions = [x.get("name") for x in data]
    formats = [x.get("numFmtId") for x in data]
    for i, name in enumerate(names):
        attrs = {"compact": "0", "showAll": "0", "defaultSubtotal": "0"}
        if i == group_index:
            attrs["axis"] = "axisRow"
        if logical(name) in {logical(n) for n in data_names}:
            attrs["dataField"] = "1"
        if i >= len(values):
            attrs.update(dragToCol="0", dragToPage="0", dragToRow="0")
        field = ET.SubElement(fields, tag("pivotField"), attrs)
        if i < len(values):
            order = [lookups[i][key(g)] for g in groups] if i == group_index else list(range(len(values[i])))
            items = ET.SubElement(field, tag("items"), {"count": str(len(order))})
            for j in order:
                ET.SubElement(items, tag("item"), {"x": str(j)})
    position = list(root).index(old); root.remove(old); root.insert(position, fields)
    rowfields = root.find(tag("rowFields")); rowfields.clear(); rowfields.set("count", "1")
    ET.SubElement(rowfields, tag("field"), {"x": str(group_index)})
    rowitems = root.find(tag("rowItems")); rowitems.clear(); rowitems.set("count", str(len(groups) + 1))
    for i in range(len(groups)):
        ET.SubElement(ET.SubElement(rowitems, tag("i")), tag("x"), {"v": str(i)})
    ET.SubElement(ET.SubElement(rowitems, tag("i"), {"t": "grand"}), tag("x"), {"v": "0"})
    data.clear(); data.set("count", str(len(data_names)))
    for i, name in enumerate(data_names):
        attrs = {"name": captions[i] if i < len(captions) else "完整样本站数", "fld": str(ids[logical(name)]), "subtotal": "sum", "baseField": "0", "baseItem": "0"}
        if i < len(formats) and formats[i] is not None:
            attrs["numFmtId"] = formats[i]
        ET.SubElement(data, tag("dataField"), attrs)
    cols = root.find(tag("colItems"))
    if cols is not None:
        cols.clear(); cols.set("count", str(len(data_names)))
        for i in range(len(data_names)):
            ET.SubElement(ET.SubElement(cols, tag("i"), {"i": str(i)}), tag("x"), {"v": str(i)})
    # Only this fixed one-row-field/no-page-filter layout is supported.
    for unsupported in ("pageFields", "filters", "conditionalFormats"):
        node = root.find(tag(unsupported))
        if node is not None:
            raise ExportError(f"模板透视表包含暂不支持重映射的设置：{unsupported}")
    if not global_formats_only(root):
        raise ExportError("模板含按项目/字段绑定的透视格式，暂不支持自动重映射。")
    location = root.find(tag("location")); start = location.get("ref").split(":")[0]
    startrow = int(re.search(r"\d+", start).group())
    location.set("ref", f"A{startrow}:{col(len(data_names)+1)}{startrow+len(groups)+1}")
    pkg.write(part, root)
    return startrow


def rebuild_slicer(pkg, part, values, lookup, extension_id):
    root = pkg.root(part); tab = root.find(f".//{{{X14}}}tabular")
    if tab is None:
        raise ExportError("模板含非普通表格切片器。")
    tab.set("pivotCacheId", str(extension_id))
    items = tab.find(f"{{{X14}}}items")
    if items is not None:
        tab.remove(items)
    items = ET.Element(f"{{{X14}}}items", {"count": str(len(values))})
    for k in values:
        ET.SubElement(items, f"{{{X14}}}i", {"x": str(lookup[k]), "s": "1"})
    tab.insert(0, items); pkg.write(part, root)


def global_formats_only(root):
    formats = root.find(tag("formats"))
    if formats is None:
        return True
    return all(local(n.tag) == "format" and len(n) == 1 and local(n[0].tag) == "pivotArea"
               and n[0].get("type") == "all" and len(n[0]) == 0 and "field" not in n[0].attrib
               for n in formats)


def arithmetic(formula, sums, calculations, stack=()):
    """Evaluate only arithmetic field references for initial pivot cached cells."""
    if formula in stack:
        raise ExportError("模板计算字段存在循环引用。")
    fields = {**{k: v for k, v in sums.items()}, **{k: None for k in calculations}}
    names = sorted(fields, key=len, reverse=True)
    translated = formula
    substitutions = {}
    for i, name in enumerate(names):
        alias = f"v_{i}"
        pattern = re.escape(name)
        quoted = "'" + name.replace("'", "''") + "'"
        if quoted in translated or name in translated:
            translated = translated.replace(quoted, alias)
            translated = re.sub(r"(?<![\w])" + pattern + r"(?![\w])", alias, translated)
            substitutions[alias] = name
    try:
        node = ast.parse(translated.lstrip("="), mode="eval").body
    except SyntaxError as exc:
        raise ExportError("模板含暂不支持预计算的自定义交易计算字段。") from exc
    def visit(n):
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)):
            return n.value
        if isinstance(n, ast.Name) and n.id in substitutions:
            name = substitutions[n.id]
            if name in sums:
                return sums[name]
            return arithmetic(calculations[name], sums, calculations, stack + (formula,))
        if isinstance(n, ast.UnaryOp) and isinstance(n.op, (ast.UAdd, ast.USub)):
            return visit(n.operand) * (-1 if isinstance(n.op, ast.USub) else 1)
        if isinstance(n, ast.BinOp) and isinstance(n.op, (ast.Add, ast.Sub, ast.Mult, ast.Div)):
            a, b = visit(n.left), visit(n.right)
            if a is None or b is None:
                return None
            if isinstance(n.op, ast.Add): return a + b
            if isinstance(n.op, ast.Sub): return a - b
            if isinstance(n.op, ast.Mult): return a * b
            return a / b
        raise ExportError("模板含暂不支持的自定义交易公式；只支持原模板四则计算字段。")
    try:
        return visit(node)
    except ZeroDivisionError:
        return None
