"""Synthetic ZIP fixtures only: no uploaded or private business data."""
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as E
import unittest

from benchmark.export import ExportError, generate_workbook, validate_template
from benchmark.hours import build_hours_model
from benchmark.native import CT, NS, PKGREL, REL, X14, Package, cell, col, parse, shared_strings, tag, value, xml


def synthetic_headers():
    headers = [f"指标{i}" for i in range(67)]
    for i, name in {0:"来源位置",1:"月份",2:"结算单元\n名称",3:"调度名称",4:"所在地区",5:"场站类型",6:"装机容量\n（MW）",7:"投产时间",20:"实际上网电量",22:"实时电能量电费",25:"日前电能量电费",28:"中长期合约电费",10:"省间日前结算电费",13:"省间中长期结算电费",59:"总结算电量",61:"总结算电费",66:"集团"}.items(): headers[i]=name
    return headers


def synthetic_records():
    result=[]
    for month, energy, capacity in [(1,1000,100),(3,3000,200)]:
        row=[0]*67
        for i,item in {0:"合成来源",1:f"{month}月份",2:"合成站",3:"合成调度",4:"合成地区",5:"陆风",6:capacity,7:None,20:energy,22:energy*200,59:energy,61:energy*200,66:"合成集团"}.items():row[i]=item
        result.append({"values":row,"notes":["合成备注"]})
    return result


def synthetic_template():
    headers=synthetic_headers()
    names=["计算-实时电价","原计算1","计算-单位电量中长期收益","计算-单位电量日前收益","原计算4","原计算5","原计算6","原计算7","计算-单位电量省间收益","原计算9","原计算10","原计算11","原计算12","原计算13","计算-利用小时数","计算-单位电量交易电价"]
    formulas=["实时电能量电费/总结算电量","1","中长期合约电费/总结算电量","日前电能量电费/总结算电量","1","1","1","1","(省间日前结算电费+省间中长期结算电费)/总结算电量","1","1","1","1","1","实际上网电量/'装机容量_x000a_（MW）'","'计算-实时电价'+'计算-单位电量中长期收益'+'计算-单位电量日前收益'+'计算-单位电量省间收益'"]
    parts={};wb=E.Element(tag("workbook"));views=E.SubElement(wb,tag("bookViews"));E.SubElement(views,tag("workbookView"),{"activeTab":"0"});sheets=E.SubElement(wb,tag("sheets"))
    rels=E.Element(f"{{{PKGREL}}}Relationships")
    for i,name in enumerate(["Sheet5","Sheet2","Sheet1","汇总勾稽"],1):
        E.SubElement(sheets,tag("sheet"),{"name":name,"sheetId":str(i),f"{{{REL}}}id":f"rId{i}"})
        E.SubElement(rels,f"{{{PKGREL}}}Relationship",{"Id":f"rId{i}","Type":REL+"/worksheet","Target":f"worksheets/sheet{i}.xml"})
        root=E.Element(tag("worksheet"));E.SubElement(root,tag("dimension"),{"ref":"A1:BS60"});data=E.SubElement(root,tag("sheetData"))
        if name=="Sheet2":
            row=E.SubElement(data,tag("row"),{"r":"4"})
            for c,h in enumerate(headers+["利用小时","存量增量","退补调整","核对说明"],1):row.append(cell(f"{col(c)}4",h))
            E.SubElement(root,tag("autoFilter"),{"ref":"A4:BS5"})
        elif name!="Sheet5":E.SubElement(data,tag("row"),{"r":"1"}).append(cell("A1","历史合成快照"))
        parts[f"xl/worksheets/sheet{i}.xml"]=xml(root)
    caches=E.SubElement(wb,tag("pivotCaches"));E.SubElement(caches,tag("pivotCache"),{"cacheId":"0",f"{{{REL}}}id":"rId5"});E.SubElement(wb,tag("calcPr"),{"calcMode":"auto"})
    E.SubElement(rels,f"{{{PKGREL}}}Relationship",{"Id":"rId5","Type":REL+"/pivotCacheDefinition","Target":"pivotCache/pivotCacheDefinition1.xml"})
    parts["xl/workbook.xml"]=xml(wb);parts["xl/_rels/workbook.xml.rels"]=xml(rels)
    root=E.Element(tag("pivotCacheDefinition"),{f"{{{REL}}}id":"rId1","recordCount":"0"});source=E.SubElement(root,tag("cacheSource"),{"type":"worksheet"});E.SubElement(source,tag("worksheetSource"),{"sheet":"Sheet2","ref":"A4:BO5"});fields=E.SubElement(root,tag("cacheFields"),{"count":"83"})
    for h in headers:E.SubElement(E.SubElement(fields,tag("cacheField"),{"name":h.replace("\n","_x000a_")}),tag("sharedItems"),{"count":"0"})
    for n,f in zip(names,formulas):E.SubElement(fields,tag("cacheField"),{"name":n,"databaseField":"0","formula":f})
    ext=E.SubElement(E.SubElement(root,tag("extLst")),tag("ext"),{"uri":"{725AE2AE-9491-48be-B2B4-4EB974FC3084}"});E.SubElement(ext,f"{{{X14}}}pivotCacheDefinition",{"pivotCacheId":"1"})
    parts["xl/pivotCache/pivotCacheDefinition1.xml"]=xml(root);parts["xl/pivotCache/pivotCacheRecords1.xml"]=xml(E.Element(tag("pivotCacheRecords"),{"count":"0"}))
    cr=E.Element(f"{{{PKGREL}}}Relationships");E.SubElement(cr,f"{{{PKGREL}}}Relationship",{"Id":"rId1","Type":REL+"/pivotCacheRecords","Target":"pivotCacheRecords1.xml"});parts["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"]=xml(cr)
    sheetrels=E.Element(f"{{{PKGREL}}}Relationships")
    for number,rowfield,indices,start in [(1,66,[67,69,70,75,82,81],3),(2,3,[20],57)]:
        pivot=E.Element(tag("pivotTableDefinition"),{"name":f"数据透视表{3-number}","cacheId":"0","dataCaption":"值"});E.SubElement(pivot,tag("location"),{"ref":f"A{start}:{'G' if number==1 else 'B'}{start+2}","firstHeaderRow":"0","firstDataRow":"1","firstDataCol":"1"});pf=E.SubElement(pivot,tag("pivotFields"),{"count":"83"})
        for i in range(83):E.SubElement(pf,tag("pivotField"),{})
        E.SubElement(E.SubElement(pivot,tag("rowFields"),{"count":"1"}),tag("field"),{"x":str(rowfield)});E.SubElement(pivot,tag("rowItems"),{"count":"0"})
        if number==1:E.SubElement(E.SubElement(pivot,tag("colFields"),{"count":"1"}),tag("field"),{"x":"-2"});E.SubElement(pivot,tag("colItems"),{"count":"0"})
        data=E.SubElement(pivot,tag("dataFields"),{"count":str(len(indices))})
        for i in indices:E.SubElement(data,tag("dataField"),{"fld":str(i),"name":names[i-67] if i>=67 else headers[i]})
        parts[f"xl/pivotTables/pivotTable{number}.xml"]=xml(pivot)
        pr=E.Element(f"{{{PKGREL}}}Relationships");E.SubElement(pr,f"{{{PKGREL}}}Relationship",{"Id":"rId1","Type":REL+"/pivotCacheDefinition","Target":"../pivotCache/pivotCacheDefinition1.xml"});parts[f"xl/pivotTables/_rels/pivotTable{number}.xml.rels"]=xml(pr)
        E.SubElement(sheetrels,f"{{{PKGREL}}}Relationship",{"Id":f"rId{number}","Type":REL+"/pivotTable","Target":f"../pivotTables/pivotTable{number}.xml"})
        for j,field in enumerate(["月份","场站类型","集团"],1):
            sc=E.Element(f"{{{X14}}}slicerCacheDefinition",{"name":f"合成切片{number}{j}","sourceName":field});E.SubElement(E.SubElement(sc,f"{{{X14}}}pivotTables"),f"{{{X14}}}pivotTable",{"name":f"数据透视表{3-number}","tabId":"1"});E.SubElement(E.SubElement(sc,f"{{{X14}}}data"),f"{{{X14}}}tabular",{"pivotCacheId":"1"});parts[f"xl/slicerCaches/slicerCache{(number-1)*3+j}.xml"]=xml(sc)
    parts["xl/worksheets/_rels/sheet1.xml.rels"]=xml(sheetrels)
    parts["[Content_Types].xml"]=xml(E.Element(f"{{{CT}}}Types"));parts["xl/styles.xml"]=xml(E.Element(tag("styleSheet")))
    parts["opaque-preserved.bin"]=b"synthetic native attachment bytes"
    out=BytesIO()
    with ZipFile(out,"w",ZIP_DEFLATED) as z:
        for name,data in parts.items():z.writestr(name,data)
    return out.getvalue()


class ExportTests(unittest.TestCase):
    def test_rebuilds_both_caches_and_preserves_history(self):
        template=synthetic_template();records=synthetic_records();model=build_hours_model(records,synthetic_headers())
        self.assertTrue(validate_template(template)["compatible"])
        result=generate_workbook(template,records,model,[{"action":"新增","detail":"合成记录"}])
        pkg=Package(result);sheets=pkg.sheets()
        self.assertIn("小时联动源",sheets);self.assertIn("本次更新记录",sheets)
        with ZipFile(BytesIO(template)) as old,ZipFile(BytesIO(result)) as new:
            for p in ["xl/worksheets/sheet3.xml","xl/worksheets/sheet4.xml","opaque-preserved.bin"]:self.assertEqual(old.read(p),new.read(p))
        wb=pkg.root("xl/workbook.xml");self.assertEqual(len(wb.find(tag("pivotCaches"))),2)
        self.assertEqual(pkg.root("xl/pivotCache/pivotCacheDefinition1.xml").get("recordCount"),"2")
        self.assertEqual(pkg.root("xl/pivotCache/pivotCacheDefinition2.xml").get("recordCount"),str(len(model["rows"])))
        self.assertEqual(pkg.root("xl/pivotTables/pivotTable1.xml").get("cacheId"),"1")
        self.assertEqual(pkg.root("xl/pivotTables/pivotTable2.xml").get("cacheId"),"0")
        shared=shared_strings(pkg);view=pkg.root(sheets["Sheet5"]);cells={c.get("r"):c for row in view.find(tag("sheetData")) for c in row}
        self.assertAlmostEqual(value(cells["G4"],shared),26.6666666666667)
        self.assertEqual(value(cells["H4"],shared),1)
        self.assertEqual(value(cells["B58"],shared),4000)
        self.assertNotIn(b"AND(",pkg.read("xl/pivotCache/pivotCacheDefinition2.xml"))

    def test_rejects_plain_or_broken_upload(self):
        with self.assertRaises(ExportError):validate_template(b"not a zip")
        pkg=Package(synthetic_template());root=pkg.root("xl/pivotTables/pivotTable2.xml");root.find(tag("rowFields"))[0].set("x","66");pkg.write("xl/pivotTables/pivotTable2.xml",root)
        with self.assertRaises(ExportError):validate_template(pkg.finish())

    def test_rejects_misaligned_hours_without_altering_input(self):
        records=synthetic_records();model=build_hours_model(records,synthetic_headers());model["rows"][0][20]=99
        template=synthetic_template()
        with self.assertRaises(ExportError):generate_workbook(template,records,model,[])
        self.assertTrue(validate_template(template)["compatible"])

    def test_preserves_an_added_raw_value_and_global_pivot_format(self):
        pkg=Package(synthetic_template());main=pkg.root("xl/pivotTables/pivotTable1.xml");data=main.find(tag("dataFields"));data.insert(0,E.Element(tag("dataField"),{"name":"新增日前费","fld":"25"}));data.set("count",str(len(data)));pkg.write("xl/pivotTables/pivotTable1.xml",main)
        lower=pkg.root("xl/pivotTables/pivotTable2.xml");formats=E.SubElement(lower,tag("formats"),{"count":"1"});E.SubElement(E.SubElement(formats,tag("format"),{"dxfId":"0"}),tag("pivotArea"),{"type":"all","fieldPosition":"0"});pkg.write("xl/pivotTables/pivotTable2.xml",lower)
        records=synthetic_records();model=build_hours_model(records,synthetic_headers());result=Package(generate_workbook(pkg.finish(),records,model,[]));view=result.root(result.sheets()["Sheet5"]);cells={c.get("r"):c for r in view.find(tag("sheetData")) for c in r};shared=shared_strings(result)
        self.assertAlmostEqual(value(cells["H4"],shared),26.6666666666667);self.assertEqual(value(cells["I4"],shared),1)
        self.assertEqual(result.root("xl/pivotTables/pivotTable2.xml").find(tag("formats"))[0][0].get("type"),"all")


if __name__=="__main__":unittest.main()
