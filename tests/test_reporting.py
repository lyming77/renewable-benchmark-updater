from benchmark.reporting import audit_row


def test_quantity_difference_is_explained_with_both_numbers_and_gap():
    report = audit_row({"company": "测试公司", "month": 3, "station": "测试站",
                        "severity": "warning", "message": "总量组成不一致",
                        "fields": ["总结算电量", "实际上网电量"],
                        "reported": 101, "known_sum": 100, "difference": 1}, "待核对")
    assert report["原值"] == 101
    assert report["新值或计算值"] == 100
    assert report["差额"] == 1
    assert "已知组成项合计=100" in report["核对明细"]
    assert report["场站"] == "测试站"


def test_added_record_report_keeps_identity_without_dumping_private_structure():
    values = [None] * 67
    values[2] = "新增测试站"
    report = audit_row({"kind": "added_record", "after": {
        "values": values, "company": "测试公司", "month": 10, "year": 2026,
        "source": "合成汇总.xlsx"}}, "变更")
    assert report["类型"] == "新增记录"
    assert report["场站"] == "新增测试站"
    assert report["集团"] == "测试公司"
    assert report["月份"] == 10
    assert "详见Sheet2" in report["新值或计算值"]
