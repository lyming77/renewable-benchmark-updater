"""Streamlit entry point: upload two workbook versions and download an update."""
from __future__ import annotations

import hashlib
import json
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from benchmark.service import preview_summary, update_workbook
from benchmark.reporting import audit_row


def readable_rows(rows):
    """Audit old/new values can mix numbers and text in the same column."""
    return [{key: "" if value is None else
             json.dumps(value, ensure_ascii=False, default=str) if isinstance(value, (list, dict, tuple)) else str(value)
             for key, value in row.items()} for row in rows]

st.set_page_config(page_title="新能源交易对标更新", page_icon="📊", layout="wide")
st.title("新能源交易对标更新")
st.write("上传上一版分析表与最新收集表，交叉补全数据，核对量价费，并重新生成带筛选联动的分析工作簿。")

left, right = st.columns(2, gap="large")
with left:
    previous = st.file_uploader("上一版分析表", type=["xlsx"], key="previous",
                                help="使用此前带有 Sheet2 明细、Sheet5 原透视表的分析工作簿。")
with right:
    latest = st.file_uploader("最新收集汇总表", type=["xlsx"], accept_multiple_files=True, key="latest",
                              help="可同时上传多个公司或月份的 .xlsx 收集表；按字段名称识别不同列数。")

setting_a, setting_b = st.columns([1, 3])
with setting_a:
    year = int(st.number_input("统计年份", 2000, 2100, 2026, step=1))
with setting_b:
    policy = st.radio("已存在数值发生冲突时", ["保留旧值，仅补空缺", "采用新值，记录覆盖明细"],
                      horizontal=True, key="conflict_policy")
    st.caption("同一批新文件互相冲突、场站身份无法唯一匹配时，会列入待核对。真实零值不会当作缺失。")

st.caption("上传文件仅用于本次会话处理，不写入 GitHub。单文件上限 50 MB，最多 15 份收集表。")

fingerprint = None
if previous is not None and latest:
    digest = hashlib.sha256(previous.getvalue())
    digest.update(previous.name.encode("utf-8"))
    for item in latest:
        digest.update(item.name.encode("utf-8")); digest.update(item.getvalue())
    digest.update(f"{year}|{policy}".encode("utf-8"))
    fingerprint = digest.hexdigest()

if st.session_state.get("result_fingerprint") != fingerprint:
    st.session_state.pop("processing_result", None)

if st.button("核对并生成新版分析表", type="primary", disabled=not (previous is not None and latest), width="stretch"):
    st.session_state.pop("processing_result", None)
    with st.status("正在处理上传文件", expanded=True) as status:
        stage = st.empty()
        try:
            result = update_workbook(
                previous.getvalue(), [(item.name, item.getvalue()) for item in latest],
                previous_name=previous.name, year=year,
                conflict_policy="keep_old" if policy.startswith("保留") else "prefer_new",
                progress=stage.write,
            )
        except (ValueError, KeyError) as exc:
            status.update(label="暂时无法生成", state="error", expanded=True)
            st.error(str(exc))
        except Exception:
            status.update(label="生成未完成", state="error", expanded=True)
            st.error("工作簿处理失败。本次没有生成可下载结果，请检查文件格式、原透视表结构及数据规模。")
        else:
            st.session_state.processing_result = result
            st.session_state.result_fingerprint = fingerprint
            status.update(label="新版分析表已生成", state="complete", expanded=False)

result = st.session_state.get("processing_result")
if result is not None:
    a, b, c, d = st.columns(4)
    a.metric("原分析记录", f"{result.stats['old_records']:,}")
    b.metric("更新后记录", f"{result.stats['output_records']:,}")
    c.metric("本次变更", f"{len(result.changes):,}")
    pending = [row for row in result.issues if row.get("severity") != "info"]
    d.metric("待核对事项", f"{len(pending):,}")
    now = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y%m%d_%H%M")
    st.download_button("下载新版分析表", result.workbook,
                       file_name=f"{year}年新能源交易对标_更新_{now}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", width="stretch")
    st.info("新版保留原透视表和切片器。利用小时仅纳入所选各月数据完整的场站，交易指标保留全部已收录记录。生成后若再增改源数据，请重新上传生成。")

    preview, changes, issues = st.tabs(["同一筛选下的统计预览", "本次变更", "待核对事项"])
    with preview:
        months = sorted({r.get("month") for r in result.records if r.get("month")})
        companies = sorted({r["values"][66] for r in result.records if r["values"][66]})
        types = sorted({r["values"][5] for r in result.records if r["values"][5]})
        m, co, ty = st.columns(3)
        selected_months = m.multiselect("月份（可任意多选）", months, default=months,
                                       format_func=lambda value: f"{value}月", key=f"months_{fingerprint}")
        selected_companies = co.multiselect("集团", companies, default=companies, key=f"companies_{fingerprint}")
        selected_types = ty.multiselect("场站类型", types, default=types, key=f"types_{fingerprint}")
        if selected_months and selected_companies and selected_types:
            try:
                table = preview_summary(result.records, selected_months, selected_companies, selected_types, year)
                st.dataframe(pd.DataFrame(table), hide_index=True, width="stretch")
            except ValueError as exc:
                st.warning(str(exc))
        else:
            st.caption("请至少选择一个月份、集团和场站类型。")
        st.caption("容量按所选月份日历天数计算期间平均值。预览筛选不修改下载文件的默认全选状态。")
    with changes:
        if result.changes:
            st.dataframe(pd.DataFrame(readable_rows([audit_row(row, "变更", result.headers) for row in result.changes])),
                         hide_index=True, width="stretch")
        else:
            st.write("本次没有需要变更的数据。")
    with issues:
        if result.issues:
            show_info = st.checkbox("同时显示舍入精度、零电量调整等说明")
            visible = result.issues if show_info else pending
            st.dataframe(pd.DataFrame(readable_rows([audit_row(row, "待核对", result.headers) for row in visible])),
                         hide_index=True, width="stretch")
            st.caption("相应来源、原值、差额和处理方式也会写入下载文件的本次更新记录。")
        else:
            st.write("本次未发现待核对事项。")

with st.expander("处理口径与后续更新"):
    st.write("按公司、月份、场站及明确分期信息匹配记录。默认补空缺并保留已有数值，冲突逐项记录。容量只有在来源单位明确时自动换算为 MW。")
    st.write("量价费核验会分别列出电量组成、电费组成和差额。只有总量、总费均有充分证据时，才据此修正错误价格；其他差异保留待核对。")
    st.write("每次生成都会重建利用小时辅助列和月份控制行，支持新增月份和场站。原工作簿历史审计页作为旧版本快照保留，请以本次更新记录查看本轮处理。")
    st.write("目前接收同一年度数据。原生 WPS 透视字段数量有限，若新增分类变化导致字段超限，程序会说明原因并停止生成。")
