#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import tempfile
import csv
import io
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook

from generate_report import (
    MAIN_SHEET,
    SIGN_SHEET,
    generate_report_data,
)


MAX_FILE_MB = 50


def _check_file_size(uploaded_file, max_mb=MAX_FILE_MB):
    size_mb = uploaded_file.size / (1024 * 1024)
    if size_mb > max_mb:
        raise ValueError(f"{uploaded_file.name} 超过 {max_mb}MB，当前 {size_mb:.1f}MB。")


def _check_required_sheet(path, expected_sheet, file_label):
    wb = load_workbook(path, read_only=True, data_only=True)
    if expected_sheet not in wb.sheetnames:
        raise ValueError(f"{file_label} 缺少工作表：{expected_sheet}")


def _save_uploaded(uploaded_file, target_path):
    target_path.write_bytes(uploaded_file.getbuffer())


def _trend_text(v):
    if v is None:
        return "--"
    if v > 0:
        return f"↑ {v * 100:.2f}%"
    if v < 0:
        return f"↓ {abs(v * 100):.2f}%"
    return "→ 0.00%"


st.set_page_config(page_title="GUS_排班情况总览", layout="wide")
st.title("GUS_排班情况总览（公网交互版）")
st.caption("上传数据 -> 一键生成考勤晾晒结果 -> 在线查看报表 -> 下载 HTML/JSON")

with st.expander("上传说明", expanded=False):
    st.markdown(
        "- 必传：主明细.xlsx + 签字报表.xlsx\n"
        "- 系统将自动解析 xlsx，生成结构化 JSON 与考勤报表\n"
        "- 可选：规则文档（用于留档，不参与计算）\n"
        f"- 单文件上限：{MAX_FILE_MB}MB"
    )

col1, col2 = st.columns(2)
with col1:
    main_file = st.file_uploader("1) 主明细（xlsx）", type=["xlsx"], key="main_file")
with col2:
    sign_file = st.file_uploader("2) 签字报表（xlsx）", type=["xlsx"], key="sign_file")

rule_doc = st.file_uploader("可选：规则文档（docx）", type=["docx"], key="rule_doc")

if st.button("生成报表", type="primary", use_container_width=True):
    if not all([main_file, sign_file]):
        st.error("请先上传2个必传文件后再生成。")
        st.stop()

    try:
        for f in [main_file, sign_file]:
            _check_file_size(f)

        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            main_path = td_path / "main.xlsx"
            sign_path = td_path / "sign.xlsx"
            _save_uploaded(main_file, main_path)
            _save_uploaded(sign_file, sign_path)
            if rule_doc:
                _save_uploaded(rule_doc, td_path / "rule.docx")

            # 输入校验：必须工作表
            _check_required_sheet(main_path, MAIN_SHEET, "主明细文件")
            _check_required_sheet(sign_path, SIGN_SHEET, "签字报表文件")

            with st.spinner("正在生成，请稍候..."):
                result, html = generate_report_data(
                    main_file=main_path,
                    sign_file=sign_path,
                )

        st.success("生成成功。")
        st.session_state["report_result"] = result
        st.session_state["report_html"] = html
        st.session_state["report_json_text"] = json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        st.error(f"生成失败：{e}")

if "report_result" in st.session_state and "report_html" in st.session_state:
    result = st.session_state["report_result"]
    html = st.session_state["report_html"]
    json_text = st.session_state["report_json_text"]
    kpi = result.get("kpi", {})
    trend = result.get("kpi_trend", {})
    detail_rows = result.get("detail_rows", [])

    st.subheader("核心指标")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("总人数", int(kpi.get("总人数", 0)))
    c2.metric("未排班数", int(kpi.get("未排班数", 0)))
    c3.metric("排班率", f"{(kpi.get('排班率') or 0)*100:.2f}%", _trend_text(trend.get("排班率")))
    c4.metric("HUB排班正确率", f"{(kpi.get('HUB排班正确率') or 0)*100:.2f}%", _trend_text(trend.get("HUB排班正确率")))
    c5.metric("缺卡率", f"{(kpi.get('缺卡率') or 0)*100:.2f}%", _trend_text(trend.get("缺卡率")))

    st.subheader("下载结果")
    d1, d2, d3 = st.columns(3)
    d1.download_button(
        "下载 attendance_report.html",
        data=html.encode("utf-8"),
        file_name="attendance_report.html",
        mime="text/html",
        use_container_width=True,
    )
    d2.download_button(
        "下载 attendance_report_data.json",
        data=json_text.encode("utf-8"),
        file_name="attendance_report_data.json",
        mime="application/json",
        use_container_width=True,
    )
    if detail_rows:
        csv_buf = io.StringIO()
        writer = csv.DictWriter(csv_buf, fieldnames=list(detail_rows[0].keys()))
        writer.writeheader()
        writer.writerows(detail_rows)
        d3.download_button(
            "下载 attendance_report_detail.csv",
            data=csv_buf.getvalue().encode("utf-8-sig"),
            file_name="attendance_report_detail.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.subheader("在线预览报表")
    st.components.v1.html(html, height=1200, scrolling=True)
