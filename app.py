from __future__ import annotations

import hashlib
import html
import re
from collections import Counter
from datetime import datetime

import pandas as pd
import streamlit as st

from config.column_mapping import DEFAULT_KEYWORD, MAX_KIOT_DELAY_DAYS
from core.bank_smart_matcher import reconcile_bank_smart
from core.bank_smart_reader import read_bank_history_file, read_sales_smart_file
from core.bank_smart_reporter import (
    bank_smart_results_to_dataframe,
    build_bank_smart_excel_report,
)
from core.excel_reader import read_hddt_files, read_kiot_file, read_smart_file
from core.invoice_matcher import reconcile_invoices
from core.history_store import (
    delete_all_history_records,
    delete_history_record,
    get_history_record,
    list_history_records,
    save_history_record,
)
from core.models import ExcelFile, Severity
from core.payment_matcher import reconcile_payments
from core.report_exporter import (
    build_excel_report,
    invoice_results_to_dataframe,
    issues_to_dataframe,
    payment_results_to_dataframe,
)


st.set_page_config(
    page_title="Đối chiếu công nợ nhà cung cấp",
    page_icon="✓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      :root {
        --navy-950: #0b1f33;
        --navy-800: #123a5a;
        --blue-700: #17628f;
        --blue-600: #1476a8;
        --blue-100: #e8f3f9;
        --slate-700: #34475a;
        --slate-500: #66788a;
        --slate-300: #cdd8e2;
        --slate-100: #eef3f7;
        --surface: #ffffff;
        --green-700: #1f6e50;
        --green-100: #eaf7f1;
        --amber-700: #8b5a00;
        --amber-100: #fff7e3;
        --red-700: #a13b36;
        --red-100: #fff0ef;
        --history-700: #684c7c;
        --history-500: #81649a;
        --history-100: #f3eef7;
      }

      html, body, [class*="css"] {
        font-family: "Segoe UI", Arial, sans-serif;
      }

      [data-testid="stAppViewContainer"] {
        background:
          radial-gradient(circle at 92% 2%, rgba(20, 118, 168, .08), transparent 24rem),
          #f5f8fb;
      }

      [data-testid="stHeader"] {
        background: rgba(245, 248, 251, .92);
      }

      .block-container {
        max-width: 1480px;
        padding-top: 2rem;
        padding-bottom: 3.5rem;
        padding-left: clamp(1rem, 3vw, 3rem);
        padding-right: clamp(1rem, 3vw, 3rem);
      }

      .accounting-hero {
        position: relative;
        overflow: hidden;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1.5rem;
        min-height: 148px;
        margin: .25rem 0 1.35rem;
        padding: 1.55rem 1.8rem;
        border: 1px solid rgba(23, 98, 143, .2);
        border-radius: 18px;
        background: linear-gradient(118deg, #0b2741 0%, #114a70 68%, #16739d 100%);
        box-shadow: 0 12px 30px rgba(15, 49, 77, .12);
        color: #fff;
      }

      .accounting-hero::after {
        content: "";
        position: absolute;
        right: -65px;
        top: -105px;
        width: 280px;
        height: 280px;
        border: 42px solid rgba(255, 255, 255, .06);
        border-radius: 50%;
      }

      .hero-copy {position: relative; z-index: 1; min-width: 0;}
      .hero-eyebrow {
        display: inline-flex;
        align-items: center;
        gap: .45rem;
        margin-bottom: .45rem;
        color: #bfe8f7;
        font-size: .84rem;
        font-weight: 700;
        letter-spacing: .095em;
        line-height: 1.5;
        text-transform: uppercase;
      }
      .hero-eyebrow::before {
        content: "";
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: #46d3a2;
        box-shadow: 0 0 0 4px rgba(70, 211, 162, .15);
      }
      .hero-title {
        margin: 0;
        padding: .08em 0 .12em;
        color: #fff;
        font-size: clamp(1.65rem, 3vw, 2.45rem);
        font-weight: 760;
        letter-spacing: -.025em;
        line-height: 1.3;
        overflow: visible;
      }
      .hero-subtitle {
        max-width: 780px;
        margin: .22rem 0 0;
        color: #d9ebf5;
        font-size: .96rem;
        line-height: 1.6;
      }
      .hero-security {
        position: relative;
        z-index: 1;
        flex: 0 0 auto;
        display: flex;
        align-items: center;
        gap: .65rem;
        padding: .72rem .9rem;
        border: 1px solid rgba(255, 255, 255, .18);
        border-radius: 12px;
        background: rgba(255, 255, 255, .09);
        color: #f4fbff;
        font-size: .9rem;
        font-weight: 650;
        backdrop-filter: blur(5px);
      }
      .security-check {
        display: grid;
        width: 28px;
        height: 28px;
        place-items: center;
        border-radius: 50%;
        background: #46d3a2;
        color: #0b3b2d;
        font-size: .95rem;
        font-weight: 900;
      }

      .workflow-strip {
        display: grid;
        grid-template-columns: 1fr 30px 1fr 30px 1fr;
        align-items: center;
        margin: 0 0 1.2rem;
        padding: .85rem 1rem;
        border: 1px solid #dce6ee;
        border-radius: 14px;
        background: rgba(255, 255, 255, .82);
      }
      .workflow-step {display: flex; align-items: center; gap: .75rem; min-width: 0;}
      .step-number {
        display: grid;
        flex: 0 0 34px;
        width: 34px;
        height: 34px;
        place-items: center;
        border-radius: 10px;
        background: var(--blue-100);
        color: var(--blue-700);
        font-weight: 800;
      }
      .step-label {color: var(--navy-950); font-size: 1rem; font-weight: 700; line-height: 1.4;}
      .step-note {color: var(--slate-500); font-size: .88rem; line-height: 1.45;}
      .workflow-arrow {color: #95a9ba; font-size: 1.25rem; text-align: center;}

      .section-kicker {
        margin-bottom: .45rem;
        color: var(--blue-700);
        font-size: .82rem;
        font-weight: 800;
        letter-spacing: .08em;
        text-transform: uppercase;
      }
      .section-title {
        margin: 0 0 .55rem;
        color: var(--navy-950);
        font-size: 1.25rem;
        font-weight: 760;
        line-height: 1.45;
      }
      .section-description {
        margin: 0 0 1.45rem;
        color: var(--slate-500);
        font-size: 1rem;
        line-height: 1.6;
      }
      .upload-heading {
        margin: 0 0 .38rem;
        color: var(--navy-800);
        font-size: 1rem;
        font-weight: 720;
        line-height: 1.45;
      }
      .upload-note {
        margin: 0 0 .65rem;
        min-height: 2.4rem;
        color: var(--slate-500);
        font-size: .9rem;
        line-height: 1.5;
      }

      [data-testid="stVerticalBlockBorderWrapper"] {
        border-color: #d9e4ec !important;
        border-radius: 15px !important;
        background: rgba(255, 255, 255, .94);
        box-shadow: 0 5px 18px rgba(26, 58, 83, .045);
      }
      [data-testid="stFileUploaderDropzone"] {
        min-height: 112px;
        border: 1.5px dashed #aac2d3;
        border-radius: 12px;
        background: #f8fbfd;
      }
      [data-testid="stFileUploaderDropzone"]:hover {
        border-color: var(--blue-600);
        background: #f0f8fc;
      }
      [data-testid="stFileUploaderDropzone"] button {
        border-color: #9db8ca;
        color: var(--navy-800);
        background: #fff;
      }
      [data-testid="stFileUploaderFile"] {
        align-items: center;
        min-height: 52px;
        margin-top: .42rem;
        padding: .42rem .5rem;
        border: 1px solid #dce6ee;
        border-radius: 10px;
        background: #fff;
      }
      [data-testid="stFileUploaderFile"] > div:first-child {
        position: relative;
        display: grid;
        flex: 0 0 34px;
        width: 34px;
        height: 38px;
        margin-right: .4rem;
        padding: 0;
        place-items: center;
        overflow: hidden;
        border-radius: 5px;
        background: linear-gradient(145deg, #17864d, #0f6c3a);
        box-shadow: 0 3px 8px rgba(15, 108, 58, .2);
        color: #fff;
      }
      [data-testid="stFileUploaderFile"] > div:first-child svg {display: none;}
      [data-testid="stFileUploaderFile"] > div:first-child::before {
        content: "X";
        position: relative;
        z-index: 1;
        font-family: "Segoe UI", Arial, sans-serif;
        font-size: 1rem;
        font-weight: 800;
        line-height: 1;
      }
      [data-testid="stFileUploaderFile"] > div:first-child::after {
        content: "";
        position: absolute;
        top: 0;
        right: 0;
        width: 9px;
        height: 9px;
        background: rgba(255, 255, 255, .62);
        clip-path: polygon(0 0, 100% 100%, 100% 0);
      }
      [data-testid="stFileUploaderFileName"] {
        color: var(--navy-950);
        font-weight: 650;
      }

      .stTabs [data-baseweb="tab-list"] {
        gap: .45rem;
        padding: .35rem;
        border: 1px solid #dce6ee;
        border-radius: 13px;
        background: #edf3f7;
      }
      .stTabs [data-baseweb="tab"] {
        height: 44px;
        padding: 0 1.15rem;
        border-radius: 9px;
        color: #516577;
        font-weight: 700;
      }
      .stTabs [aria-selected="true"] {
        background: #fff !important;
        color: var(--navy-800) !important;
        box-shadow: 0 2px 8px rgba(22, 64, 94, .08);
      }
      .stTabs [data-baseweb="tab-highlight"] {display: none;}

      [data-testid="stMetric"] {
        min-height: 104px;
        padding: .9rem 1rem;
        border: 1px solid #dce6ee;
        border-radius: 12px;
        background: #fff;
        box-shadow: 0 3px 12px rgba(26, 58, 83, .04);
      }
      [data-testid="stMetricLabel"] {color: var(--slate-500); font-weight: 650;}
      [data-testid="stMetricValue"] {color: var(--navy-950); font-weight: 760;}
      [data-testid="stCaptionContainer"] p {font-size: .9rem; line-height: 1.5;}

      .stButton > button[kind="primary"],
      .stDownloadButton > button {
        min-height: 46px;
        border-radius: 10px;
        font-weight: 720;
        box-shadow: 0 4px 12px rgba(20, 91, 132, .12);
      }
      .stButton > button[kind="primary"] {
        border-color: var(--blue-700);
        background: linear-gradient(135deg, #17628f, #1476a8);
      }
      .stButton > button[kind="primary"]:hover {background: #104f77;}
      .stButton > button[kind="secondary"] {
        min-height: 44px;
        border: 1px solid #aa98bb;
        border-radius: 10px;
        background: var(--history-100);
        color: var(--history-700);
        font-weight: 750;
        box-shadow: 0 3px 10px rgba(104, 76, 124, .1);
      }
      .stButton > button[kind="secondary"]:hover {
        border-color: var(--history-500);
        background: #ebe2f2;
        color: #503760;
      }
      .stDownloadButton > button {
        border-color: var(--blue-700);
        color: var(--blue-700);
        background: #fff;
      }

      [data-testid="stDataFrame"] {
        overflow: hidden;
        border: 1px solid #dce6ee;
        border-radius: 12px;
      }
      [data-testid="stExpander"] {
        border-color: #dce6ee;
        border-radius: 11px;
        background: #fff;
      }

      .locked-box, .warning-box, .success-box {
        margin: .8rem 0;
        padding: 1rem 1.15rem;
        border-radius: 11px;
        font-size: 1rem;
        line-height: 1.55;
      }
      .locked-box {border: 1px solid #ecd18a; background: var(--amber-100); color: var(--amber-700);}
      .warning-box {border: 1px solid #ecd18a; background: var(--amber-100); color: var(--amber-700);}
      .success-box {border: 1px solid #a9dbc7; background: var(--green-100); color: var(--green-700);}

      .app-footer {
        display: flex;
        justify-content: space-between;
        gap: 1rem;
        color: #718294;
        font-size: .86rem;
        line-height: 1.5;
      }

      .history-separator {
        display: flex;
        align-items: center;
        gap: .8rem;
        margin: 1.15rem 0 .7rem;
        color: var(--history-700);
        font-size: .78rem;
        font-weight: 800;
        letter-spacing: .075em;
        text-transform: uppercase;
      }
      .history-separator::before,
      .history-separator::after {
        content: "";
        height: 1px;
        background: #d8cce2;
      }
      .history-separator::before {width: 28px;}
      .history-separator::after {flex: 1;}
      .history-tab-gap {height: 1rem;}
      .history-panel-banner {
        margin: 0 0 1.1rem;
        padding: 1rem 1.15rem;
        border: 1px solid #d8cce2;
        border-left: 5px solid var(--history-500);
        border-radius: 12px;
        background: linear-gradient(120deg, #f6f2f9 0%, #eee6f4 100%);
      }
      .history-kicker {
        margin-bottom: .4rem;
        color: var(--history-700);
        font-size: .8rem;
        font-weight: 800;
        letter-spacing: .07em;
        text-transform: uppercase;
      }
      .history-title {
        margin-bottom: .42rem;
        color: #3f2d4b;
        font-size: 1.3rem;
        font-weight: 760;
      }
      .history-description {color: #695b73; font-size: .94rem; line-height: 1.55;}
      .history-current-label {
        margin: .15rem 0 .38rem;
        color: var(--history-700);
        font-size: .9rem;
        font-weight: 750;
      }
      .history-current-record {
        min-height: 54px;
        margin-bottom: .65rem;
        padding: .85rem 1rem;
        border: 1.5px solid #aa98bb;
        border-radius: 10px;
        background: #fff;
        color: #34263d;
        font-size: .98rem;
        font-weight: 650;
        line-height: 1.5;
        box-shadow: 0 2px 8px rgba(104, 76, 124, .08);
      }
      [data-testid="stPopover"] button {
        min-height: 44px;
        border: 1px solid #9a82ad;
        border-radius: 10px;
        background: var(--history-100);
        color: var(--history-700) !important;
        font-weight: 750;
      }
      .st-key-request_delete_selected_history button,
      .st-key-request_delete_all_history button {
        border-color: #d9a2a2 !important;
        background: #fff8f8 !important;
        color: #a32121 !important;
        box-shadow: none !important;
      }
      .st-key-request_delete_selected_history button:hover,
      .st-key-request_delete_all_history button:hover {
        border-color: #b42323 !important;
        background: #fff0f0 !important;
        color: #861818 !important;
      }
      .st-key-confirm_delete_selected_history button,
      .st-key-confirm_delete_all_history button {
        border-color: #a51d1d !important;
        background: #b42323 !important;
        color: #fff !important;
        box-shadow: 0 4px 12px rgba(180, 35, 35, .18) !important;
      }
      .st-key-confirm_delete_selected_history button:hover,
      .st-key-confirm_delete_all_history button:hover {
        border-color: #861818 !important;
        background: #961c1c !important;
        color: #fff !important;
      }

      @media (max-width: 760px) {
        .block-container {padding-top: 1rem;}
        .accounting-hero {align-items: flex-start; flex-direction: column; padding: 1.25rem;}
        .hero-security {width: 100%;}
        .workflow-strip {grid-template-columns: 1fr; gap: .65rem;}
        .workflow-arrow {transform: rotate(90deg);}
        .upload-note {min-height: auto;}
        .app-footer {flex-direction: column;}
      }
    </style>
    """,
    unsafe_allow_html=True,
)


def uploaded_to_source(uploaded) -> ExcelFile:
    return ExcelFile(name=uploaded.name, content=uploaded.getvalue())


def file_signature(files) -> str:
    digest = hashlib.sha256()
    for uploaded in files:
        digest.update(uploaded.name.encode("utf-8"))
        digest.update(uploaded.getvalue())
    return digest.hexdigest()


def render_issues(issues, title="Kiểm tra dữ liệu"):
    if not issues:
        st.success(f"{title}: không phát hiện lỗi cấu trúc.")
        return
    errors = [issue for issue in issues if issue.severity is Severity.ERROR]
    warnings = [issue for issue in issues if issue.severity is Severity.WARNING]
    if errors:
        st.error(f"Phát hiện {len(errors)} lỗi cần xử lý.")
    if warnings:
        st.warning(f"Có {len(warnings)} cảnh báo.")
    with st.expander(f"Chi tiết lỗi/cảnh báo ({len(issues)})", expanded=bool(errors)):
        st.dataframe(issues_to_dataframe(issues), use_container_width=True, hide_index=True)


def highlight_status(row):
    status = str(row.get("Trạng thái", "")).upper()
    if status == "KHỚP":
        color = "background-color: #e2f0d9"
    elif "SMART THIẾU" in status:
        color = "background-color: #ddebf7; color: #1f4e78"
    elif "KIOT THIẾU" in status:
        color = "background-color: #f4cccc; color: #9c342f"
    elif any(marker in status for marker in ("CẦN KIỂM TRA", "TRÙNG", "CẢNH BÁO", "THIẾU NỘI DUNG")):
        color = "background-color: #fff2cc"
    else:
        color = "background-color: #fce4d6"
    return [color if column == "Trạng thái" else "" for column in row.index]


def render_filtered_table(dataframe: pd.DataFrame, key: str):
    if dataframe.empty:
        st.info("Không có dữ liệu để hiển thị.")
        return
    statuses = sorted(str(value) for value in dataframe["Trạng thái"].dropna().unique())
    selected = st.multiselect(
        "Lọc trạng thái",
        statuses,
        default=statuses,
        key=f"status_filter_{key}",
    )
    filtered = dataframe[dataframe["Trạng thái"].isin(selected)] if selected else dataframe.iloc[0:0]
    formatters = {
        column: "{:,.0f}"
        for column in [
            "Tổng tiền HĐĐT",
            "Tổng tiền Smart",
            "Chênh lệch",
            "Số tiền Smart",
            "Số tiền KIOT",
            "Số tiền ngân hàng",
        ]
        if column in filtered.columns
    }
    if "Số ngày chênh lệch" in filtered.columns:
        formatters["Số ngày chênh lệch"] = "{:.0f}"
    st.dataframe(
        filtered.style.apply(highlight_status, axis=1).format(formatters, na_rep=""),
        use_container_width=True,
        hide_index=True,
        height=min(620, 80 + 35 * max(1, len(filtered))),
    )


def safe_supplier_filename(name: str) -> str:
    text = re.sub(r"[^A-Za-z0-9À-ỹ]+", "_", name.strip(), flags=re.UNICODE).strip("_")
    return text[:60] or "NHA_CUNG_CAP"


def reset_result_if_inputs_changed(signature_key: str, signature: str, result_keys: list[str]):
    previous = st.session_state.get(signature_key)
    if previous is not None and previous != signature:
        for key in result_keys:
            st.session_state.pop(key, None)
    st.session_state[signature_key] = signature


def save_history_snapshot(
    *,
    reconciliation_type: str,
    title: str,
    input_files: list[str],
    status_counts: Counter,
    summary: dict[str, object],
    results: pd.DataFrame,
    issues: pd.DataFrame,
    report_filename: str,
    report_bytes: bytes,
) -> str | None:
    try:
        return save_history_record(
            reconciliation_type=reconciliation_type,
            title=title,
            input_files=input_files,
            status_counts=dict(status_counts),
            summary=summary,
            results=results,
            issues=issues,
            report_filename=report_filename,
            report_bytes=report_bytes,
        )
    except Exception as exc:
        st.warning(f"Đối chiếu đã hoàn tất nhưng chưa lưu được lịch sử: {exc}")
        return None


def _history_label(record) -> str:
    try:
        created = datetime.fromisoformat(record.created_at).strftime("%d/%m/%Y %H:%M:%S")
    except ValueError:
        created = record.created_at
    return f"{created} · {record.title}"


def render_history_delete_controls(record, record_label: str):
    pending = st.session_state.get("pending_history_delete")
    if (
        pending
        and pending.get("scope") == "selected"
        and pending.get("record_id") != record.record_id
    ):
        st.session_state.pop("pending_history_delete", None)
        pending = None

    if not pending:
        label_column, delete_selected_column, delete_all_column = st.columns(
            [3.8, 1.6, 1.7],
            vertical_alignment="center",
        )
        with label_column:
            st.caption("Thao tác với lịch sử đang xem")
        with delete_selected_column:
            if st.button(
                "Xóa lần đối chiếu này",
                width="stretch",
                key="request_delete_selected_history",
            ):
                st.session_state.pending_history_delete = {
                    "scope": "selected",
                    "record_id": record.record_id,
                }
                st.rerun()
        with delete_all_column:
            if st.button(
                "Xóa toàn bộ lịch sử",
                width="stretch",
                key="request_delete_all_history",
            ):
                st.session_state.pending_history_delete = {"scope": "all"}
                st.rerun()
        return

    deleting_all = pending.get("scope") == "all"
    st.error(
        "Xóa toàn bộ lịch sử?"
        if deleting_all
        else "Xóa lần đối chiếu này?"
    )
    if deleting_all:
        st.warning(
            "Toàn bộ lịch sử đối chiếu và tất cả báo cáo Excel đã lưu sẽ bị "
            "xóa vĩnh viễn. Các file dữ liệu gốc không bị ảnh hưởng."
        )
    else:
        st.warning(
            "Lịch sử và báo cáo Excel của lần đối chiếu này sẽ bị xóa vĩnh "
            "viễn. Các file dữ liệu gốc không bị ảnh hưởng."
        )
        st.caption(record_label)

    cancel_column, delete_column = st.columns(2)
    with cancel_column:
        if st.button(
            "Không, quay lại",
            width="stretch",
            key="cancel_history_delete",
        ):
            st.session_state.pop("pending_history_delete", None)
            st.rerun()
    with delete_column:
        confirm_label = "Xóa toàn bộ" if deleting_all else "Xóa vĩnh viễn"
        confirm_key = (
            "confirm_delete_all_history"
            if deleting_all
            else "confirm_delete_selected_history"
        )
        if st.button(
            confirm_label,
            type="primary",
            width="stretch",
            key=confirm_key,
        ):
            try:
                if deleting_all:
                    deleted_count = delete_all_history_records()
                    notice = (
                        f"Đã xóa toàn bộ {deleted_count} lần đối chiếu khỏi lịch sử."
                    )
                else:
                    deleted = delete_history_record(record.record_id)
                    notice = (
                        "Đã xóa lần đối chiếu đã chọn."
                        if deleted
                        else "Lần đối chiếu này đã được xóa trước đó."
                    )
            except Exception as exc:
                st.error(f"Không thể xóa lịch sử: {exc}")
                return
            st.session_state.pop("pending_history_delete", None)
            st.session_state.pop("selected_history_record", None)
            st.session_state.history_notice = notice
            st.rerun()


def render_history_panel():
    st.markdown(
        """
        <div class="history-panel-banner">
          <div class="history-kicker">Chỉ xem lại kết quả · Lưu trên máy tính này</div>
          <div class="history-title">Lịch sử hoạt động</div>
          <div class="history-description">Đây không phải tab đối chiếu. Khu vực này dùng để mở lại kết quả và tải lại báo cáo sau khi làm mới trang.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    history_notice = st.session_state.pop("history_notice", None)
    if history_notice:
        st.success(history_notice)
    try:
        records = list_history_records(limit=100)
    except Exception as exc:
        st.error(f"Không thể đọc lịch sử hoạt động: {exc}")
        return
    if not records:
        st.info("Chưa có lịch sử. Kết quả sẽ tự động được lưu sau lần đối chiếu thành công tiếp theo.")
        return

    records_by_id = {record.record_id: record for record in records}
    record_ids = list(records_by_id)
    history_choice_key = "selected_history_record"
    if st.session_state.get(history_choice_key) not in records_by_id:
        st.session_state[history_choice_key] = record_ids[0]
    selected_id = st.session_state[history_choice_key]
    current_label = html.escape(_history_label(records_by_id[selected_id]))
    st.markdown(
        '<div class="history-current-label">Lần đối chiếu đang xem</div>'
        f'<div class="history-current-record">{current_label}</div>',
        unsafe_allow_html=True,
    )
    with st.popover("Chọn lần đối chiếu khác", width="stretch"):
        st.caption("Chọn một dòng bên dưới. Nội dung lịch sử chỉ đọc và không thể chỉnh sửa.")
        st.radio(
            "Danh sách các lần đối chiếu",
            options=record_ids,
            format_func=lambda record_id: _history_label(records_by_id[record_id]),
            key=history_choice_key,
            label_visibility="collapsed",
        )
    selected_id = st.session_state[history_choice_key]
    record = get_history_record(selected_id)
    if record is None:
        st.warning("Bản ghi lịch sử không còn tồn tại.")
        return

    render_history_delete_controls(record, _history_label(record))
    st.caption("File đã sử dụng: " + " · ".join(record.input_files))
    if record.status_counts:
        statuses = list(record.status_counts.items())
        for start in range(0, len(statuses), 4):
            batch = statuses[start : start + 4]
            columns = st.columns(len(batch))
            for column, (status, count) in zip(columns, batch):
                column.metric(status.title(), count)
    if record.summary:
        with st.expander("Thông tin tổng hợp", expanded=False):
            summary_frame = pd.DataFrame(
                [(key, value) for key, value in record.summary.items()],
                columns=["Nội dung", "Giá trị"],
            )
            st.dataframe(summary_frame, use_container_width=True, hide_index=True)

    render_filtered_table(record.results, f"history_{record.record_id}")
    if not record.issues.empty:
        with st.expander(f"Lỗi/cảnh báo đã ghi nhận ({len(record.issues)})", expanded=False):
            st.dataframe(record.issues, use_container_width=True, hide_index=True)
    st.download_button(
        "Tải lại báo cáo Excel",
        data=record.report_bytes,
        file_name=record.report_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key=f"history_download_{record.record_id}",
    )


if st.session_state.get("state_schema_version") != 5:
    for legacy_key in [
        "hddt_data",
        "smart_data",
        "invoice_smart_data",
        "payment_smart_data",
        "invoice_reconciliation",
        "payment_reconciliation",
        "kiot_data",
        "step1_input_signature",
        "step2_input_signature",
        "invoice_input_signature",
        "payment_input_signature",
        "bank_history_data",
        "sales_smart_data",
        "bank_smart_reconciliation",
        "bank_smart_input_signature",
        "invoice_report_bytes",
        "invoice_report_filename",
        "payment_report_bytes",
        "payment_report_filename",
        "bank_report_bytes",
        "bank_report_filename",
    ]:
        st.session_state.pop(legacy_key, None)
    st.session_state.state_schema_version = 5


st.markdown(
    """
    <section class="accounting-hero">
      <div class="hero-copy">
        <div class="hero-eyebrow">Hệ thống kiểm soát công nợ</div>
        <h1 class="hero-title">Đối chiếu công nợ nhà cung cấp</h1>
        <p class="hero-subtitle">Đối chiếu chính xác dữ liệu Hóa đơn điện tử, Smart Pro và KIOT Việt theo quy trình đã phê duyệt.</p>
      </div>
      <div class="hero-security">
        <span class="security-check">✓</span>
        <span>Xử lý nội bộ<br>trên máy tính này</span>
      </div>
    </section>
    <div class="workflow-strip">
      <div class="workflow-step">
        <span class="step-number">A</span>
        <span><div class="step-label">HĐĐT ↔ Smart Pro</div><div class="step-note">Đối chiếu hóa đơn độc lập</div></span>
      </div>
      <div class="workflow-arrow">•</div>
      <div class="workflow-step">
        <span class="step-number">B</span>
        <span><div class="step-label">Smart Pro ↔ KIOT Việt</div><div class="step-note">Đối chiếu thanh toán độc lập</div></span>
      </div>
      <div class="workflow-arrow">•</div>
      <div class="workflow-step">
        <span class="step-number">C</span>
        <span><div class="step-label">Ngân hàng ↔ Smart</div><div class="step-note">Đối chiếu tiền vào độc lập</div></span>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="history-separator">Khu vực xem lại</div>',
    unsafe_allow_html=True,
)
history_spacer, history_button_column = st.columns([5, 1.35])
with history_button_column:
    history_label = (
        "Đóng lịch sử"
        if st.session_state.get("show_history_panel", False)
        else "Lịch sử hoạt động"
    )
    if st.button(history_label, use_container_width=True, key="toggle_history_panel"):
        st.session_state.show_history_panel = not st.session_state.get(
            "show_history_panel", False
        )
        st.rerun()

if st.session_state.get("show_history_panel", False):
    with st.container(border=True):
        render_history_panel()

st.markdown('<div class="history-tab-gap"></div>', unsafe_allow_html=True)

tab_invoice, tab_payment, tab_bank = st.tabs(
    ["HĐĐT ↔ SMART", "SMART ↔ KIOT", "NGÂN HÀNG ↔ SMART"]
)

with tab_invoice:
    st.markdown(
        """
        <div class="section-kicker" style="margin-top:1.15rem">Đối chiếu hóa đơn</div>
        <div class="section-title">Đối chiếu hóa đơn HĐĐT với Smart Pro</div>
        <div class="section-description">Kiểm tra số hóa đơn và tổng tiền theo MST nhà cung cấp. Chức năng này hoạt động độc lập với Smart–KIOT.</div>
        """,
        unsafe_allow_html=True,
    )
    with st.container(border=True):
        invoice_upload_col1, invoice_upload_col2 = st.columns(2)
        with invoice_upload_col1:
            st.markdown(
                '<div class="upload-heading">01 · Hóa đơn điện tử theo tháng</div>'
                '<div class="upload-note">Chọn tất cả file từ tháng 1 đến tháng hiện tại. Có thể bổ sung nhiều tháng trong cùng một lần tải.</div>',
                unsafe_allow_html=True,
            )
            hddt_uploads = st.file_uploader(
                "Tải các file HĐĐT",
                type=["xlsx"],
                accept_multiple_files=True,
                help="Kỳ năm hiện hành được kiểm tra động đến tháng đang chạy tool.",
                key="invoice_hddt_uploads",
            )
        with invoice_upload_col2:
            st.markdown(
                '<div class="upload-heading">02 · Smart Pro công nợ</div>'
                '<div class="upload-note">Chọn một file Smart công nợ lũy kế dùng riêng cho đối chiếu HĐĐT.</div>',
                unsafe_allow_html=True,
            )
            invoice_smart_upload = st.file_uploader(
                "Tải file Smart công nợ cho HĐĐT",
                type=["xlsx"],
                accept_multiple_files=False,
                key="invoice_smart_upload",
            )

    invoice_files = list(hddt_uploads or []) + (
        [invoice_smart_upload] if invoice_smart_upload else []
    )
    invoice_signature = file_signature(invoice_files) if invoice_files else ""
    reset_result_if_inputs_changed(
        "invoice_input_signature",
        invoice_signature,
        [
            "hddt_data",
            "invoice_smart_data",
            "invoice_reconciliation",
            "invoice_report_bytes",
            "invoice_report_filename",
        ],
    )

    if not hddt_uploads or invoice_smart_upload is None:
        st.info("Hãy chọn ít nhất một file HĐĐT và một file Smart để bắt đầu.")
    run_invoice = st.button(
        "Kiểm tra và đối chiếu HĐĐT–Smart",
        type="primary",
        disabled=not hddt_uploads or invoice_smart_upload is None,
        use_container_width=True,
    )
    if run_invoice:
        try:
            with st.spinner("Đang đọc và đối chiếu hóa đơn..."):
                hddt_data = read_hddt_files([uploaded_to_source(file) for file in hddt_uploads])
                invoice_smart_data = read_smart_file(uploaded_to_source(invoice_smart_upload))
                invoice_reconciliation = reconcile_invoices(hddt_data, invoice_smart_data)
                st.session_state.hddt_data = hddt_data
                st.session_state.invoice_smart_data = invoice_smart_data
                st.session_state.invoice_reconciliation = invoice_reconciliation
                st.session_state.invoice_input_signature = invoice_signature
                invoice_report_filename = (
                    f"Ket_qua_HDDT_Smart_"
                    f"{safe_supplier_filename(invoice_smart_data.supplier_name)}.xlsx"
                )
                invoice_report_bytes = build_excel_report(
                    hddt_data,
                    invoice_smart_data,
                    invoice_reconciliation,
                )
                invoice_statuses = Counter(
                    result.status for result in invoice_reconciliation.results
                )
                relevant_hddt_total = sum(
                    record.amount
                    for record in hddt_data.invoices
                    if record.supplier_tax_id == invoice_smart_data.supplier_tax_id
                )
                smart_invoice_total = sum(
                    record.amount for record in invoice_smart_data.invoices
                )
                save_history_snapshot(
                    reconciliation_type="HDDT_SMART",
                    title=(
                        "HĐĐT ↔ Smart · "
                        f"{invoice_smart_data.supplier_name or 'Chưa xác định nhà cung cấp'}"
                    ),
                    input_files=[file.name for file in hddt_uploads]
                    + [invoice_smart_upload.name],
                    status_counts=invoice_statuses,
                    summary={
                        "Kỳ Smart": invoice_smart_data.period.label,
                        "MST nhà cung cấp": invoice_smart_data.supplier_tax_id,
                        "Tổng kết quả": len(invoice_reconciliation.results),
                        "Tổng tiền HĐĐT": relevant_hddt_total,
                        "Tổng tiền Smart": smart_invoice_total,
                        "Chênh lệch": relevant_hddt_total - smart_invoice_total,
                    },
                    results=invoice_results_to_dataframe(invoice_reconciliation),
                    issues=issues_to_dataframe(invoice_reconciliation.issues),
                    report_filename=invoice_report_filename,
                    report_bytes=invoice_report_bytes,
                )
                st.session_state.invoice_report_bytes = invoice_report_bytes
                st.session_state.invoice_report_filename = invoice_report_filename
        except Exception as exc:
            st.error(f"Không thể xử lý file: {exc}")

    invoice_reconciliation = st.session_state.get("invoice_reconciliation")
    hddt_data = st.session_state.get("hddt_data")
    invoice_smart_data = st.session_state.get("invoice_smart_data")
    if invoice_reconciliation and hddt_data and invoice_smart_data:
        st.caption(
            f"Nhà cung cấp Smart: {invoice_smart_data.supplier_name or 'Không xác định'} · "
            f"MST: {invoice_smart_data.supplier_tax_id or 'Không xác định'} · Kỳ Smart: {invoice_smart_data.period.label}"
        )
        statuses = Counter(result.status for result in invoice_reconciliation.results)
        special_warning_count = sum(
            issue.code == "HDDT_SPECIAL_STATUS" for issue in invoice_reconciliation.issues
        )
        metric_columns = st.columns(6)
        metric_columns[0].metric("Tổng kết quả", len(invoice_reconciliation.results))
        metric_columns[1].metric("Khớp", statuses.get("KHỚP", 0))
        metric_columns[2].metric("Lệch tiền", statuses.get("LỆCH TIỀN", 0))
        metric_columns[3].metric(
            "Smart thiếu", statuses.get("SMART THIẾU HÓA ĐƠN", 0)
        )
        metric_columns[4].metric(
            "HĐĐT thiếu", statuses.get("HĐĐT THIẾU HÓA ĐƠN", 0)
        )
        metric_columns[5].metric("Cảnh báo HĐĐT", special_warning_count)
        relevant_hddt_total = sum(
            record.amount
            for record in hddt_data.invoices
            if record.supplier_tax_id == invoice_smart_data.supplier_tax_id
        )
        smart_invoice_total = sum(record.amount for record in invoice_smart_data.invoices)
        total_columns = st.columns(3)
        total_columns[0].metric("Tổng tiền HĐĐT", f"{relevant_hddt_total:,.0f} đ")
        total_columns[1].metric("Tổng tiền Smart", f"{smart_invoice_total:,.0f} đ")
        total_columns[2].metric(
            "Chênh lệch tổng",
            f"{relevant_hddt_total - smart_invoice_total:+,.0f} đ",
        )
        if invoice_reconciliation.is_fully_matched:
            st.markdown(
                '<div class="success-box"><b>HĐĐT–SMART KHỚP HOÀN TOÀN.</b> Các cảnh báo trạng thái hóa đơn (nếu có) không làm thay đổi kết quả đối chiếu.</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div class="warning-box"><b>HĐĐT–SMART CẦN KIỂM TRA.</b> Bạn vẫn có thể thực hiện Smart–KIOT độc lập tại tab bên cạnh.</div>',
                unsafe_allow_html=True,
            )
        render_issues(invoice_reconciliation.issues)
        invoice_df = invoice_results_to_dataframe(invoice_reconciliation)
        render_filtered_table(invoice_df, "invoice")
        report_bytes = st.session_state.get("invoice_report_bytes") or build_excel_report(
            hddt_data, invoice_smart_data, invoice_reconciliation
        )
        filename = st.session_state.get("invoice_report_filename") or (
            f"Ket_qua_HDDT_Smart_{safe_supplier_filename(invoice_smart_data.supplier_name)}.xlsx"
        )
        st.download_button(
            "Tải báo cáo HĐĐT–Smart",
            data=report_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with tab_payment:
    st.markdown(
        """
        <div class="section-kicker" style="margin-top:1.15rem">Đối chiếu thanh toán</div>
        <div class="section-title">Đối chiếu thanh toán Smart Pro với KIOT Việt</div>
        <div class="section-description">Chỉ cần file Smart và KIOT. Chức năng hoạt động độc lập, không phụ thuộc kết quả HĐĐT–Smart.</div>
        """,
        unsafe_allow_html=True,
    )
    with st.container(border=True):
        payment_upload_col1, payment_upload_col2 = st.columns(2)
        with payment_upload_col1:
            st.markdown(
                '<div class="upload-heading">01 · Smart Pro công nợ</div>'
                '<div class="upload-note">Chọn file Smart có các khoản Đã thanh toán cần đối chiếu.</div>',
                unsafe_allow_html=True,
            )
            payment_smart_upload = st.file_uploader(
                "Tải file Smart công nợ cho KIOT",
                type=["xlsx"],
                accept_multiple_files=False,
                key="payment_smart_upload",
            )
        with payment_upload_col2:
            st.markdown(
                '<div class="upload-heading">02 · KIOT Việt</div>'
                '<div class="upload-note">Chọn báo cáo công nợ chi tiết của đúng nhà cung cấp.</div>',
                unsafe_allow_html=True,
            )
            kiot_upload = st.file_uploader(
                "Tải file KIOT Việt",
                type=["xlsx"],
                accept_multiple_files=False,
                key="payment_kiot_upload",
            )
        with st.expander("Thiết lập nhận diện giao dịch KIOT", expanded=False):
            setting_col1, setting_col2 = st.columns([2, 1])
            with setting_col1:
                keyword = st.text_input(
                    "Từ khóa nhận diện giao dịch",
                    value=DEFAULT_KEYWORD,
                    help="Tool tìm từ khóa không phân biệt chữ hoa, chữ thường và dấu tiếng Việt.",
                    key="payment_keyword",
                )
            with setting_col2:
                st.text_input(
                    "Độ trễ tối đa của KIOT",
                    value=f"{MAX_KIOT_DELAY_DAYS} ngày",
                    disabled=True,
                    key="payment_delay",
                )

    payment_files = (
        ([payment_smart_upload] if payment_smart_upload else [])
        + ([kiot_upload] if kiot_upload else [])
    )
    payment_base_signature = file_signature(payment_files) if payment_files else ""
    payment_signature = hashlib.sha256(
        f"{payment_base_signature}|{keyword}".encode("utf-8")
    ).hexdigest()
    reset_result_if_inputs_changed(
        "payment_input_signature",
        payment_signature,
        [
            "payment_smart_data",
            "kiot_data",
            "payment_reconciliation",
            "payment_report_bytes",
            "payment_report_filename",
        ],
    )

    missing_payment_files = []
    if payment_smart_upload is None:
        missing_payment_files.append("Smart Pro")
    if kiot_upload is None:
        missing_payment_files.append("KIOT Việt")
    if missing_payment_files:
        st.info("Hãy chọn file " + " và ".join(missing_payment_files) + " để bắt đầu.")

    run_payment = st.button(
        "Đối chiếu Smart–KIOT",
        type="primary",
        disabled=bool(missing_payment_files),
        use_container_width=True,
    )
    if run_payment:
        try:
            with st.spinner("Đang đọc file, tìm từ khóa và ghép giao dịch..."):
                payment_smart_data = read_smart_file(uploaded_to_source(payment_smart_upload))
                kiot_data = read_kiot_file(uploaded_to_source(kiot_upload), keyword)
                payment_reconciliation = reconcile_payments(
                    payment_smart_data,
                    kiot_data,
                    max_delay_days=MAX_KIOT_DELAY_DAYS,
                )
                st.session_state.payment_smart_data = payment_smart_data
                st.session_state.kiot_data = kiot_data
                st.session_state.payment_reconciliation = payment_reconciliation
                st.session_state.payment_input_signature = payment_signature
                payment_end_date = (
                    payment_smart_data.period.end.strftime("%Y-%m-%d")
                    if payment_smart_data.period.end
                    else "KET_QUA"
                )
                payment_report_filename = (
                    f"Ket_qua_Smart_KIOT_"
                    f"{safe_supplier_filename(payment_smart_data.supplier_name)}_"
                    f"{payment_end_date}.xlsx"
                )
                payment_report_bytes = build_excel_report(
                    None,
                    payment_smart_data,
                    None,
                    kiot_data,
                    payment_reconciliation,
                )
                payment_statuses = Counter(
                    result.status for result in payment_reconciliation.results
                )
                smart_payment_total = sum(
                    record.amount for record in payment_smart_data.payments
                )
                kiot_payment_total = sum(record.amount for record in kiot_data.payments)
                save_history_snapshot(
                    reconciliation_type="SMART_KIOT",
                    title=(
                        "Smart ↔ KIOT · "
                        f"{payment_smart_data.supplier_name or 'Chưa xác định nhà cung cấp'}"
                    ),
                    input_files=[payment_smart_upload.name, kiot_upload.name],
                    status_counts=payment_statuses,
                    summary={
                        "Kỳ Smart": payment_smart_data.period.label,
                        "Kỳ KIOT": kiot_data.period.label,
                        "Từ khóa": kiot_data.keyword,
                        "Tổng kết quả": len(payment_reconciliation.results),
                        "Tổng tiền Smart": smart_payment_total,
                        "Tổng tiền KIOT": kiot_payment_total,
                        "Chênh lệch": smart_payment_total - kiot_payment_total,
                    },
                    results=payment_results_to_dataframe(payment_reconciliation),
                    issues=issues_to_dataframe(payment_reconciliation.issues),
                    report_filename=payment_report_filename,
                    report_bytes=payment_report_bytes,
                )
                st.session_state.payment_report_bytes = payment_report_bytes
                st.session_state.payment_report_filename = payment_report_filename
        except Exception as exc:
            st.error(f"Không thể xử lý file Smart hoặc KIOT: {exc}")

    payment_smart_data = st.session_state.get("payment_smart_data")
    kiot_data = st.session_state.get("kiot_data")
    payment_reconciliation = st.session_state.get("payment_reconciliation")
    if payment_smart_data and kiot_data and payment_reconciliation:
        st.caption(
            f"Nhà cung cấp Smart: {payment_smart_data.supplier_name or 'Không xác định'} · "
            f"Tên NCC KIOT (tham khảo): {kiot_data.supplier_name or 'Không xác định'} · "
            f"Kỳ KIOT: {kiot_data.period.label} · Từ khóa: {kiot_data.keyword}"
        )
        statuses = Counter(result.status for result in payment_reconciliation.results)
        metric_columns = st.columns(5)
        metric_columns[0].metric("Tổng kết quả", len(payment_reconciliation.results))
        metric_columns[1].metric("Khớp", statuses.get("KHỚP", 0))
        metric_columns[2].metric("Smart thiếu", statuses.get("SMART THIẾU", 0))
        metric_columns[3].metric("KIOT thiếu", statuses.get("KIOT THIẾU", 0))
        metric_columns[4].metric("Cần kiểm tra", statuses.get("CẦN KIỂM TRA", 0))
        smart_payment_total = sum(record.amount for record in payment_smart_data.payments)
        kiot_payment_total = sum(record.amount for record in kiot_data.payments)
        total_columns = st.columns(3)
        total_columns[0].metric("Đã thanh toán Smart", f"{smart_payment_total:,.0f} đ")
        total_columns[1].metric("Ghi có KIOT có từ khóa", f"{kiot_payment_total:,.0f} đ")
        total_columns[2].metric(
            "Chênh lệch tổng",
            f"{smart_payment_total - kiot_payment_total:+,.0f} đ",
        )
        render_issues(payment_reconciliation.issues)
        payment_df = payment_results_to_dataframe(payment_reconciliation)
        render_filtered_table(payment_df, "payment")

        report_bytes = st.session_state.get("payment_report_bytes") or build_excel_report(
            None, payment_smart_data, None, kiot_data, payment_reconciliation
        )
        end_date = (
            payment_smart_data.period.end.strftime("%Y-%m-%d")
            if payment_smart_data.period.end
            else "KET_QUA"
        )
        filename = st.session_state.get("payment_report_filename") or (
            f"Ket_qua_Smart_KIOT_"
            f"{safe_supplier_filename(payment_smart_data.supplier_name)}_{end_date}.xlsx"
        )
        st.download_button(
            "Tải báo cáo Smart–KIOT",
            data=report_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with tab_bank:
    st.markdown(
        """
        <div class="section-kicker" style="margin-top:1.15rem">Đối chiếu tiền vào</div>
        <div class="section-title">Đối chiếu giao dịch ngân hàng với hóa đơn Smart</div>
        <div class="section-description">Chỉ đối chiếu ngày và số tiền. Smart được phép hạch toán cùng ngày hoặc trễ tối đa một ngày; nội dung và tên khách hàng chỉ hiển thị tham khảo.</div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(
        "Thiết lập hiện tại: gom Smart theo ID Chứng từ · tìm trong toàn bộ chứng từ HDBR · "
        "chỉ xuất một dòng cho mỗi giao dịch tiền vào ngân hàng · không bắt buộc khớp tên hoặc nội dung."
    )
    with st.container(border=True):
        bank_upload_col, sales_smart_upload_col = st.columns(2)
        with bank_upload_col:
            st.markdown(
                '<div class="upload-heading">01 · Lịch sử giao dịch ngân hàng</div>'
                '<div class="upload-note">Chọn file có ngày giao dịch, số tiền có dấu vào/ra và nội dung chuyển khoản.</div>',
                unsafe_allow_html=True,
            )
            bank_history_upload = st.file_uploader(
                "Tải file lịch sử giao dịch ngân hàng",
                type=["xlsx"],
                accept_multiple_files=False,
                key="bank_history_upload",
            )
        with sales_smart_upload_col:
            st.markdown(
                '<div class="upload-heading">02 · Smart bán hàng</div>'
                '<div class="upload-note">Chọn báo cáo có ID chứng từ, số HĐ, ngày HĐ và Tổng Tiền.</div>',
                unsafe_allow_html=True,
            )
            sales_smart_upload = st.file_uploader(
                "Tải file Smart bán hàng",
                type=["xlsx"],
                accept_multiple_files=False,
                key="sales_smart_upload",
            )

    bank_smart_files = (
        ([bank_history_upload] if bank_history_upload else [])
        + ([sales_smart_upload] if sales_smart_upload else [])
    )
    bank_smart_signature = file_signature(bank_smart_files) if bank_smart_files else ""
    reset_result_if_inputs_changed(
        "bank_smart_input_signature",
        bank_smart_signature,
        [
            "bank_history_data",
            "sales_smart_data",
            "bank_smart_reconciliation",
            "bank_report_bytes",
            "bank_report_filename",
        ],
    )

    missing_bank_files = []
    if bank_history_upload is None:
        missing_bank_files.append("lịch sử giao dịch ngân hàng")
    if sales_smart_upload is None:
        missing_bank_files.append("Smart bán hàng")
    if missing_bank_files:
        st.info("Hãy chọn file " + " và ".join(missing_bank_files) + " để bắt đầu.")

    run_bank_smart = st.button(
        "Đối chiếu Ngân hàng–Smart",
        type="primary",
        disabled=bool(missing_bank_files),
        use_container_width=True,
    )
    if run_bank_smart:
        try:
            with st.spinner("Đang đọc giao dịch tiền vào, gom hóa đơn Smart và đối chiếu..."):
                bank_history_data = read_bank_history_file(
                    uploaded_to_source(bank_history_upload)
                )
                sales_smart_data = read_sales_smart_file(
                    uploaded_to_source(sales_smart_upload)
                )
                bank_smart_reconciliation = reconcile_bank_smart(
                    bank_history_data,
                    sales_smart_data,
                )
                st.session_state.bank_history_data = bank_history_data
                st.session_state.sales_smart_data = sales_smart_data
                st.session_state.bank_smart_reconciliation = bank_smart_reconciliation
                st.session_state.bank_smart_input_signature = bank_smart_signature
                bank_end_date = (
                    bank_history_data.period.end.strftime("%Y-%m-%d")
                    if bank_history_data.period.end
                    else "KET_QUA"
                )
                bank_report_filename = f"Ket_qua_Ngan_hang_Smart_{bank_end_date}.xlsx"
                bank_report_bytes = build_bank_smart_excel_report(
                    bank_history_data,
                    sales_smart_data,
                    bank_smart_reconciliation,
                )
                bank_statuses = Counter(
                    result.status for result in bank_smart_reconciliation.results
                )
                bank_total = sum(
                    record.amount for record in bank_history_data.transactions
                )
                paired_results = [
                    result
                    for result in bank_smart_reconciliation.results
                    if result.smart is not None
                ]
                paired_smart_total = sum(result.smart.amount for result in paired_results)
                save_history_snapshot(
                    reconciliation_type="BANK_SMART",
                    title="Ngân hàng ↔ Smart bán hàng",
                    input_files=[bank_history_upload.name, sales_smart_upload.name],
                    status_counts=bank_statuses,
                    summary={
                        "Kỳ ngân hàng": bank_history_data.period.label,
                        "Kỳ Smart": sales_smart_data.period.label,
                        "Tổng kết quả": len(bank_smart_reconciliation.results),
                        "Tổng tiền vào ngân hàng": bank_total,
                        "Tổng Smart đã ghép": paired_smart_total,
                    },
                    results=bank_smart_results_to_dataframe(bank_smart_reconciliation),
                    issues=issues_to_dataframe(bank_smart_reconciliation.issues),
                    report_filename=bank_report_filename,
                    report_bytes=bank_report_bytes,
                )
                st.session_state.bank_report_bytes = bank_report_bytes
                st.session_state.bank_report_filename = bank_report_filename
        except Exception as exc:
            st.error(f"Không thể xử lý file ngân hàng hoặc Smart bán hàng: {exc}")

    bank_history_data = st.session_state.get("bank_history_data")
    sales_smart_data = st.session_state.get("sales_smart_data")
    bank_smart_reconciliation = st.session_state.get("bank_smart_reconciliation")
    if bank_history_data and sales_smart_data and bank_smart_reconciliation:
        st.caption(
            f"Kỳ ngân hàng: {bank_history_data.period.label} · "
            f"Kỳ Smart: {sales_smart_data.period.label} · "
            f"Smart đã gom: {len(sales_smart_data.invoices):,} hóa đơn"
        )
        statuses = Counter(result.status for result in bank_smart_reconciliation.results)
        not_matched = sum(
            statuses.get(status, 0)
            for status in (
                "KHÔNG KHỚP SỐ TIỀN",
                "KHÔNG KHỚP NGÀY",
                "THIẾU NỘI DUNG CHUYỂN KHOẢN",
                "CẦN KIỂM TRA",
            )
        )
        metric_columns = st.columns(5)
        metric_columns[0].metric("Giao dịch tiền vào", len(bank_history_data.transactions))
        metric_columns[1].metric("Khớp", statuses.get("KHỚP", 0))
        metric_columns[2].metric("Lệch 1 đồng", statuses.get("CẢNH BÁO LỆCH 1 ĐỒNG", 0))
        metric_columns[3].metric("Smart trùng", statuses.get("SMART TRÙNG HÓA ĐƠN", 0))
        metric_columns[4].metric("Khác cần xử lý", not_matched)

        bank_total = sum(record.amount for record in bank_history_data.transactions)
        paired_results = [
            result for result in bank_smart_reconciliation.results if result.smart is not None
        ]
        paired_smart_total = sum(result.smart.amount for result in paired_results)
        paired_bank_total = sum(result.bank.amount for result in paired_results)
        total_columns = st.columns(3)
        total_columns[0].metric("Tổng tiền vào ngân hàng", f"{bank_total:,.0f} đ")
        total_columns[1].metric("Tổng Smart đã ghép duy nhất", f"{paired_smart_total:,.0f} đ")
        total_columns[2].metric(
            "Chênh lệch các cặp đã ghép",
            f"{paired_smart_total - paired_bank_total:+,.0f} đ",
        )
        render_issues(bank_smart_reconciliation.issues, "Kiểm tra file Ngân hàng–Smart")
        bank_smart_df = bank_smart_results_to_dataframe(bank_smart_reconciliation)
        render_filtered_table(bank_smart_df, "bank_smart")
        report_bytes = st.session_state.get("bank_report_bytes") or (
            build_bank_smart_excel_report(
                bank_history_data,
                sales_smart_data,
                bank_smart_reconciliation,
            )
        )
        end_date = (
            bank_history_data.period.end.strftime("%Y-%m-%d")
            if bank_history_data.period.end
            else "KET_QUA"
        )
        st.download_button(
            "Tải báo cáo Ngân hàng–Smart",
            data=report_bytes,
            file_name=st.session_state.get("bank_report_filename")
            or f"Ket_qua_Ngan_hang_Smart_{end_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

st.divider()
st.markdown(
    """
    <div class="app-footer">
      <span>Tool đối chiếu công nợ nhà cung cấp</span>
      <span>Xử lý cục bộ · Không thay đổi file nguồn · Không gửi dữ liệu ra Internet</span>
    </div>
    """,
    unsafe_allow_html=True,
)
