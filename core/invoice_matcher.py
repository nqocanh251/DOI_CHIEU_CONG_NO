from __future__ import annotations

from collections import defaultdict
from datetime import date

from core.excel_reader import expected_months, hddt_covered_months
from core.models import (
    HddtData,
    InvoiceMatchResult,
    InvoiceReconciliation,
    InvoiceRecord,
    ReportPeriod,
    Severity,
    SmartData,
    ValidationIssue,
)
from core.normalization import is_advisory_invoice_status, is_normal_invoice_status


def _record_sort_key(record: InvoiceRecord):
    return record.invoice_date or record.source_row, record.invoice_number, record.source_row


def _reason_with_hddt_warning(reason: str, record: InvoiceRecord | None) -> str:
    if record and is_advisory_invoice_status(record.status):
        return f"{reason}; cảnh báo trạng thái HĐĐT: {record.status}"
    return reason


def _expected_hddt_months(smart_period: ReportPeriod, as_of_date: date | None) -> set[tuple[int, int]]:
    """Mở kỳ của năm hiện hành tới tháng đang chạy tool, không khóa ở kỳ Smart cũ."""
    if not smart_period.start or not smart_period.end:
        return expected_months(smart_period)
    today = as_of_date or date.today()
    expected_end = smart_period.end
    if smart_period.end.year == today.year and today > smart_period.end:
        expected_end = today
    return expected_months(ReportPeriod(smart_period.start, expected_end))


def reconcile_invoices(
    hddt: HddtData,
    smart: SmartData,
    as_of_date: date | None = None,
) -> InvoiceReconciliation:
    issues = list(hddt.issues) + list(smart.issues)

    expected = _expected_hddt_months(smart.period, as_of_date)
    covered = hddt_covered_months(hddt)
    missing = sorted(expected - covered)
    if expected and missing:
        label = ", ".join(f"{month:02d}/{year}" for year, month in missing)
        issues.append(
            ValidationIssue(
                Severity.ERROR,
                "MISSING_HDDT_MONTHS",
                f"Thiếu file HĐĐT của các tháng: {label}",
            )
        )

    supplier_tax_id = smart.supplier_tax_id
    relevant_hddt = [
        record for record in hddt.invoices if supplier_tax_id and record.supplier_tax_id == supplier_tax_id
    ]
    if not supplier_tax_id:
        relevant_hddt = []
    if supplier_tax_id and not relevant_hddt:
        issues.append(
            ValidationIssue(
                Severity.ERROR,
                "NO_HDDT_FOR_SMART_SUPPLIER",
                f"Không có hóa đơn HĐĐT nào thuộc MST nhà cung cấp {supplier_tax_id}",
            )
        )

    for record in relevant_hddt:
        if is_advisory_invoice_status(record.status):
            issues.append(
                ValidationIssue(
                    Severity.WARNING,
                    "HDDT_SPECIAL_STATUS",
                    f"Hóa đơn {record.invoice_number_raw or record.invoice_number}: {record.status}",
                    record.source_file,
                    record.sheet,
                    record.source_row,
                )
            )

    hddt_index: dict[tuple[str, str], list[InvoiceRecord]] = defaultdict(list)
    smart_index: dict[tuple[str, str], list[InvoiceRecord]] = defaultdict(list)
    for record in relevant_hddt:
        hddt_index[record.match_key].append(record)
    for record in smart.invoices:
        smart_index[record.match_key].append(record)

    results: list[InvoiceMatchResult] = []
    for key in sorted(set(hddt_index) | set(smart_index)):
        h_records = sorted(hddt_index.get(key, []), key=_record_sort_key)
        s_records = sorted(smart_index.get(key, []), key=_record_sort_key)

        if len(h_records) > 1 or len(s_records) > 1:
            results.append(
                InvoiceMatchResult(
                    status="TRÙNG HÓA ĐƠN",
                    reason=(
                        f"Số hóa đơn xuất hiện {len(h_records)} lần trong HĐĐT và "
                        f"{len(s_records)} lần trong Smart"
                    ),
                    hddt=h_records[0] if h_records else None,
                    smart=s_records[0] if s_records else None,
                )
            )
            continue

        h_record = h_records[0] if h_records else None
        s_record = s_records[0] if s_records else None

        if (
            h_record
            and not is_normal_invoice_status(h_record.status)
            and not is_advisory_invoice_status(h_record.status)
        ):
            details = [f"Trạng thái HĐĐT: {h_record.status}"]
            if s_record is None:
                details.append("không tìm thấy hóa đơn tương ứng trong Smart")
            elif h_record.amount != s_record.amount:
                details.append(f"chênh lệch tiền {h_record.amount - s_record.amount:+,} đồng")
            results.append(
                InvoiceMatchResult(
                    status=h_record.status,
                    reason="; ".join(details),
                    hddt=h_record,
                    smart=s_record,
                )
            )
            continue

        if h_record is None:
            results.append(
                InvoiceMatchResult(
                    status="HĐĐT THIẾU HÓA ĐƠN",
                    reason="Có trong Smart nhưng không tìm thấy trong các file HĐĐT đã tải",
                    smart=s_record,
                )
            )
        elif s_record is None:
            results.append(
                InvoiceMatchResult(
                    status="SMART THIẾU HÓA ĐƠN",
                    reason=_reason_with_hddt_warning(
                        "Có trong HĐĐT nhưng không tìm thấy trong Smart",
                        h_record,
                    ),
                    hddt=h_record,
                )
            )
        elif h_record.amount != s_record.amount:
            difference = h_record.amount - s_record.amount
            results.append(
                InvoiceMatchResult(
                    status="LỆCH TIỀN",
                    reason=_reason_with_hddt_warning(
                        f"HĐĐT và Smart chênh lệch {difference:+,} đồng",
                        h_record,
                    ),
                    hddt=h_record,
                    smart=s_record,
                )
            )
        else:
            results.append(
                InvoiceMatchResult(
                    status="KHỚP",
                    reason=_reason_with_hddt_warning(
                        "Khớp số hóa đơn và tổng tiền",
                        h_record,
                    ),
                    hddt=h_record,
                    smart=s_record,
                )
            )

    if not results:
        issues.append(
            ValidationIssue(
                Severity.ERROR,
                "NO_INVOICE_RESULTS",
                "Không có hóa đơn hợp lệ để đối chiếu HĐĐT–Smart",
            )
        )
    return InvoiceReconciliation(results=results, issues=issues)
