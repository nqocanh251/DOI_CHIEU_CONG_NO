from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, time
from enum import Enum
from typing import Iterable


class Severity(str, Enum):
    WARNING = "CẢNH BÁO"
    ERROR = "LỖI"


@dataclass(frozen=True)
class ExcelFile:
    name: str
    content: bytes

    @classmethod
    def from_path(cls, path: str) -> "ExcelFile":
        from pathlib import Path

        source = Path(path)
        return cls(name=source.name, content=source.read_bytes())


@dataclass(frozen=True)
class ValidationIssue:
    severity: Severity
    code: str
    message: str
    source_file: str = ""
    sheet: str = ""
    row: int | None = None

    @property
    def is_error(self) -> bool:
        return self.severity is Severity.ERROR


@dataclass(frozen=True)
class ReportPeriod:
    start: date | None = None
    end: date | None = None

    @property
    def label(self) -> str:
        if self.start and self.end:
            return f"{self.start:%d/%m/%Y} – {self.end:%d/%m/%Y}"
        return "Không xác định"


@dataclass(frozen=True)
class InvoiceRecord:
    supplier_tax_id: str
    supplier_name: str
    invoice_number: str
    invoice_number_raw: str
    invoice_symbol: str
    invoice_date: date | None
    amount: int
    status: str
    source_file: str
    sheet: str
    source_row: int
    reference: str = ""

    @property
    def match_key(self) -> tuple[str, str]:
        # Ký hiệu và tên doanh nghiệp chỉ dùng để tham khảo vì có thể trùng
        # giữa nhiều công ty. MST + số hóa đơn chuẩn hóa là khóa đối chiếu.
        return self.supplier_tax_id, self.invoice_number


@dataclass(frozen=True)
class PaymentRecord:
    source: str
    payment_date: date
    amount: int
    content: str
    source_file: str
    sheet: str
    source_row: int
    payment_time: time | None = None
    reference: str = ""


@dataclass(frozen=True)
class BankTransactionRecord:
    transaction_date: date
    amount: int
    content: str
    transaction_id: str
    source_file: str
    sheet: str
    source_row: int
    counterparty_account: str = ""


@dataclass(frozen=True)
class SalesInvoiceRecord:
    document_id: str
    invoice_number: str
    invoice_date: date
    amount: int
    customer_name: str
    debit_account: str
    source_file: str
    sheet: str
    source_rows: tuple[int, ...]


@dataclass
class HddtData:
    invoices: list[InvoiceRecord] = field(default_factory=list)
    periods: list[ReportPeriod] = field(default_factory=list)
    issues: list[ValidationIssue] = field(default_factory=list)
    files: list[str] = field(default_factory=list)


@dataclass
class SmartData:
    supplier_tax_id: str = ""
    supplier_name: str = ""
    invoices: list[InvoiceRecord] = field(default_factory=list)
    payments: list[PaymentRecord] = field(default_factory=list)
    period: ReportPeriod = field(default_factory=ReportPeriod)
    issues: list[ValidationIssue] = field(default_factory=list)
    source_file: str = ""


@dataclass
class KiotData:
    supplier_name: str = ""
    payments: list[PaymentRecord] = field(default_factory=list)
    period: ReportPeriod = field(default_factory=ReportPeriod)
    issues: list[ValidationIssue] = field(default_factory=list)
    source_file: str = ""
    keyword: str = "KHUYẾN NÔNG"


@dataclass
class BankHistoryData:
    transactions: list[BankTransactionRecord] = field(default_factory=list)
    period: ReportPeriod = field(default_factory=ReportPeriod)
    issues: list[ValidationIssue] = field(default_factory=list)
    source_file: str = ""


@dataclass
class SalesSmartData:
    invoices: list[SalesInvoiceRecord] = field(default_factory=list)
    period: ReportPeriod = field(default_factory=ReportPeriod)
    issues: list[ValidationIssue] = field(default_factory=list)
    source_file: str = ""


@dataclass(frozen=True)
class InvoiceMatchResult:
    status: str
    reason: str
    hddt: InvoiceRecord | None = None
    smart: InvoiceRecord | None = None


@dataclass
class InvoiceReconciliation:
    results: list[InvoiceMatchResult]
    issues: list[ValidationIssue] = field(default_factory=list)

    @property
    def is_fully_matched(self) -> bool:
        return bool(self.results) and all(r.status == "KHỚP" for r in self.results) and not any(
            issue.is_error for issue in self.issues
        )

    @property
    def can_continue(self) -> bool:
        """Tương thích báo cáo cũ; không còn được dùng để khóa Smart–KIOT."""
        return self.is_fully_matched


@dataclass(frozen=True)
class PaymentMatchResult:
    status: str
    reason: str
    smart: PaymentRecord | None = None
    kiot: PaymentRecord | None = None
    day_difference: int | None = None


@dataclass
class PaymentReconciliation:
    results: list[PaymentMatchResult]
    issues: list[ValidationIssue] = field(default_factory=list)


@dataclass(frozen=True)
class BankSmartMatchResult:
    status: str
    reason: str
    bank: BankTransactionRecord
    smart: SalesInvoiceRecord | None = None
    amount_difference: int | None = None
    day_difference: int | None = None
    name_matched: bool | None = None
    candidate_count: int = 0
    candidate_invoice_numbers: tuple[str, ...] = ()
    candidate_customer_names: tuple[str, ...] = ()


@dataclass
class BankSmartReconciliation:
    results: list[BankSmartMatchResult]
    issues: list[ValidationIssue] = field(default_factory=list)


def has_errors(issues: Iterable[ValidationIssue]) -> bool:
    return any(issue.is_error for issue in issues)
