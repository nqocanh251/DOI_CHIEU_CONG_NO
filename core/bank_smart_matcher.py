from __future__ import annotations

from collections import defaultdict
import re

from config.column_mapping import (
    MAX_BANK_SMART_AMOUNT_DIFFERENCE,
    MAX_BANK_SMART_DELAY_DAYS,
)
from core.models import (
    BankHistoryData,
    BankSmartMatchResult,
    BankSmartReconciliation,
    SalesInvoiceRecord,
    SalesSmartData,
)
from core.normalization import fold_text


_NAME_STOPWORDS = {
    "ANH",
    "BA",
    "BAC",
    "CHI",
    "CHU",
    "CO",
    "EM",
    "ONG",
    "THI",
    "VAN",
    "CUA",
    "HANG",
    "DICH",
    "VU",
    "NONG",
    "NGHIEP",
    "KHUYEN",
    "HO",
    "KINH",
    "DOANH",
    "CONG",
    "TY",
    "TNHH",
    "CP",
    "CTY",
    "DVNN",
    "CHUYEN",
    "TIEN",
    "THANH",
    "TOAN",
    "TTCN",
    "CK",
    "NHANH",
    "QUA",
    "ZALO",
    "THON",
    "AP",
    "XA",
    "PHUONG",
    "HUYEN",
    "TINH",
}


def _tokens(value: object) -> list[str]:
    return re.findall(r"[A-Z0-9]+", fold_text(value))


def _meaningful_name_tokens(value: object) -> list[str]:
    original = _tokens(value)
    meaningful = [
        token
        for token in original
        if token not in _NAME_STOPWORDS and not token.isdigit() and len(token) >= 2
    ]
    return meaningful or [token for token in original if not token.isdigit() and len(token) >= 2]


def customer_name_match_evidence(bank_content: str, customer_name: str) -> tuple[bool, tuple[str, ...]]:
    if not bank_content.strip() or not customer_name.strip():
        return False, ()
    content_tokens = set(_tokens(bank_content))
    customer_tokens = _meaningful_name_tokens(customer_name)
    matched = tuple(dict.fromkeys(token for token in customer_tokens if token in content_tokens))
    if matched:
        return True, matched

    compact_content = "".join(_tokens(bank_content))
    compact_customer = "".join(customer_tokens)
    if len(compact_customer) >= 4 and compact_customer in compact_content:
        return True, (compact_customer,)
    for token in customer_tokens:
        if len(token) >= 3 and token in compact_content:
            return True, (token,)
    return False, ()


def customer_name_matches(bank_content: str, customer_name: str) -> bool:
    return customer_name_match_evidence(bank_content, customer_name)[0]


def _invoice_label(record: SalesInvoiceRecord) -> str:
    return record.invoice_number or record.document_id


def reconcile_bank_smart(
    bank_data: BankHistoryData,
    smart_data: SalesSmartData,
    max_delay_days: int = MAX_BANK_SMART_DELAY_DAYS,
    max_amount_difference: int = MAX_BANK_SMART_AMOUNT_DIFFERENCE,
) -> BankSmartReconciliation:
    issues = list(bank_data.issues) + list(smart_data.issues)
    banks = sorted(
        bank_data.transactions,
        key=lambda record: (record.transaction_date, record.source_row),
    )
    smart = list(smart_data.invoices)

    valid_candidates: dict[int, list[tuple[int, int, int]]] = {}
    for bank_index, bank in enumerate(banks):
        candidates = []
        for smart_index, invoice in enumerate(smart):
            day_difference = (invoice.invoice_date - bank.transaction_date).days
            amount_difference = invoice.amount - bank.amount
            if 0 <= day_difference <= max_delay_days and abs(amount_difference) <= max_amount_difference:
                candidates.append((smart_index, day_difference, amount_difference))
        if candidates:
            best_amount_difference = min(abs(item[2]) for item in candidates)
            candidates = [item for item in candidates if abs(item[2]) == best_amount_difference]
        valid_candidates[bank_index] = candidates

    single_claims: dict[int, list[int]] = defaultdict(list)
    for bank_index, candidates in valid_candidates.items():
        if len(candidates) == 1:
            single_claims[candidates[0][0]].append(bank_index)
    conflicting_banks = {
        bank_index
        for bank_indexes in single_claims.values()
        if len(bank_indexes) > 1
        for bank_index in bank_indexes
    }

    results = []
    for bank_index, bank in enumerate(banks):
        candidates = valid_candidates[bank_index]
        candidate_records = [smart[index] for index, _, _ in candidates]
        candidate_numbers = tuple(_invoice_label(record) for record in candidate_records)
        candidate_names = tuple(record.customer_name for record in candidate_records)

        if bank_index in conflicting_banks:
            invoice = candidate_records[0]
            results.append(
                BankSmartMatchResult(
                    status="CẦN KIỂM TRA",
                    reason="Một hóa đơn Smart đang đồng thời phù hợp với nhiều giao dịch ngân hàng",
                    bank=bank,
                    candidate_count=1,
                    candidate_invoice_numbers=(_invoice_label(invoice),),
                    candidate_customer_names=(invoice.customer_name,),
                )
            )
            continue

        if len(candidates) > 1:
            results.append(
                BankSmartMatchResult(
                    status="SMART TRÙNG HÓA ĐƠN",
                    reason=f"Có {len(candidates)} hóa đơn Smart cùng số tiền và ngày hợp lệ",
                    bank=bank,
                    candidate_count=len(candidates),
                    candidate_invoice_numbers=candidate_numbers,
                    candidate_customer_names=candidate_names,
                )
            )
            continue

        if not candidates:
            amount_candidates = [
                invoice
                for invoice in smart
                if abs(invoice.amount - bank.amount) <= max_amount_difference
            ]
            if amount_candidates:
                closest = min(
                    amount_candidates,
                    key=lambda invoice: abs((invoice.invoice_date - bank.transaction_date).days),
                )
                difference = (closest.invoice_date - bank.transaction_date).days
                status = "KHÔNG KHỚP NGÀY"
                reason = (
                    f"Có hóa đơn Smart cùng số tiền nhưng ngày Smart lệch {difference:+d} ngày; "
                    f"chỉ cho phép từ 0 đến {max_delay_days} ngày"
                )
            else:
                status = "KHÔNG KHỚP SỐ TIỀN"
                reason = "Không tìm thấy hóa đơn Smart có số tiền chênh lệch không quá 1 đồng"
            results.append(BankSmartMatchResult(status=status, reason=reason, bank=bank))
            continue

        smart_index, day_difference, amount_difference = candidates[0]
        invoice = smart[smart_index]
        common_details = (
            f"Smart trễ ngân hàng {day_difference} ngày; chênh lệch tiền {amount_difference:+d} đồng"
        )
        if abs(amount_difference) == 1:
            status = "CẢNH BÁO LỆCH 1 ĐỒNG"
            reason = f"{common_details}; nội dung chuyển khoản và tên khách hàng chỉ để tham khảo"
        else:
            status = "KHỚP"
            reason = f"{common_details}; nội dung chuyển khoản và tên khách hàng chỉ để tham khảo"
        results.append(
            BankSmartMatchResult(
                status=status,
                reason=reason,
                bank=bank,
                smart=invoice,
                amount_difference=amount_difference,
                day_difference=day_difference,
                name_matched=None,
                candidate_count=1,
                candidate_invoice_numbers=(_invoice_label(invoice),),
                candidate_customer_names=(invoice.customer_name,),
            )
        )

    return BankSmartReconciliation(results=results, issues=issues)
