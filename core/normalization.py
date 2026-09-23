from __future__ import annotations

import math
import re
import unicodedata
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


_SPACE_RE = re.compile(r"\s+")
_DATE_RE = re.compile(r"(?P<day>\d{1,2})/(?P<month>\d{1,2})/(?P<year>\d{2,4})")
_TIME_RE = re.compile(r"(?P<hour>\d{1,2}):(?P<minute>\d{2})(?::(?P<second>\d{2}))?")


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return _SPACE_RE.sub(" ", str(value).replace("\u00a0", " ")).strip()


def fold_text(value: object) -> str:
    text = clean_text(value).upper().replace("Đ", "D")
    decomposed = unicodedata.normalize("NFD", text)
    return "".join(char for char in decomposed if unicodedata.category(char) != "Mn")


def normalize_tax_id(value: object) -> str:
    return re.sub(r"\D", "", clean_text(value))


def normalize_invoice_number(value: object) -> str:
    text = clean_text(value)
    if re.fullmatch(r"[+-]?\d+\.0+", text):
        text = text.split(".", 1)[0]
    text = _SPACE_RE.sub("", text).upper()
    if re.fullmatch(r"\d+", text):
        return text.lstrip("0") or "0"
    return text


def normalize_symbol(value: object) -> str:
    return re.sub(r"\s+", "", clean_text(value)).upper()


def normalize_status(value: object) -> str:
    return clean_text(value).upper()


def is_normal_invoice_status(value: object) -> bool:
    status = fold_text(value)
    return status in {"HOA DON MOI", "MOI", "HOP LE"}


def is_advisory_invoice_status(value: object) -> bool:
    """Các trạng thái vẫn được đối chiếu nhưng phải hiện cảnh báo cho kế toán."""
    status = fold_text(value)
    return status in {
        "HOA DON THAY THE",
        "HOA DON BI THAY THE",
        "HOA DON DA BI THAY THE",
        "HOA DON DIEU CHINH",
        "HOA DON BI DIEU CHINH",
        "HOA DON DA BI DIEU CHINH",
    }


def money_to_int(value: object) -> int:
    if value is None or value == "":
        raise ValueError("Số tiền đang để trống")
    if isinstance(value, bool):
        raise ValueError("Giá trị logic không phải số tiền")
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError("Số tiền không hữu hạn")
        number = Decimal(str(value))
    else:
        text = clean_text(value).replace(" ", "")
        if not text:
            raise ValueError("Số tiền đang để trống")
        if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", text):
            text = text.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", text):
            text = text.replace(",", "")
        elif "," in text and "." not in text:
            text = text.replace(",", ".")
        try:
            number = Decimal(text)
        except InvalidOperation as exc:
            raise ValueError(f"Không đọc được số tiền: {value}") from exc
    rounded = number.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    if number != rounded:
        raise ValueError(f"Số tiền không phải số nguyên VND: {value}")
    return int(rounded)


def parse_date(value: object) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    match = _DATE_RE.search(clean_text(value))
    if not match:
        return None
    year = int(match.group("year"))
    if year < 100:
        year += 2000
    try:
        return date(year, int(match.group("month")), int(match.group("day")))
    except ValueError:
        return None


def parse_time(value: object) -> time | None:
    if isinstance(value, datetime):
        return value.time().replace(microsecond=0)
    if isinstance(value, time):
        return value.replace(microsecond=0)
    match = _TIME_RE.search(clean_text(value))
    if not match:
        return None
    try:
        return time(
            int(match.group("hour")),
            int(match.group("minute")),
            int(match.group("second") or 0),
        )
    except ValueError:
        return None


def parse_report_period(value: object) -> tuple[date | None, date | None]:
    text = clean_text(value)
    matches = list(_DATE_RE.finditer(text))
    if len(matches) < 2:
        return None, None
    return parse_date(matches[0].group(0)), parse_date(matches[1].group(0))


def format_date(value: date | None) -> str:
    return value.strftime("%d/%m/%Y") if value else ""


def format_time(value: time | None) -> str:
    return value.strftime("%H:%M:%S").removesuffix(":00") if value else ""
