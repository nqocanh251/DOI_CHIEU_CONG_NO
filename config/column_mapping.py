DEFAULT_KEYWORD = "KHUYẾN NÔNG"
MAX_KIOT_DELAY_DAYS = 2
MAX_BANK_SMART_DELAY_DAYS = 1
MAX_BANK_SMART_AMOUNT_DIFFERENCE = 1

HDDT_REQUIRED_HEADERS = {
    "invoice_number": ("SO HOA DON",),
    "supplier_tax_id": ("MST NGUOI BAN", "MST NGUOI BAN/MST NGUOI XUAT HANG"),
    "supplier_name": ("TEN NGUOI BAN", "TEN NGUOI BAN/TEN NGUOI XUAT HANG"),
    "invoice_symbol": ("KY HIEU HOA DON",),
    "invoice_date": ("NGAY LAP",),
    "amount": ("TONG TIEN THANH TOAN",),
    "status": ("TRANG THAI HOA DON",),
}

KIOT_REQUIRED_HEADERS = {
    "datetime": ("THOI GIAN",),
    "description": ("DIEN GIAI",),
    "debit": ("GHI NO",),
    "credit": ("GHI CO",),
}

BANK_HISTORY_HEADERS = {
    "transaction_datetime": ("THOI GIAN GIAO DICH", "NGAY GIAO DICH"),
    "posting_date": ("NGAY HACH TOAN",),
    "amount": ("SO TIEN GIAO DICH", "SO TIEN", "GIA TRI GIAO DICH"),
    "credit": ("GHI CO", "TIEN VAO", "CREDIT"),
    "content": ("NOI DUNG", "DIEN GIAI", "MO TA"),
    "transaction_id": ("SO ID", "MA GIAO DICH", "ID GIAO DICH"),
    "counterparty_account": ("TAI KHOAN DOI UNG", "TK DOI UNG"),
    "currency": ("LOAI TIEN", "DON VI TIEN TE"),
}

SALES_SMART_HEADERS = {
    "document_id": ("ID CHUNG TU", "MA CHUNG TU"),
    "document_type": ("CHUNG TU", "LOAI CHUNG TU"),
    "invoice_number": ("SO HD", "SO HOA DON"),
    "invoice_date": ("NGAY HD", "NGAY HOA DON"),
    "document_date": ("NGAY C.TU", "NGAY CHUNG TU"),
    "total_amount": ("TONG TIEN", "TONG TIEN VND"),
    "customer_name": ("TEN KHACH HANG", "TEN KH", "KHACH HANG"),
    "debit_account": ("TK NO", "TAI KHOAN NO"),
}
