import re
from datetime import datetime
import xlrd


def extract_last4(workbook: xlrd.Book, statement_type: str) -> str:
    sheet = workbook.sheet_by_index(0)
    if statement_type == "cc":
        for row_idx in range(min(10, sheet.nrows)):
            for col_idx in range(sheet.ncols):
                cell = str(sheet.cell_value(row_idx, col_idx))
                if "Credit Card No." in cell:
                    match = re.search(r'(\d{4})\s*$', cell.strip())
                    if match:
                        return match.group(1)
    elif statement_type == "bank":
        for row_idx in range(min(20, sheet.nrows)):
            for col_idx in range(sheet.ncols):
                cell = str(sheet.cell_value(row_idx, col_idx))
                if re.search(r'Account No\s*:', cell):
                    match = re.search(r'Account No\s*:\s*(\d+)', cell)
                    if match:
                        return match.group(1)[-4:]
    return "XXXX"


def detect_statement_type(workbook: xlrd.Book) -> str:
    for sheet_name in workbook.sheet_names():
        sheet = workbook.sheet_by_name(sheet_name)
        for row_idx in range(min(30, sheet.nrows)):
            for col_idx in range(sheet.ncols):
                cell = str(sheet.cell_value(row_idx, col_idx))
                if "Credit Card No." in cell:
                    return "cc"
        if sheet.nrows > 0:
            first_cell = str(sheet.cell_value(0, 0))
            if "Statement of accounts" in first_cell or "HDFC BANK" in first_cell:
                for row_idx in range(min(25, sheet.nrows)):
                    row = [str(sheet.cell_value(row_idx, c)) for c in range(sheet.ncols)]
                    if row[0] == "Date" and "Narration" in row[1]:
                        return "bank"
    return "unknown"


def _parse_amount(value) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _extract_ref(description: str) -> str:
    match = re.search(r'Ref#\s*(\S+)', description)
    if not match:
        return ""
    return match.group(1).rstrip(")")


def _clean_cc_payee(description: str) -> str:
    payee = re.sub(r'\s*\(.*?\)', '', description).strip()
    payee = re.sub(r'(\s+[A-Z]{2,}){1,2}$', '', payee).strip()
    return payee or description.strip()


def parse_cc_statement(workbook: xlrd.Book) -> list[dict]:
    sheet = workbook.sheet_by_index(0)
    transactions = []
    header_row = None

    for row_idx in range(sheet.nrows):
        if str(sheet.cell_value(row_idx, 0)).strip() == "Transaction type":
            header_row = row_idx
            break

    if header_row is None:
        raise ValueError("Could not find transaction header row in CC statement.")

    valid_types = {"Domestic", "International"}

    for row_idx in range(header_row + 1, sheet.nrows):
        tx_type = str(sheet.cell_value(row_idx, 0)).strip()
        if tx_type not in valid_types:
            continue

        date_str = str(sheet.cell_value(row_idx, 9)).strip()
        description = str(sheet.cell_value(row_idx, 12)).strip()
        amount_raw = sheet.cell_value(row_idx, 20)
        debit_credit = str(sheet.cell_value(row_idx, 23)).strip()

        try:
            tx_date = datetime.strptime(date_str.split(" / ")[0], "%d/%m/%Y").date()
        except (ValueError, IndexError):
            continue

        amount = _parse_amount(amount_raw)
        if amount == 0.0:
            continue

        is_credit = debit_credit.lower() == "cr"
        payee = _clean_cc_payee(description)
        reference = _extract_ref(description)

        transactions.append({
            "date": tx_date,
            "amount": amount,
            "is_credit": is_credit,
            "payee": payee,
            "description": description,
            "reference": reference,
        })

    return transactions


def _extract_bank_payee(narration: str) -> str:
    narration = narration.strip()

    if narration.startswith("UPI-"):
        parts = narration.split("-")
        if len(parts) >= 2:
            return parts[1].strip()

    if narration.startswith("ATW-"):
        return "ATM Withdrawal"

    if "IB BILLPAY DR-HDFCWI" in narration:
        return "HDFC Credit Card Payment"

    if "BILLPAY" in narration.upper():
        match = re.search(r'BILLPAY\s+DR-(\S+)', narration, re.IGNORECASE)
        if match:
            return match.group(1)
        return "Bill Payment"

    parts = narration.split("-")

    if len(parts) >= 2 and parts[0].isdigit() and len(parts[0]) >= 8:
        return parts[-1].strip()[:60]

    if len(parts) >= 2 and len(parts[0]) <= 10:
        return "-".join(parts[1:]).strip()[:60]

    return narration[:60]


def parse_bank_statement(workbook: xlrd.Book) -> list[dict]:
    sheet = workbook.sheet_by_index(0)
    transactions = []
    header_row = None

    for row_idx in range(sheet.nrows):
        row = [str(sheet.cell_value(row_idx, c)).strip() for c in range(sheet.ncols)]
        if row[0] == "Date" and len(row) > 1 and "Narration" in row[1]:
            header_row = row_idx
            break

    if header_row is None:
        raise ValueError("Could not find transaction header row in bank statement.")

    for row_idx in range(header_row + 1, sheet.nrows):
        date_val = str(sheet.cell_value(row_idx, 0)).strip()

        if not date_val or date_val.startswith("*") or date_val.startswith("-"):
            continue

        try:
            tx_date = datetime.strptime(date_val, "%d/%m/%y").date()
        except ValueError:
            continue

        narration = str(sheet.cell_value(row_idx, 1)).strip()
        reference = str(sheet.cell_value(row_idx, 2)).strip()
        withdrawal_raw = sheet.cell_value(row_idx, 4)
        deposit_raw = sheet.cell_value(row_idx, 5)

        withdrawal = _parse_amount(withdrawal_raw)
        deposit = _parse_amount(deposit_raw)

        if withdrawal == 0.0 and deposit == 0.0:
            continue

        payee = _extract_bank_payee(narration)

        transactions.append({
            "date": tx_date,
            "withdrawal": withdrawal,
            "deposit": deposit,
            "payee": payee,
            "description": narration,
            "reference": reference,
        })

    return transactions
