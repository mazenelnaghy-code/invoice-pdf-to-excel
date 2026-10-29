"""Extract a folder of PDF invoices into one Excel workbook, and check every invoice.

Usage:
    python extract.py data/invoices
    python extract.py path/to/pdfs --out results

Writes <out>/invoices.xlsx with three sheets:
    Invoices      one row per invoice: supplier, number, date, subtotal, VAT, total, status
    Line items    one row per product or service on every invoice
    Needs review  every invoice that failed a check, and why

Each invoice is checked: quantity x unit price = line amount, the lines add up to the subtotal, and
subtotal + VAT = total. Nothing is guessed: an invoice that fails a check is still extracted, and flagged.
"""
import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber

MONEY = re.compile(r"-?\d[\d,]*\.\d{2}")
LINE_ITEM = re.compile(r"^(?P<description>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s+(?:EGP\s*)?(?P<unit_price>\d[\d,]*\.\d{2})"
                       r"\s+(?:EGP\s*)?(?P<amount>\d[\d,]*\.\d{2})$")
INVOICE_NUMBER = re.compile(r"\b(?:invoice|bill)\s*(?:no\.?|number|#)\s*[:#]?\s*([A-Z0-9][A-Z0-9/\-]*)", re.I)
DATE_LABEL = re.compile(r"\b(?:invoice date|issue date|date|dated)\b\s*:?\s*(.+)", re.I)
DATE_PATTERNS = [  # (regex, strptime format)
    (r"\d{4}-\d{2}-\d{2}", "%Y-%m-%d"),
    (r"\d{1,2}/\d{1,2}/\d{4}", "%d/%m/%Y"),  # slashes are day-first, as on Egyptian invoices
    (r"\d{1,2}-\d{1,2}-\d{4}", "%d-%m-%Y"),
    (r"\d{1,2} [A-Za-z]{3} \d{4}", "%d %b %Y"),
    (r"\d{1,2} [A-Za-z]{4,9} \d{4}", "%d %B %Y"),
    (r"[A-Za-z]{3} \d{1,2}, \d{4}", "%b %d, %Y"),
    (r"[A-Za-z]{4,9} \d{1,2}, \d{4}", "%B %d, %Y"),
]
TOTAL_LABELS = {  # field -> labels a line can start with
    "subtotal": re.compile(r"^(?:sub-?\s?total|net amount|net total)\b", re.I),
    "vat": re.compile(r"^(?:vat|tax)\b", re.I),
    "total": re.compile(r"^(?:grand total|total due|amount due|total)\b", re.I),
}
TOLERANCE = 0.015  # money is compared to the cent


def parse_money(text):
    """'EGP 1,250.00' -> 1250.0; None when there is no amount."""
    found = MONEY.findall(text)
    return float(found[-1].replace(",", "")) if found else None


def parse_date(text):
    """First date in the text, in any of DATE_PATTERNS, or None."""
    for pattern, fmt in DATE_PATTERNS:
        for match in re.finditer(pattern, text):
            try:
                return datetime.strptime(match.group(), fmt).date()
            except ValueError:
                continue
    return None


def find_supplier(page):
    """The supplier is the biggest text on the page that isn't the word INVOICE."""
    lines = {}
    for word in page.extract_words(extra_attrs=["size"]):
        key = (round(word["top"]), round(word["size"], 1))
        lines.setdefault(key, []).append(word["text"])
    candidates = [(size, " ".join(words)) for (_, size), words in lines.items()
                  if "invoice" not in " ".join(words).lower()]
    return max(candidates)[1] if candidates else None


def read_invoice(path):
    """Return a dict with the invoice fields, its line items and a list of problems found."""
    invoice = {"file": path.name, "supplier": None, "invoice_no": None, "date": None,
               "subtotal": None, "vat": None, "total": None, "items": [], "problems": []}
    with pdfplumber.open(path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        if not text.strip():
            invoice["problems"].append("No text layer (scanned image): needs OCR or manual entry")
            return invoice
        invoice["supplier"] = find_supplier(pdf.pages[0])

    for line in (l.strip() for l in text.splitlines()):
        if invoice["invoice_no"] is None and (m := INVOICE_NUMBER.search(line)):
            invoice["invoice_no"] = m.group(1)
        if invoice["date"] is None and (m := DATE_LABEL.search(line)):
            invoice["date"] = parse_date(m.group(1))
        if m := LINE_ITEM.match(line):
            invoice["items"].append({"description": m["description"], "qty": float(m["qty"]),
                                     "unit_price": parse_money(m["unit_price"]), "amount": parse_money(m["amount"])})
            continue
        for field, label in TOTAL_LABELS.items():
            if invoice[field] is None and label.match(line):
                invoice[field] = parse_money(line)
                break

    invoice["problems"] += check(invoice)
    return invoice


def check(invoice):
    """Every reason not to trust this invoice as it stands."""
    problems = [f"Missing {name}" for name in ("invoice_no", "date", "subtotal", "vat", "total") if invoice[name] is None]
    if not invoice["items"]:
        problems.append("No line items found")
    for item in invoice["items"]:
        if abs(item["qty"] * item["unit_price"] - item["amount"]) > TOLERANCE:
            problems.append(f"Line '{item['description']}': {item['qty']:g} x {item['unit_price']:,.2f} "
                            f"is not {item['amount']:,.2f}")
    if invoice["subtotal"] is not None and invoice["items"]:
        lines_sum = sum(i["amount"] for i in invoice["items"])
        if abs(lines_sum - invoice["subtotal"]) > TOLERANCE:
            problems.append(f"Lines add up to {lines_sum:,.2f}, but the subtotal says {invoice['subtotal']:,.2f}")
    if None not in (invoice["subtotal"], invoice["vat"], invoice["total"]):
        expected = invoice["subtotal"] + invoice["vat"]
        if abs(expected - invoice["total"]) > TOLERANCE:
            problems.append(f"Subtotal + VAT = {expected:,.2f}, but the total says {invoice['total']:,.2f} "
                            f"({invoice['total'] - expected:+,.2f})")
    return problems


def to_tables(invoices):
    """The three sheets as DataFrames."""
    summary = pd.DataFrame([{
        "file": inv["file"], "supplier": inv["supplier"], "invoice_no": inv["invoice_no"], "date": inv["date"],
        "items": len(inv["items"]), "subtotal": inv["subtotal"], "vat": inv["vat"],
        "vat_rate": round(inv["vat"] / inv["subtotal"], 4) if inv["vat"] and inv["subtotal"] else None,
        "total": inv["total"], "status": "Needs review" if inv["problems"] else "OK",
    } for inv in invoices])
    items = pd.DataFrame([{"file": inv["file"], "invoice_no": inv["invoice_no"], **item}
                          for inv in invoices for item in inv["items"]],
                         columns=["file", "invoice_no", "description", "qty", "unit_price", "amount"])
    review = pd.DataFrame([{"file": inv["file"], "invoice_no": inv["invoice_no"], "problem": problem}
                           for inv in invoices for problem in inv["problems"]],
                          columns=["file", "invoice_no", "problem"])
    return summary, items, review


def write_workbook(tables, path):
    money_cols = {"subtotal", "vat", "total", "unit_price", "amount"}
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        for sheet, df in zip(("Invoices", "Line items", "Needs review"), tables):
            df.to_excel(xl, sheet_name=sheet, index=False)
            ws = xl.sheets[sheet]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col, name in zip(ws.columns, df.columns):
                width = max([len(str(name))] + [len(str(v)) for v in df[name].head(200)]) + 2
                ws.column_dimensions[col[0].column_letter].width = min(width, 70)
                for cell in col[1:]:
                    if name in money_cols:
                        cell.number_format = "#,##0.00"
                    elif name == "vat_rate":
                        cell.number_format = "0%"
                    elif name == "date":
                        cell.number_format = "yyyy-mm-dd"


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("folder", type=Path, help="folder of PDF invoices")
    parser.add_argument("--out", type=Path, default=Path("output"))
    args = parser.parse_args(argv)

    pdfs = sorted(args.folder.glob("*.pdf"))
    if not pdfs:
        sys.exit(f"No PDF files in {args.folder}")
    invoices = []
    for pdf in pdfs:
        inv = read_invoice(pdf)
        invoices.append(inv)
        status = "OK" if not inv["problems"] else "NEEDS REVIEW"
        total = f"{inv['total']:>12,.2f}" if inv["total"] is not None else " " * 12
        print(f"{pdf.name:<28} {inv['supplier'] or '-':<22} {total}  {status}")

    args.out.mkdir(parents=True, exist_ok=True)
    summary, items, review = to_tables(invoices)
    write_workbook((summary, items, review), args.out / "invoices.xlsx")
    ok = (summary["status"] == "OK").sum()
    print(f"\n{len(invoices)} invoices, {len(items)} line items, {ok} OK, {len(invoices) - ok} need review"
          f" -> {args.out / 'invoices.xlsx'}")


if __name__ == "__main__":
    main()
