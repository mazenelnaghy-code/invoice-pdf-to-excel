"""Create a folder of sample invoice PDFs to try the extractor on.

Usage:
    python make_sample_invoices.py            # writes data/invoices/*.pdf

Three fictional suppliers, each with its own layout and its own words for the same fields
("Invoice No" / "Invoice #" / "Bill number", "Subtotal" / "Net amount" / "Sub-total" ...).
Two invoices are deliberately bad, the way real ones are: one has a total that doesn't add up,
and one is a scan with no text layer.
"""
import random
from datetime import date, timedelta
from pathlib import Path

import pdfplumber
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

OUT = Path("data/invoices")
VAT = 0.14
WIDTH, HEIGHT = A4
NAVY = colors.HexColor("#14213D")
AMBER = colors.HexColor("#E0A040")
GREY = colors.HexColor("#5C6370")

CUSTOMER = ["Bill to:", "Sinai Retail Group", "14 El Gomhoureya St.", "Port Said, Egypt"]

CATALOGUE = {
    "nile": [("A4 copy paper, box of 5 reams", 450), ("Printer toner 85A", 1250), ("Ballpoint pens, pack of 50", 180),
             ("Arch lever files", 35), ("Whiteboard markers, set of 4", 95), ("Desk organiser", 260)],
    "delta": [("Business cards, 500 pcs", 650), ("Roll-up banner 85x200", 1400), ("Flyers A5, 1000 pcs", 900),
              ("Branded notebooks", 120), ("Sticker sheets", 15)],
    "canal": [("Network maintenance visit", 600), ("Laptop repair and cleaning", 450), ("Firewall configuration", 900),
              ("On-site support", 350)],
}


def money(value):
    return f"{value:,.2f}"


def pick_items(rng, supplier):
    """2-5 random lines from the supplier's catalogue as (description, qty, unit_price, amount)."""
    lines = rng.sample(CATALOGUE[supplier], rng.randint(2, min(5, len(CATALOGUE[supplier]))))
    items = []
    for name, price in lines:
        qty = rng.choice([0.5, 1.5, 2.5, 3, 4]) if supplier == "canal" else rng.randint(1, 20)
        items.append((name, qty, float(price), round(qty * price, 2)))
    return items


def totals(items):
    subtotal = round(sum(i[3] for i in items), 2)
    tax = round(subtotal * VAT, 2)
    return subtotal, tax, round(subtotal + tax, 2)


def fmt_qty(qty):
    return f"{qty:g}"


def nile_layout(c, number, day, items, printed_total=None):
    """Grid table, labels on the right: Invoice No / Date dd/mm/yyyy / Subtotal / VAT 14% / Total."""
    subtotal, tax, total = totals(items)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 22)
    c.drawString(50, HEIGHT - 70, "Nile Paper Co.")
    c.setFont("Helvetica", 9)
    c.setFillColor(GREY)
    for i, line in enumerate(["Office supplies wholesale", "27 Corniche Rd., Alexandria", "Tax ID 512-884-301"]):
        c.drawString(50, HEIGHT - 88 - i * 12, line)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 16)
    c.drawRightString(WIDTH - 50, HEIGHT - 120, "INVOICE")
    c.setFont("Helvetica", 10)
    c.drawRightString(WIDTH - 50, HEIGHT - 138, f"Invoice No: {number}")
    c.drawRightString(WIDTH - 50, HEIGHT - 152, f"Date: {day:%d/%m/%Y}")
    draw_customer(c, 50, HEIGHT - 175)

    top = HEIGHT - 260
    cols = [50, 330, 390, 470, WIDTH - 50]
    rows = [("Description", "Qty", "Unit Price", "Amount (EGP)")] + [
        (d, fmt_qty(q), money(p), money(a)) for d, q, p, a in items]
    for r, row in enumerate(rows):
        y = top - r * 22
        if r == 0:
            c.setFillColor(colors.HexColor("#E8ECF3"))
            c.rect(cols[0], y - 6, cols[-1] - cols[0], 22, fill=1, stroke=0)
        c.setFillColor(NAVY)
        c.setFont("Helvetica-Bold" if r == 0 else "Helvetica", 10)
        c.drawString(cols[0] + 6, y + 2, row[0])
        for k in range(1, 4):
            c.drawRightString(cols[k + 1] - 6, y + 2, row[k])
    c.setStrokeColor(colors.HexColor("#B8C0CC"))
    bottom = top - (len(rows) - 1) * 22 - 6
    for x in cols:
        c.line(x, top + 16, x, bottom)
    for r in range(len(rows) + 1):
        c.line(cols[0], top + 16 - r * 22, cols[-1], top + 16 - r * 22)

    y = bottom - 30
    for label, value, bold in [("Subtotal", subtotal, False), ("VAT 14%", tax, False),
                               ("Total", printed_total or total, True)]:
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 11 if bold else 10)
        c.drawString(370, y, label)
        c.drawRightString(WIDTH - 56, y, money(value))
        y -= 18
    footer(c, "Payment within 30 days. Thank you for your business.")


def delta_layout(c, number, day, items, printed_total=None):
    """Centred title, no grid: Invoice # / Issue date dd Mon yyyy / Net amount / VAT (14%) / Amount due."""
    subtotal, tax, total = totals(items)
    c.setFillColor(AMBER)
    c.rect(0, HEIGHT - 12, WIDTH, 12, fill=1, stroke=0)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(WIDTH / 2, HEIGHT - 60, "TAX INVOICE")
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(WIDTH / 2, HEIGHT - 80, "Delta Print House")
    c.setFont("Helvetica", 9)
    c.drawCentredString(WIDTH / 2, HEIGHT - 94, "Printing and branding · 9 Galaa St., Mansoura")
    c.setFont("Helvetica", 10)
    c.drawString(50, HEIGHT - 130, f"Invoice #: {number}")
    c.drawString(50, HEIGHT - 144, f"Issue date: {day:%d %b %Y}")
    draw_customer(c, 340, HEIGHT - 130)

    top = HEIGHT - 230
    c.setFont("Helvetica-Bold", 10)
    for x, text, right in [(50, "Item", False), (360, "Quantity", True), (450, "Rate", True), (WIDTH - 50, "Line total", True)]:
        (c.drawRightString if right else c.drawString)(x, top, text)
    c.setStrokeColor(colors.black)
    c.line(50, top - 6, WIDTH - 50, top - 6)
    c.setFont("Helvetica", 10)
    for r, (d, q, p, a) in enumerate(items, start=1):
        y = top - 6 - r * 20
        c.drawString(50, y, d)
        c.drawRightString(360, y, fmt_qty(q))
        c.drawRightString(450, y, money(p))
        c.drawRightString(WIDTH - 50, y, money(a))
    y = top - 6 - (len(items) + 1) * 20 - 10
    c.line(330, y + 14, WIDTH - 50, y + 14)
    for label, value, bold in [("Net amount", subtotal, False), ("VAT (14%)", tax, False),
                               ("Amount due EGP", printed_total or total, True)]:
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 10)
        c.drawString(330, y, label)
        c.drawRightString(WIDTH - 50, y, money(value))
        y -= 16
    footer(c, "Bank transfer to Delta Print House, account 0021-44518-01.")


def canal_layout(c, number, day, items, printed_total=None):
    """Side bar, hours as quantity: Bill number / Dated Month d, yyyy / Sub-total / Tax / Grand total."""
    subtotal, tax, total = totals(items)
    c.setFillColor(NAVY)
    c.rect(0, 0, 28, HEIGHT, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(60, HEIGHT - 64, "Canal Tech Services")
    c.setFont("Helvetica", 9)
    c.setFillColor(GREY)
    c.drawString(60, HEIGHT - 80, "IT support for small businesses · Ismailia")
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10)
    c.drawString(60, HEIGHT - 120, f"Bill number {number}")
    c.drawString(60, HEIGHT - 134, f"Dated {day:%B} {day.day}, {day:%Y}")
    draw_customer(c, 340, HEIGHT - 120)

    top = HEIGHT - 220
    c.setFont("Helvetica-Bold", 10)
    for x, text, right in [(60, "Service", False), (360, "Hours", True), (450, "Rate", True), (WIDTH - 50, "Total", True)]:
        (c.drawRightString if right else c.drawString)(x, top, text)
    c.setFont("Helvetica", 10)
    for r, (d, q, p, a) in enumerate(items, start=1):
        y = top - r * 18
        c.drawString(60, y, d)
        c.drawRightString(360, y, fmt_qty(q))
        c.drawRightString(450, y, money(p))
        c.drawRightString(WIDTH - 50, y, money(a))
    y = top - (len(items) + 2) * 18
    for label, value, bold in [("Sub-total", subtotal, False), ("Tax", tax, False),
                               ("Grand total", printed_total or total, True)]:
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 10)
        c.drawString(330, y, label)
        c.drawRightString(WIDTH - 50, y, f"EGP {money(value)}")
        y -= 16
    footer(c, "Questions about this bill? support@canaltech.example")


def draw_customer(c, x, y):
    c.setFillColor(colors.black)
    for i, line in enumerate(CUSTOMER):
        c.setFont("Helvetica-Bold" if i == 0 else "Helvetica", 9)
        c.drawString(x, y - i * 12, line)


def footer(c, text):
    c.setFont("Helvetica-Oblique", 8)
    c.setFillColor(GREY)
    c.drawString(50, 40, text)


def make_scan(source_pdf, target_pdf):
    """Turn a PDF into an image-only copy, like a phone scan: same look, no text layer."""
    with pdfplumber.open(source_pdf) as pdf:
        image = pdf.pages[0].to_image(resolution=110).original.convert("L")
    c = canvas.Canvas(str(target_pdf), pagesize=A4)
    c.drawImage(ImageReader(image), 0, 0, width=WIDTH, height=HEIGHT)
    c.save()


def main(out=OUT):
    out = Path(out)
    rng = random.Random(7)
    out.mkdir(parents=True, exist_ok=True)
    layouts = {"nile": (nile_layout, "NPC-2026-{:04d}"), "delta": (delta_layout, "DPH/{:d}"),
               "canal": (canal_layout, "{:d}")}
    start = {"nile": 107, "delta": 1043, "canal": 7781}
    day = date(2026, 3, 2)
    written = []
    for i in range(12):
        supplier = ["nile", "delta", "canal"][i % 3]
        draw, pattern = layouts[supplier]
        number = pattern.format(start[supplier] + i // 3)
        items = pick_items(rng, supplier)
        printed_total = None
        if i == 7:  # a supplier typo: the printed total is 100 EGP too high
            printed_total = totals(items)[2] + 100
        day += timedelta(days=rng.randint(2, 9))
        path = out / f"{supplier}_{number.replace('/', '-')}.pdf"
        c = canvas.Canvas(str(path), pagesize=A4)
        c.setTitle(f"Invoice {number}")
        draw(c, number, day, items, printed_total)
        c.save()
        written.append(path)

    scan_source = written[0]
    make_scan(scan_source, out / "scanned_invoice.pdf")
    print(f"Wrote {len(written) + 1} invoices to {out}/ (one total that doesn't add up, one scan with no text)")


if __name__ == "__main__":
    main()
