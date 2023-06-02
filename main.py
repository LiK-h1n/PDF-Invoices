from glob import glob
from pandas import read_excel
from fpdf import FPDF
from datetime import datetime

filepaths = glob("Invoices/*.xlsx")

for filepath in filepaths:
    filename, date = filepath[9:].split("-")
    date = datetime.strptime(date[:-5], "%Y.%m.%d").date().strftime("%B %d, %Y")

    pdf = FPDF("P", "mm", "A4")
    pdf.add_page()
    pdf.set_font("Times", "B", 16)
    pdf.cell(50, 8, f"Invoice no: {filename}", ln=1)
    pdf.cell(50, 8, f"Date: {date}", ln=1)
    pdf.cell(50, 12, ln=1)

    df = read_excel(filepath, "Sheet 1")
    headings = [heading.replace("_", " ").title().replace("Per", "per").replace("Purchased", "")
                for heading in df.columns]

    pdf.set_font("Times", "B", 12)
    for heading in headings:
        if heading == headings[1]:
            pdf.cell(70, 8, heading, border=1)
        elif heading == headings[-1]:
            pdf.cell(30, 8, heading, border=1, ln=1)
        else:
            pdf.cell(30, 8, heading, border=1)

    pdf.set_font("Times", size=12)
    for index, row in df.iterrows():
        pdf.cell(30, 8, str(row["product_id"]), border=1)
        pdf.cell(70, 8, row["product_name"], border=1)
        pdf.cell(30, 8, str(row["amount_purchased"]), border=1)
        pdf.cell(30, 8, str(row["price_per_unit"]), border=1)
        pdf.cell(30, 8, str(row["total_price"]), border=1, ln=1)

    total = df["total_price"].sum()
    for i in range(4):
        if i == 1:
            pdf.cell(70, 8, border=1)
        else:
            pdf.cell(30, 8, border=1)
    pdf.cell(30, 8, str(total), 1, 1)
    pdf.cell(50, 12, ln=1)

    pdf.set_font("Times", "B", 16)
    pdf.cell(30, 8, f"The total due amount is {total} Euros.")

    pdf.output(f"PDFs/{filename}.pdf")
