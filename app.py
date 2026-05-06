import streamlit as st
import pdfplumber
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_bytes

# -------- PAGE CONFIG --------
st.set_page_config(
    page_title="PDF Invoice Extractor",
    layout="wide"
)

# -------- PAGE TITLE --------
st.title("📄 PDF Invoice Extractor For Swiss Military")

# -------- PDF UPLOAD --------
uploaded_file = st.file_uploader(
    "Upload Invoice PDF",
    type=["pdf"]
)

# -------- PROCESS PDF --------
if uploaded_file:

    text = ""

    # -------- READ PDF BYTES --------
    pdf_bytes = uploaded_file.getvalue()

    # -------- NORMAL PDF READ --------
    try:

        with pdfplumber.open(uploaded_file) as pdf:

            for page in pdf.pages:

                page_text = page.extract_text()

                if page_text:
                    text += page_text + "\n"

    except Exception as e:
        st.error(e)

    # -------- OCR FOR SCANNED PDF --------
    if len(text.strip()) == 0:

        st.warning("⚠️ OCR Running...")

        images = convert_from_bytes(pdf_bytes)

        for img in images:
            text += pytesseract.image_to_string(img)

    # -------- CLEAN TEXT --------
    text = text.replace("|", " ")
    text = text.replace("\n\n", "\n")

    final_data = []

    # -------- INVOICE NUMBER --------
    invoice_match = re.search(
        r'(SMCGG|SMCGM|SMLGN|SMG|SM\/HR|SM\/MH)\/\d{2}-\d{2}\/\d+',
        text
    )

    invoice_no = invoice_match.group(0) if invoice_match else ""

    # -------- E-WAY BILL --------
    eway_match = re.search(r'\b3\d{11}\b', text)
    eway_no = eway_match.group(0) if eway_match else ""

    # -------- DATE --------
    date_match = re.search(
        r'\d{1,2}-[A-Za-z]{3}-\d{2}',
        text
    )

    date = date_match.group(0) if date_match else ""

    # -------- CONSIGNEE --------
    consignee = ""

    cons_match = re.search(
        r'Consignee\s+(.*?)\s+Buyer',
        text,
        re.S
    )

    if cons_match:

        consignee_text = cons_match.group(1)

        consignee = consignee_text.split("\n")[0].strip()

        consignee = re.split(
            r'Despatch|Dated|GSTIN|State Name',
            consignee
        )[0].strip()

    # -------- BUYER --------
    buyer = ""

    buyer_match = re.search(
        r'Buyer\s+(.*?)\s+GSTIN',
        text,
        re.S
    )

    if buyer_match:

        buyer_text = buyer_match.group(1)

        buyer = buyer_text.split("\n")[0].strip()

        buyer = re.split(
            r'Despatch|Dated|GSTIN|State Name',
            buyer
        )[0].strip()

    # -------- IGST --------
    igst_match = re.search(
        r'IGST Output\s*([\d,]+\.\d{2})',
        text
    )

    igst = igst_match.group(1) if igst_match else ""

    # -------- TOTAL --------
    total_match = re.search(
        r'Total\s+\d+\s+(SET|PCS)\s+[\₹\&]?\s*([\d,]+\.\d{2})',
        text
    )

    total = total_match.group(2) if total_match else ""

    # -------- PRODUCT EXTRACTION --------
    lines = text.split("\n")

    for line in lines:

        if re.search(r'\d{8}', line) and ("SET" in line or "PCS" in line):

            try:

                clean_line = line.replace("|", " ")

                # -------- PRODUCT REGEX --------
                product_match = re.search(
                    r'([A-Z0-9_,\-/]+)\s+(\d{8})\s+(\d+)\s+(SET|PCS)',
                    clean_line
                )

                # -------- DECIMAL NUMBERS --------
                numbers = re.findall(
                    r'[\d,]+\.\d{2}',
                    clean_line
                )

                if product_match:

                    product_name = product_match.group(1)
                    hsn = product_match.group(2)
                    qty = product_match.group(3)
                    unit = product_match.group(4)

                    # -------- RATE --------
                    rate = numbers[0] if len(numbers) > 0 else ""

                    # -------- AMOUNT --------
                    amount = numbers[-1] if len(numbers) > 1 else ""

                    final_data.append([
                        invoice_no,
                        eway_no,
                        date,
                        consignee,
                        buyer,
                        product_name,
                        hsn,
                        qty,
                        unit,
                        rate,
                        amount,
                        igst,
                        total
                    ])

            except:
                pass

    # -------- DATAFRAME --------
    df = pd.DataFrame(final_data, columns=[
        "Invoice No",
        "e-Way Bill No",
        "Date",
        "Consignee",
        "Buyer",
        "Product",
        "HSN",
        "Qty",
        "Unit",
        "Rate",
        "Amount",
        "IGST",
        "Total"
    ])

    # -------- SHOW TABLE --------
    st.success("✅ Data Extracted Successfully")

    st.dataframe(
        df,
        use_container_width=True
    )

    # -------- EXCEL DOWNLOAD --------
    excel_file = "output.xlsx"

    df.to_excel(excel_file, index=False)

    with open(excel_file, "rb") as f:

        st.download_button(
            label="📥 Download Excel",
            data=f,
            file_name="invoice_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
