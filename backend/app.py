
from flask import Flask, jsonify, request, send_file, after_this_request
from flask_cors import CORS
from openpyxl import load_workbook
import fitz
import pdfplumber
import pytesseract
from PIL import Image
from io import BytesIO
import os
import re
import tempfile


# ============================================================
# FLASK APP
# ============================================================

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

CORS(
    app,
    resources={
        r"/api/*": {
            "origins": [
                "http://localhost:5173",
                "http://localhost:5174",
                "http://127.0.0.1:5173",
                "http://127.0.0.1:5174",
                "https://pdf-excel-tool.vercel.app",
            ]
        }
    },
)


# ============================================================
# TESSERACT OCR CONFIGURATION
# ============================================================

pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)

# ============================================================
# FILE SETTINGS
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")

NUMBER_PATTERN = r"\d+(?:,\d{3})*(?:\.\d+)?"


# ============================================================
# CLEAN NUMBER
# ============================================================

def clean_number(value):
    """Convert formatted numeric text into a float."""

    if value is None:
        return 0.0

    value = str(value).strip()
    value = value.replace(",", "")
    value = value.replace("₹", "")
    value = value.replace("Rs.", "")
    value = value.replace("Rs", "")
    value = value.replace("%", "")
    value = value.replace(" ", "")

    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


# ============================================================
# EXTRACT PDF TEXT
# Supports selectable text and scanned pages using OCR.
# ============================================================

def extract_pdf_text(pdf_path):
    """Extract text from a PDF, using OCR when needed."""

    extracted_pages = []

    try:
        with fitz.open(pdf_path) as pdf:
            if len(pdf) == 0:
                raise ValueError("The uploaded PDF has no pages.")

            for page_number, page in enumerate(pdf, start=1):
                text = page.get_text("text") or ""

                if len(text.strip()) < 30:
                    try:
                        pix = page.get_pixmap(
                            matrix=fitz.Matrix(2, 2),
                            alpha=False,
                        )

                        image = Image.open(
                            BytesIO(pix.tobytes("png"))
                        )

                        ocr_text = pytesseract.image_to_string(
                            image,
                            config="--psm 3",
                        )

                        if len(ocr_text.strip()) > len(text.strip()):
                            text = ocr_text

                    except Exception as ocr_error:
                        print(
                            f"OCR error on page {page_number}:",
                            ocr_error,
                        )

                extracted_pages.append(text)

        result = "\n".join(extracted_pages)

        if result.strip():
            return result

    except Exception as error:
        print("PyMuPDF extraction error: - app.py:131", error)

    # Fallback to pdfplumber and OCR.
    try:
        extracted_pages = []

        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                raise ValueError("The uploaded PDF has no pages.")

            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""

                if len(text.strip()) < 30:
                    try:
                        image = page.to_image(
                            resolution=250
                        ).original

                        ocr_text = pytesseract.image_to_string(
                            image,
                            config="--psm 3",
                        )

                        if len(ocr_text.strip()) > len(text.strip()):
                            text = ocr_text

                    except Exception as ocr_error:
                        print(
                            f"Fallback OCR error on page {page_number}:",
                            ocr_error,
                        )

                extracted_pages.append(text)

        result = "\n".join(extracted_pages)

        if result.strip():
            return result

    except Exception as error:
        print("PDF fallback extraction error: - app.py:172", error)

    raise ValueError(
        "PDF se readable text nahi mila. "
        "PDF ki validity, scan quality aur OCR installation check karein."
    )


# ============================================================
# EXTRACT PRODUCTS
# Supports numbered product rows and common unit names.
# ============================================================


def extract_products(text):
    """Extract products from multiline and common invoice layouts."""

    products = []
    lines = [" ".join(line.strip().split()) for line in text.splitlines()]
    lines = [line for line in lines if line]

    excluded_words = {
        "ITEMS", "BOX", "PACKING", "TOTAL SET", "RATE",
        "TOTAL RATE", "LOADING", "TRANSPORT", "TRANSPORTATION",
        "FREIGHT", "ROUND OFF", "ROUNDOFF", "CGST", "SGST",
        "IGST", "GRAND TOTAL", "SUB TOTAL", "SUBTOTAL",
        "TAXABLE VALUE", "TOTAL TAX", "DISCOUNT TOTAL",
        "INVOICE TOTAL", "NET AMOUNT", "NET PAYABLE",
        "AMOUNT PAYABLE", "LESS 3 %", "LESS 3%",
    }

    def number_from_line(line):
        """Return a number only when the whole line is numeric."""
        value = line.strip()
        value = re.sub(r"^(?:RS\.?|₹)\s*", "", value, flags=re.I)
        value = value.replace(",", "")
        value = value.replace("₹", "").strip()
        value = re.sub(r"\s*%$", "", value)

        if not re.fullmatch(r"-?\d+(?:\.\d+)?", value):
            return None

        try:
            return float(value)
        except ValueError:
            return None

    def is_footer_or_heading(line):
        upper = line.upper().strip()

        if upper in excluded_words:
            return True

        if re.fullmatch(r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", line):
            return True

        if upper.startswith((
            "TOTAL ", "GRAND TOTAL", "LESS ",
            "RS GIFT", "GST ", "CGST", "SGST", "IGST"
        )):
            return True

        return False

    # --------------------------------------------------------
    # FORMAT 1: MULTILINE INVOICE
    # Product name followed by:
    # Box, Packing, Total Set, Rate, Total Rate
    # --------------------------------------------------------
    header_index = -1

    for i, line in enumerate(lines):
        if line.strip().upper() == "TOTAL RATE":
            header_index = i
            break

    start_index = header_index + 1 if header_index >= 0 else 0

    for i in range(start_index, len(lines) - 5):
        name = lines[i].strip()

        # A product name must not itself be a numeric value.
        if number_from_line(name) is not None:
            continue

        if is_footer_or_heading(name):
            continue

        # The next five lines should be numeric:
        # Box, Packing, Total Set, Rate, Total Rate.
        values = [
            number_from_line(lines[i + offset])
            for offset in range(1, 6)
        ]

        if any(value is None for value in values):
            continue

        box_count, packing, total_set, rate, amount = values

        if box_count <= 0 or packing <= 0:
            continue

        if total_set <= 0 or rate <= 0 or amount <= 0:
            continue

        # Check that the extracted amount approximately matches
        # quantity multiplied by unit rate.
        expected_amount = total_set * rate
        tolerance = max(2.0, expected_amount * 0.02)

        if abs(expected_amount - amount) > tolerance:
            continue

        products.append({
            "serial_no": len(products) + 1,
            "name": name,
            "qty": total_set,
            "unit": "SET",
            "list_price": rate,
            "price": rate,
            "amount": amount,
        })

        # Skip the five numeric lines already consumed.
        # The loop advances one line at a time, so numeric lines
        # cannot become product names because they are numeric.
    
    if products:
        print("Multiline invoice products extracted:", len(products))
        return products

    # --------------------------------------------------------
    # FORMAT 2: NUMBERED PRODUCT ROWS
    # Example: 1 Product Name 12 PCS 100 80 960
    # --------------------------------------------------------
    unit_pattern = (
        r"(SETS?|PCS?|PIECES?|BOX(?:ES)?|PKTS?|PACKETS?|"
        r"DOZ(?:EN)?|KGS?|GMS?|GRAMS?|LTRS?|"
        r"LIT(?:RE|ER)S?|ML|MTRS?|MET(?:RE|ER)S?|"
        r"FT|FEET|INCH(?:ES)?|ROLLS?|BAGS?|PAIRS?|UNITS?)"
    )

    number_pattern = r"\d+(?:,\d{3})*(?:\.\d+)?"
    seen_rows = set()

    for line in lines:
        if is_footer_or_heading(line):
            continue

        serial_match = re.match(r"^(\d{1,4})[.)]?\s+(.+)$", line)

        if not serial_match:
            continue

        serial_no = int(serial_match.group(1))
        remaining = serial_match.group(2).strip()

        qty_match = re.search(
            rf"(?<![\w.])({number_pattern})\s+({unit_pattern})\b",
            remaining,
            re.IGNORECASE,
        )

        if not qty_match:
            continue

        qty = clean_number(qty_match.group(1))
        unit = qty_match.group(2).upper()

        if qty <= 0:
            continue

        name = remaining[:qty_match.start()].strip(" -:|")

        if not name:
            continue

        numeric_text = remaining[qty_match.end():]
        numbers = [
            clean_number(value)
            for value in re.findall(number_pattern, numeric_text)
        ]

        if len(numbers) < 3:
            continue

        list_price = numbers[0]
        price = numbers[-2]
        amount = numbers[-1]

        if price <= 0 or amount <= 0:
            continue

        row_key = (name.upper(), qty, price, amount)

        if row_key in seen_rows:
            continue

        seen_rows.add(row_key)

        products.append({
            "serial_no": serial_no,
            "name": name,
            "qty": qty,
            "unit": unit,
            "list_price": list_price,
            "price": price,
            "amount": amount,
        })

    print("Total products extracted:", len(products))
    return products

# ============================================================
# EXTRACT GRAND TOTAL
# ============================================================

def extract_grand_total(text):
    total_labels = [
        "GRAND TOTAL",
        "INVOICE TOTAL",
        "NET AMOUNT",
        "NET PAYABLE",
        "AMOUNT PAYABLE",
    ]

    for raw_line in text.splitlines():
        line = " ".join(raw_line.strip().split())
        upper_line = line.upper()

        if not any(label in upper_line for label in total_labels):
            continue

        values = [
            clean_number(value)
            for value in re.findall(NUMBER_PATTERN, line)
        ]

        values = [value for value in values if value > 0]

        if values:
            return values[-1]

    # Fallback: add extracted product amounts.
    products = extract_products(text)

    if products:
        return round(
            sum(product["amount"] for product in products),
            2,
        )

    return 0.0


# ============================================================
# EXTRACT TOTAL UNITS
# ============================================================

def extract_total_units(text):
    total_labels = [
        "GRAND TOTAL",
        "TOTAL QTY",
        "TOTAL QUANTITY",
        "TOTAL UNITS",
    ]

    for raw_line in text.splitlines():
        line = " ".join(raw_line.strip().split())
        upper_line = line.upper()

        if not any(label in upper_line for label in total_labels):
            continue

        values = [
            clean_number(value)
            for value in re.findall(NUMBER_PATTERN, line)
        ]

        values = [value for value in values if value > 0]

        if values:
            # This is only an estimate for different invoice layouts.
            return min(values)

    return 0.0


# ============================================================
# EXTRA CHARGE PER UNIT
# ============================================================

def calculate_extra_charge_per_unit(
    total_extra_charge,
    unit_rate,
    total_amount,
):
    if total_amount <= 0:
        return 0.0

    extra_charge_per_unit = (
        total_extra_charge * unit_rate / total_amount
    )

    return round(extra_charge_per_unit, 3)


# ============================================================
# SELLING PRICE
# ============================================================

def calculate_selling_price(
    unit_price,
    extra_charge_per_unit,
    profit_percentage,
):
    cost_price = unit_price + extra_charge_per_unit
    profit_amount = cost_price * profit_percentage / 100
    selling_price = cost_price + profit_amount

    return round(selling_price, 2)


# ============================================================
# CREATE EXCEL
# ============================================================


def create_excel(
    products,
    output_path,
    invoice_total,
    item_code,
    category,
    hsn,
    item_location,
    minimum_stock,
    transport_charge,
    gst_amount,
    any_extra_charge,
    profit_percentage,
    gst_rate,
):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(
            "template.xlsx not found at: " + TEMPLATE_PATH
        )

    workbook = load_workbook(TEMPLATE_PATH)

    try:
        if "Export Items" in workbook.sheetnames:
            worksheet = workbook["Export Items"]
        else:
            worksheet = workbook.active

        # Remove previous product rows, keep header.
        if worksheet.max_row >= 2:
            worksheet.delete_rows(2, worksheet.max_row - 1)

        total_extra_charge = (
            transport_charge
            + gst_amount
            + any_extra_charge
        )

        tax_rate_text = f"GST@{gst_rate:g}%"

        
# Read the starting item code, e.g. A1, A105, B25 or 500.
base_code = str(item_code or "1").strip()
match = re.fullmatch(r"([A-Za-z]*)(\d+)", base_code)

if not match:
    raise ValueError(
        "Item Code mein A1, A105, B25 ya 500 jaisa code enter karein."
    )

code_prefix = match.group(1)
start_number = int(match.group(2))

        for index, product in enumerate(products, start=1):
            unit_rate = product["price"]

            extra_charge_per_unit = calculate_extra_charge_per_unit(
                total_extra_charge,
                unit_rate,
                invoice_total,
            )

            selling_price = calculate_selling_price(
                unit_rate,
                extra_charge_per_unit,
                profit_percentage,
            )

            default_mrp = round(selling_price * 2, 2)

            row_number = worksheet.max_row + 1

            # Example: ITEM001, ITEM002, ITEM003...
            generated_code = f"{base_code}{index:03d}"

         
# Continue numbering from the user's starting code.
unique_code = f"{code_prefix}{start_number + index - 1}"


            values = {
                1: product["name"],       # A - Item name
                2: unique_code,            # B - Unique item code
                3: category,               # C - Category
                4: hsn,                    # D - HSN
                5: default_mrp,            # E - Default MRP
                6: selling_price,          # F - Sale price
                7: unit_rate,              # G - Online Store Price
                8: product["qty"],         # H - Opening stock quantity
                9: minimum_stock,          # I - Minimum stock
                10: item_location,         # J - Item Location
                11: tax_rate_text,         # K - Tax Rate
                12: "Inclusive",           # L - Inclusive Of Tax
                13: product["unit"],       # M - Base Unit
                14: "",                    # N - Secondary Unit
                15: "",                    # O - Conversion Rate
            }

            for column_number, value in values.items():
                worksheet.cell(
                    row=row_number,
                    column=column_number,
                    value=value,
                )

        workbook.save(output_path)

    finally:
        workbook.close()
# ============================================================
# HOME
# ============================================================

@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "success": True,
        "message": "PDF to Excel Import Tool API is running",
    })


# ============================================================
# HEALTH CHECK
# ============================================================

@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({
        "success": True,
        "status": "ok",
        "message": "PDF to Excel backend is running",
    })


# ============================================================
# GENERATE EXCEL
# ============================================================

@app.route("/api/generate-excel", methods=["POST"])
def generate_excel():
    pdf_file = request.files.get("pdf")

    if not pdf_file:
        return jsonify({
            "success": False,
            "error": "Please upload a PDF file.",
        }), 400

    if not pdf_file.filename:
        return jsonify({
            "success": False,
            "error": "Invalid PDF filename.",
        }), 400

    if not pdf_file.filename.lower().endswith(".pdf"):
        return jsonify({
            "success": False,
            "error": "Please upload a PDF file.",
        }), 400

    # --------------------------------------------------------
    # USER INPUTS
    # --------------------------------------------------------

    item_code = request.form.get("item_code", "").strip()
    category = request.form.get("category", "").strip()
    hsn = request.form.get("hsn", "").strip()
    item_location = request.form.get("item_location", "").strip()

    minimum_stock = clean_number(
        request.form.get("minimum_stock", 3)
    )

    transport_charge = clean_number(
        request.form.get("transport_charge", 0)
    )

    gst_amount = clean_number(
        request.form.get("gst_amount", 0)
    )

    any_extra_charge = clean_number(
        request.form.get("any_extra_charge", 0)
    )

    profit_percentage = clean_number(
        request.form.get("profit_percentage", 0)
    )

    gst_rate = clean_number(
        request.form.get("gst_rate", 0)
    )

    # --------------------------------------------------------
    # VALIDATION
    # --------------------------------------------------------

    if not item_code:
        return jsonify({
            "success": False,
            "error": "Item Code is required.",
        }), 400

    if not category:
        return jsonify({
            "success": False,
            "error": "Category is required.",
        }), 400

    if not hsn:
        return jsonify({
            "success": False,
            "error": "HSN is required.",
        }), 400

    numeric_inputs = [
        minimum_stock,
        transport_charge,
        gst_amount,
        any_extra_charge,
        profit_percentage,
        gst_rate,
    ]

    if any(value < 0 for value in numeric_inputs):
        return jsonify({
            "success": False,
            "error": "Numeric inputs cannot be negative.",
        }), 400

    temp_pdf = None
    output_path = None

    try:
        # ----------------------------------------------------
        # SAVE TEMPORARY PDF
        # ----------------------------------------------------

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".pdf",
        ) as temp_file:
            pdf_file.save(temp_file.name)
            temp_pdf = temp_file.name

        # ----------------------------------------------------
        # EXTRACT TEXT
        # ----------------------------------------------------

        text = extract_pdf_text(temp_pdf)

        print("\n - app.py:656" + "=" * 60)
        print("PDF TEXT EXTRACTION - app.py:657")
        print("= - app.py:658" * 60)
        print(text[:5000])
        print("= - app.py:660" * 60)

        # ----------------------------------------------------
        # EXTRACT PRODUCTS
        # ----------------------------------------------------

        products = extract_products(text)

        if not products:
            return jsonify({
                "success": False,
                "error": (
                    "No product rows could be extracted. "
                    "This PDF may use a different invoice layout. "
                    "Check the extracted text in the backend terminal."
                ),
            }), 400

        # ----------------------------------------------------
        # EXTRACT INVOICE TOTAL
        # ----------------------------------------------------

        total_amount = extract_grand_total(text)

        if total_amount <= 0:
            # Use product amounts if the invoice total is missing.
            total_amount = round(
                sum(product["amount"] for product in products),
                2,
            )

        if total_amount <= 0:
            return jsonify({
                "success": False,
                "error": (
                    "Invoice total could not be detected. "
                    "Check the extracted text and invoice amounts."
                ),
            }), 400

        displayed_total_units = extract_total_units(text)

        extracted_total_units = sum(
            product["qty"] for product in products
        )

        total_extra_charge = (
            transport_charge
            + gst_amount
            + any_extra_charge
        )

        # ----------------------------------------------------
        # PROCESSING SUMMARY
        # ----------------------------------------------------

        print("\n - app.py:716" + "=" * 60)
        print("PDF PROCESSING SUMMARY - app.py:717")
        print("= - app.py:718" * 60)
        print("Products extracted: - app.py:719", len(products))
        print("Invoice total: - app.py:720", total_amount)
        print("PDF displayed units: - app.py:721", displayed_total_units)
        print("Extracted product units: - app.py:722", extracted_total_units)
        print(
            "Unit difference:",
            displayed_total_units - extracted_total_units,
        )
        print("Transport charge: - app.py:727", transport_charge)
        print("GST amount: - app.py:728", gst_amount)
        print("Any extra charge: - app.py:729", any_extra_charge)
        print("Total extra charge: - app.py:730", total_extra_charge)
        print("Profit percentage: - app.py:731", profit_percentage)
        print("GST rate: - app.py:732", gst_rate)
        print("= - app.py:733" * 60)

        # ----------------------------------------------------
        # FIRST PRODUCT CALCULATION CHECK
        # ----------------------------------------------------

        first_product = products[0]
        first_unit_rate = first_product["price"]

        first_extra_per_unit = calculate_extra_charge_per_unit(
            total_extra_charge,
            first_unit_rate,
            total_amount,
        )

        first_selling_price = calculate_selling_price(
            first_unit_rate,
            first_extra_per_unit,
            profit_percentage,
        )

        first_default_mrp = round(first_selling_price * 2, 2)

        print("\nFIRST PRODUCT CHECK - app.py:756")
        print("Product: - app.py:757", first_product["name"])
        print("Unit price: - app.py:758", first_unit_rate)
        print("Extra charge per unit: - app.py:759", first_extra_per_unit)
        print("Selling price: - app.py:760", first_selling_price)
        print("Default MRP: - app.py:761", first_default_mrp)

        # ----------------------------------------------------
        # CREATE A UNIQUE OUTPUT FILE
        # ----------------------------------------------------

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx",
            dir=BASE_DIR,
        ) as output_file:
            output_path = output_file.name

        create_excel(
            products=products,
            output_path=output_path,
            invoice_total=total_amount,
            item_code=item_code,
            category=category,
            hsn=hsn,
            item_location=item_location,
            minimum_stock=minimum_stock,
            transport_charge=transport_charge,
            gst_amount=gst_amount,
            any_extra_charge=any_extra_charge,
            profit_percentage=profit_percentage,
            gst_rate=gst_rate,
        )

        # Remove temporary files after the response is sent.
        @after_this_request
        def cleanup_files(response):
            for path in (temp_pdf, output_path):
                if path and os.path.exists(path):
                    try:
                        os.remove(path)
                    except OSError as cleanup_error:
                        print("Temporary file cleanup error: - app.py:798", cleanup_error)

            return response

        # ----------------------------------------------------
        # SEND EXCEL FILE
        # ----------------------------------------------------

        return send_file(
            output_path,
            as_attachment=True,
            download_name="product_import.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    except Exception as exc:
        print("ERROR: - app.py:817", str(exc))

        return jsonify({
            "success": False,
            "error": str(exc),
        }), 500

    finally:
        if temp_pdf and os.path.exists(temp_pdf):
            try:
                os.remove(temp_pdf)
            except OSError:
                pass


# ============================================================
# ERROR HANDLERS
# ============================================================

@app.errorhandler(413)
def file_too_large(error):
    return jsonify({
        "success": False,
        "error": "PDF file 50 MB se chhoti honi chahiye.",
    }), 413


@app.errorhandler(500)
def internal_server_error(error):
    return jsonify({
        "success": False,
        "error": "Server error. Backend terminal mein details check karein.",
    }), 500


# ============================================================
# RUN SERVER
# ============================================================

if __name__ == "__main__":
    print("\n - app.py:857" + "=" * 60)
    print("PDF TO EXCEL BACKEND - app.py:858")
    print("= - app.py:859" * 60)
    print("Template: - app.py:860", TEMPLATE_PATH)
    print("Server: http://127.0.0.1:5000 - app.py:861")
    print("= - app.py:862" * 60)

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True,
    )