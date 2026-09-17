import re
import os
import html
from decimal import Decimal
from PyPDF2 import PdfReader
import docx
import pyodbc

# -----------------------------
# Configuration
# -----------------------------
TABLE = "dbo.Fuel_Prices"

PRODUCT_TO_COLUMNS = {
    "Gasoline": ["Gasoline"],
    "Diesel": ["Diesel"],
    "PG-64-22": ["PG6422"],
    "PG-67-22": ["PG6722"],
    "PG-76-22": ["PG7622"],
    "PG-82-22": ["PG8222"],
    "SS-1": ["SS1"],
    "CRS-2": ["RS2C_CRS2"],
    "RS-2C": ["RS2C_CRS2"],
    "CRS-2P": ["CRS2P"],
    # ✅ full-line matching
    "EA-1, EPR-1": ["EA1_ERP1_AEP"],  # removed the end because of the & character
    "CSS-1 & 1H (Undiluted)": ["CSS1_Undiluted"],
    "CSS-1 & 1H": ["CSS1_1H"],
}

SYNONYMS = {
    # "EA-1, EPR-1": ["CRS 2", "CRS‑2", "RS-2C/CRS-2"],
    # "RS-2C": ["RS 2C", "RS‑2C"],
    # "EA-1": ["EA 1"],
    # "ERP-1": ["ERP 1"],
}

# -----------------------------
# Helpers
# -----------------------------
def load_pdf_text(path: str) -> str:
    reader = PdfReader(path)
    return "\n".join(
        html.unescape(page.extract_text() or "")
        for page in reader.pages
    )

PRICE_RE = re.compile(r"\$?\s*([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2,3})")

def try_parse_price(line: str):
    m = PRICE_RE.search(line)
    if not m:
        return None
    return Decimal(m.group(1).replace(",", ""))

MONTH_YEAR_RE = re.compile(
    r"(January|February|March|April|May|June|July|August|September|October|November|December)"
    r"\s+\d{4}",
    re.IGNORECASE,
)

def extract_month_year(text: str):
    m = MONTH_YEAR_RE.search(text)
    if not m:
        raise ValueError("Month/Year not found in PDF")
    parts = m.group(0).split()
    return parts[0].title(), int(parts[1])

def normalize(s: str):
    return re.sub(r"[\s\-]", "", s).lower()

def candidate_labels_for(key):
    return [key] + SYNONYMS.get(key, [])

# -----------------------------
# MAIN ENTRY POINT (CALLED BY UI)
# -----------------------------
def process_fuel_prices(pdf_path: str):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(pdf_path)

    text = load_pdf_text(pdf_path)
    month, year = extract_month_year(text)

    column_values = {
        "Gasoline": None,
        "Diesel": None,
        "PG6422": None,
        "PG6722": None,
        "PG7622": None,
        "PG8222": None,
        "SS1": None,
        "RS2C_CRS2": None,
        "CRS2P": None,
        "EA1_ERP1_AEP": None,
        "CSS1_Undiluted": None,
        "CSS1_1H": None,
    }

    lines = [(ln, normalize(ln)) for ln in text.splitlines() if ln.strip()]

    for product, columns in PRODUCT_TO_COLUMNS.items():
        for label in candidate_labels_for(product):
            nlabel = normalize(label)
            for raw, nraw in lines:
                if nlabel in nraw:
                    # ✅ if I'm NOT searching for undiluted, don't allow an undiluted row to match
                    if "undiluted" not in nlabel and "undiluted" in nraw:
                        continue
                    price = try_parse_price(raw)
                    if price:
                        for c in columns:
                            column_values[c] = price
                        break

    return {
    "Month": month,
    "Year": year,
    "Gasoline": str(column_values["Gasoline"]) if column_values["Gasoline"] is not None else None,
    "Diesel": str(column_values["Diesel"]) if column_values["Diesel"] is not None else None,
    "PG6422": str(column_values["PG6422"]) if column_values["PG6422"] is not None else None,
    "PG6722": str(column_values["PG6722"]) if column_values["PG6722"] is not None else None,
    "PG7622": str(column_values["PG7622"]) if column_values["PG7622"] is not None else None,
    "PG8222": str(column_values["PG8222"]) if column_values["PG8222"] is not None else None,
    "SS1": str(column_values["SS1"]) if column_values["SS1"] is not None else None,
    "RS2C_CRS2": str(column_values["RS2C_CRS2"]) if column_values["RS2C_CRS2"] is not None else None,
    "CRS2P": str(column_values["CRS2P"]) if column_values["CRS2P"] is not None else None,
    "EA1_ERP1_AEP": str(column_values["EA1_ERP1_AEP"]) if column_values["EA1_ERP1_AEP"] is not None else None,
    "CSS1_Undiluted": str(column_values["CSS1_Undiluted"]) if column_values["CSS1_Undiluted"] is not None else None,
    "CSS1_1H": str(column_values["CSS1_1H"]) if column_values["CSS1_1H"] is not None else None
}