import re
import os
import html
from decimal import Decimal

import docx.styles.style
import docx.text.paragraph
from PyPDF2 import PdfReader
from docx import Document
import pyodbc

# -----------------------------
# Configuration
# -----------------------------
TABLE = "dbo.Fuel_Prices"


# -----------------------------
# Helpers
# -----------------------------
def extract_word_text(doc_path:str):
    word_doc = Document(doc_path)
    extracted_text = []
    # Extract text from paragraphs
    for paragraphs in word_doc.paragraphs:
        if paragraphs.text.strip():
            extracted_text.append(paragraphs.text)

    extracted_row_text = []
    for table in word_doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            extracted_row_text.append(cells)

    return extracted_text,extracted_row_text


PRICE_RE = re.compile(r"\$?\s*([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2,3})")

def try_parse_price(line: str):
    m = PRICE_RE.search(line)
    if not m:
        return None
    return Decimal(m.group(1).replace(",", ""))

def extract_month_year(text: str):
    # This function intends to find text matching only the Month Year together in the text.  Example: September 2026
    # \s +      Requires the string requires one or more whitespace characters. See below examples.
    # \d{4}     Require exactly four digits
    # re.IGNORECASE     Ignores capitalization
    #           Examples of Non-Matching: September 01, 2026, Sep 2026, 09/2026, September, 2026,Septembe r 2026
    #           Examples of Matches: September 2026, September  2026, SEPTEMBER 2026, september  2026
    MONTH_YEAR_RE = re.compile(
        r"(January|February|March|April|May|June|July|August|September|October|November|December)"
        r"\s+\d{4}",
        re.IGNORECASE,
    )
    m = MONTH_YEAR_RE.search(text)
    if not m:
        raise ValueError("Month/Year not found in PDF")
    parts = m.group(0).split()
    return parts[0].title(), int(parts[1])


def extract_gasoline_diesel(text:str):
    # This function intends to find text matching only the Gas or Diesel prices together in the text.  Example: Gasoline $3.000
    # \s +      Requires the string requires one or more whitespace characters. See below examples.
    # \t        Require string contains a tab
    # $([0-9]+\.[0-9]{2,4})     String expects a number 0-9, decimal(.),
    GAS_DIESEL_RE = re.compile(
        r"(Gasoline|Diesel)\s+\$([0-9]+\.[0-9]{2,4})",
        re.IGNORECASE,
    )

    matches = GAS_DIESEL_RE.findall(text)
    gasoline_price = matches[0][1]
    diesel_price = matches[1][1]

    if "Gasoline" not in text:
        raise ValueError(
            "Gasoline price not found in Word document"
        )

    if "Diesel" not in text:
        raise ValueError(
            "Diesel price not found in Word document"
        )
    return gasoline_price,diesel_price

def normalize(s: str):
    return re.sub(r"[\s\-]", "", s).lower()

# -----------------------------
# MAIN ENTRY POINT (CALLED BY UI)
# -----------------------------

def process_fuel_prices(file_path: str):
    if not os.path.exists(file_path):
        raise FileNotFoundError(file_path)
# Obtaining values from paragraphs and table values in the Word document
# Much of the document is a paragraph but the fuel tables are actually in a table witin the document.
    paragraphs, table_rows = extract_word_text(file_path)
    month, year = extract_month_year("\n".join(paragraphs))
    gasoline, diesel = extract_gasoline_diesel("\n".join(paragraphs))

    fuel_dic = {}
    for each_row in table_rows:
        row_label = each_row[0]
        nrow_label = normalize(row_label)
        fuel_value = each_row[1]
        nfuel_value = try_parse_price(fuel_value)
        fuel_dic.update({nrow_label:nfuel_value})

    # Clean up the Fuel dictionary "fuel_dic" and match the SQL table
    sql_table_for_import = {
        "Month": month,
        "Year": year,
        "Gasoline": gasoline,
        "Diesel": diesel,
        "PG6422": str(fuel_dic.get("gradepg6422")) if fuel_dic.get("gradepg6422") is not None else None,
        "PG6722": str(fuel_dic.get("gradepg6722")) if fuel_dic.get("gradepg6722") is not None else None,
        "PG7622": str(fuel_dic.get("gradepg7622")) if fuel_dic.get("gradepg7622") is not None else None,
        "PG8222": str(fuel_dic.get("gradepg8222")) if fuel_dic.get("gradepg8222") is not None else None,
        "SS1": str(fuel_dic.get("gradess1&ea4")) if fuel_dic.get("gradess1&ea4") is not None else None,
        "RS2C_CRS2": str(fuel_dic.get("graders2c(crs2)")) if fuel_dic.get("graders2c(crs2)") is not None else None,
        "CRS2P": str(fuel_dic.get("gradecrs2p")) if fuel_dic.get("gradecrs2p") is not None else None,
        "EA1_ERP1_AEP": str(fuel_dic.get("gradeea1,epr1,&aep")) if fuel_dic.get("gradeea1,epr1,&aep") is not None else None,
        "CSS1_Undiluted": str(fuel_dic.get("gradecss1&lh(undiluted)")) if fuel_dic.get("gradecss1&lh(undiluted)") is not None else None,
        "CSS1_1H":str(fuel_dic.get("gradecss1&1h")) if fuel_dic.get("gradecss1&1h") is not None else None
    }

    return sql_table_for_import
