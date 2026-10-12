#
#
## This module is used to test / troubleshoot the code when problems occur
### Problem 9.16.26 - The Power Automate is not completing the HTTP request.  The error indicates that the month and
###                    the year are not being read on the PDF document "09 Fuel Letter Sep 2026.pdf"
##
#

from main import process_fuel_prices, extract_word_text
import main
import docx.styles.style
import docx.text.paragraph
from PyPDF2 import PdfReader
from docx import Document
import pyodbc


pdf_file = r"C:\GitProjects\Fuel-Prices-Extracted\Test Letters\09 Fuel Letter Sep 2026.pdf"
word_file = r"C:\GitProjects\Fuel-Prices-Extracted\Test Letters\10 Fuel Letter Oct 2026.docx"
# September 2026 fuel switched code to work with a Word document instead of a PDF.  
# October 12, 2026. the automated letter upload had an error when processing.  10 Fuel Letter Oct 2026.docx used for testing. 
#       ValidationFailed. The schema validation failed. "CSS1_Undiluted": null appears to be the error in Parse JSON. 
#       Troubleshooting: "CSS1_Undiluted": str(fuel_dic.get("gradecss1&1h(undiluted)")) if fuel_dic.get("gradecss1&1h(undiluted)") is not None else None,
#       "gradecss1&lh(undiluted)")) had an l instead of a 1.  They look very similar. 


process_fuel_prices(word_file)
