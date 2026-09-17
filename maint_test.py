#
#
## This module is used to test / troubleshoot the code when problems occur
### Problem 9.16.26 - The Power Automate is not completing the HTTP request.  The error indicates that the month and
###                    the year are not being read on the PDF document "09 Fuel Letter Sep 2026.pdf"
##
#

from main import process_fuel_prices, extract_word_text
import main

pdf_file = r"C:\GitProjects\Fuel-Prices-Extracted\Test Letters\09 Fuel Letter Sep 2026.pdf"
word_file = r"C:\GitProjects\Fuel-Prices-Extracted\Test Letters\09 Fuel Letter Sep 2026.docx"

process_fuel_prices(word_file)
