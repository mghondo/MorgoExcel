import warnings
import logging
from .OriginatingEntity import process_pdf
from .Drivers import extract_driver_names_with_camelot, extract_driver_names_with_pdfplumber, extract_driver_names_with_ocr
from .MNumber import extract_m_numbers
import os

# Suppress CropBox warnings from pdfplumber
class CropBoxWarningFilter(logging.Filter):
    def filter(self, record):
        return "CropBox missing from /Page, defaulting to MediaBox" not in record.getMessage()

logging.getLogger().addFilter(CropBoxWarningFilter())

# Suppress other warnings (e.g., CryptographyDeprecationWarning)
warnings.filterwarnings("ignore")

def extract_data(pdf_file):
    """
    Extract data from the PDF using OriginatingEntity, Drivers, and MNumber scripts.
    Args:
        pdf_file (str): Path to the PDF file.
    Returns:
        dict: Extracted data structured for BuildingScan.js.
    """
    result = {
        "company": "",
        "drivers": "",
        "items": []
    }

    # Check if the PDF exists
    if not os.path.exists(pdf_file):
        result["error"] = f"PDF file not found: {pdf_file}"
        return result

    try:
        # 1. Originating Entity
        filename, company_result = process_pdf(pdf_file)
        result["company"] = company_result if company_result != "Could not locate 'Originating Entity'" else "Not Found"

        # 2. Driver Names
        driver_names = []
        filename, camelot_driver_names = extract_driver_names_with_camelot(pdf_file)
        driver_names.extend(camelot_driver_names)
        filename, pdfplumber_driver_names = extract_driver_names_with_pdfplumber(pdf_file)
        driver_names.extend(pdfplumber_driver_names)
        for page_num in [0, 1]:
            filename, ocr_driver_names = extract_driver_names_with_ocr(pdf_file, page_number=page_num)
            driver_names.extend(ocr_driver_names)
        unique_driver_names = list(dict.fromkeys(driver_names))
        result["drivers"] = " / ".join(unique_driver_names) if unique_driver_names else "Not Found"

        # 3. M Numbers
        filename, m_numbers = extract_m_numbers(pdf_file)
        result["items"] = [
            {
                "item_number": item["M_Number"],
                "package_id": item["Package_ID"],
                "m_number": item["M_Number"],
                "name": item["Name"],
                "type": item["Category"],
                "strain": item["Strain"],
                "days": item["Days"],
                "weight": item["Weight"],
                "Item_Details": item["Item_Details"]  # Add Item_Details to the response
            }
            for item in m_numbers
        ]

    except Exception as e:
        result["error"] = f"Error processing PDF: {str(e)}"
        logging.error(f"Error in extract_data: {str(e)}")

    return result