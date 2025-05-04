import pdfplumber
import re

def extract_data_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text()

    # Extract Originating Entity
    company_match = re.search(r'1\.\s*Outbound\s*Transporter\s*(.*?)\n', text, re.IGNORECASE)
    company = company_match.group(1) if company_match else "Not found"

    # Extract Drivers
    drivers = re.findall(r'Name of Person Transporting\s*(.*?)\n', text)
    drivers_str = " / ".join(drivers) if drivers else "Not found"

    # Extract Package IDs, M Numbers, Names, and Details
    items = re.findall(r'(\d+)\.\s*Package\s*\|\s*Shipped(.*?)(?=\d+\.\s*Package\s*\|\s*Shipped|\Z)', text, re.DOTALL)

    extracted_items = []
    for item_num, item_text in items:
        package_id_match = re.search(r'(\dA\d+)', item_text)
        m_number_match = re.search(r'(M\d{11}):\s*(.*?)(?:\s*\(|$)', item_text, re.DOTALL)
        details_match = re.search(r'(?:Brand|Wet|Qty).*', item_text, re.IGNORECASE)

        package_id = package_id_match.group(1) if package_id_match else "Not found"
        details = details_match.group(0) if details_match else "Not found"

        name = "Not found"
        if m_number_match:
            m_number_cell = m_number_match.group(2).strip()
            name_match = re.search(r'(?:- |-)([^-|]+)(?:\s*-|\s*\||$)', m_number_cell)
            if name_match:
                name = name_match.group(1).strip()

        strain = "Not found"
        if m_number_match:
            m_number_cell = m_number_match.group(2).strip()
            if "Indica" in m_number_cell or "Indica" in details:
                strain = "Indica"
            elif "Sativa" in m_number_cell or "Sativa" in details:
                strain = "Sativa"
            elif "Hybrid" in m_number_cell or "Hybrid" in details:
                strain = "Hybrid"

        type_info = "Undetermined"
        if m_number_match:
            combined_text = (m_number_cell + " " + details).lower()
            if any(weight in combined_text for weight in ["2.83", "5.66", "14.15", "28.3"]):
                type_info = "Flower"
            elif any(edible_type in combined_text for edible_type in ["gummies", "edibles", "drink"]):
                type_info = "Edibles"
            elif any(cartridge_type in combined_text for cartridge_type in ["distillate", "cart", "cartridge"]):
                type_info = "Cartridge"

        days_match = re.search(r'Supply:\s*(\d+)', details, re.IGNORECASE)
        days = days_match.group(1) if days_match else "Not found"

        weight = "Not found"
        if any(weight_value in combined_text for weight_value in ["2.83", "5.66", "14.15", "28.3"]):
            weight_matches = [value for value in ["2.83", "5.66", "14.15", "28.3"] if value in combined_text]
            weight = weight_matches[0] if weight_matches else weight

        extracted_items.append({
            "item_number": item_num,
            "package_id": package_id,
            "m_number": m_number_match.group(1) if m_number_match else "Not found",
            "name": name,
            "strain": strain,
            "type": type_info,
            "days": days,
            "weight": weight
        })

    return company, drivers_str, extracted_items
