import fitz
import re
from openpyxl import Workbook

pdf_path = r"E:\01.NDRS_WORKON\OpenWork\PEA_INVERTER_LIST\รายชื่อทะเบียนผลิตภัณฑ์-อินเวอร์เตอร์-(Inverter)-24-มิถุนายน-2569-0917.pdf"
excel_path = r"E:\01.NDRS_WORKON\OpenWork\PEA_INVERTER_LIST\inverter_list.xlsx"

doc = fitz.open(pdf_path)

all_rows = []

for page_num in range(len(doc)):
    page = doc[page_num]
    tables = page.find_tables()
    for table in tables.tables:
        data = table.extract()
        # data[0] is header row
        if not data:
            continue
        
        # Check if this table has our expected columns
        header_row = data[0]
        header_text = " ".join(str(h) for h in header_row)
        if "Brand" in header_text and "Model" in header_text and "Rated" in header_text and "Power" in header_text:
            # Find the column indices
            col_names = [str(h).strip().replace("\n", " ") for h in header_row]
            
            # Expected columns: No, Brand, Model, Rated Power, Phase, AC Rated Voltage (v), Issue Date, Expiry Date
            # Map them
            col_map = {}
            for i, name in enumerate(col_names):
                name_clean = name.lower().strip()
                if name_clean == "no":
                    col_map["No"] = i
                elif name_clean == "brand":
                    col_map["Brand"] = i
                elif name_clean == "model":
                    col_map["Model"] = i
                elif "rated" in name_clean and "power" in name_clean:
                    col_map["Rated Power"] = i
                elif name_clean == "phase":
                    col_map["Phase"] = i
                elif "rated" in name_clean or "voltage" in name_clean or "ac" in name_clean:
                    col_map["AC Rated Voltage (v)"] = i
                elif "issue" in name_clean or "date" in name_clean:
                    if "Issue Date" not in col_map:
                        col_map["Issue Date"] = i
                elif "expiry" in name_clean or ("date" in name_clean and "issue" not in name_clean):
                    col_map["Expiry Date"] = i
            
            print(f"Page {page_num+1}: columns = {col_names}, map = {col_map}")
            
            for row in data[1:]:  # Skip header
                # Skip empty rows
                if not any(str(cell).strip() for cell in row):
                    continue
                
                # Extract values
                no_val = str(row[col_map["No"]]).strip() if "No" in col_map and col_map["No"] < len(row) else ""
                brand = str(row[col_map["Brand"]]).strip() if "Brand" in col_map and col_map["Brand"] < len(row) else ""
                model = str(row[col_map["Model"]]).strip() if "Model" in col_map and col_map["Model"] < len(row) else ""
                rated_power = str(row[col_map["Rated Power"]]).strip() if "Rated Power" in col_map and col_map["Rated Power"] < len(row) else ""
                phase = str(row[col_map["Phase"]]).strip() if "Phase" in col_map and col_map["Phase"] < len(row) else ""
                voltage = str(row[col_map["AC Rated Voltage (v)"]]).strip() if "AC Rated Voltage (v)" in col_map and col_map["AC Rated Voltage (v)"] < len(row) else ""
                issue_date = str(row[col_map["Issue Date"]]).strip() if "Issue Date" in col_map and col_map["Issue Date"] < len(row) else ""
                expiry_date = str(row[col_map["Expiry Date"]]).strip() if "Expiry Date" in col_map and col_map["Expiry Date"] < len(row) else ""
                
                # Clean Rated Power: extract just the number
                power_match = re.search(r'(\d+\.?\d*)', rated_power)
                if power_match:
                    rated_power = power_match.group(1)
                
                all_rows.append([no_val, brand, model, rated_power, phase, voltage, issue_date, expiry_date])
                print(f"  Row: {no_val}, {brand}, {model}, {rated_power}")

print(f"\nTotal rows extracted: {len(all_rows)}")

# Write to Excel
wb = Workbook()
ws = wb.active
ws.title = "Inverter List"

headers = ["No", "Brand", "Model", "Rated Power", "Phase", "AC Rated Voltage (v)", "Issue Date", "Expiry Date"]
ws.append(headers)

# Remove duplicates based on "No" column
seen_nos = set()
unique_rows = []
for row in all_rows:
    if row[0] not in seen_nos:
        seen_nos.add(row[0])
        unique_rows.append(row)

print(f"Unique rows: {len(unique_rows)}")

for row in unique_rows:
    ws.append(row)

# Auto-fit column widths
for col in ws.columns:
    max_len = 0
    col_letter = col[0].column_letter
    for cell in col:
        if cell.value:
            max_len = max(max_len, len(str(cell.value)))
    ws.column_dimensions[col_letter].width = min(max_len + 3, 50)

wb.save(excel_path)
print(f"Saved to {excel_path}")
