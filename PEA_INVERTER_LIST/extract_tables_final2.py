import fitz
import re
import sys
from openpyxl import Workbook

# Force UTF-8 for stdout
sys.stdout.reconfigure(encoding='utf-8')

pdf_path = r"E:\01.NDRS_WORKON\OpenWork\PEA_INVERTER_LIST\รายชื่อทะเบียนผลิตภัณฑ์-อินเวอร์เตอร์-(Inverter)-24-มิถุนายน-2569-0917.pdf"
excel_path = r"E:\01.NDRS_WORKON\OpenWork\PEA_INVERTER_LIST\inverter_list.xlsx"

doc = fitz.open(pdf_path)

all_rows = []

for page_num in range(len(doc)):
    page = doc[page_num]
    tables = page.find_tables()
    for table in tables.tables:
        data = table.extract()
        if not data:
            continue
        
        header_row = data[0]
        header_text = " ".join(str(h) for h in header_row)
        
        if "Brand" in header_text and "Model" in header_text and "Rated" in header_text:
            col_names = [str(h).strip().replace("\n", " ") for h in header_row]
            print(f"Page {page_num+1}: headers = {col_names}")
            
            # Build column index map
            col_map = {}
            for i, name in enumerate(col_names):
                name_clean = name.lower().strip()
                if name_clean == "no":
                    col_map["No"] = i
                elif name_clean == "brand" or "brand" in name_clean:
                    col_map["Brand"] = i
                elif name_clean == "model" or "model" in name_clean:
                    col_map["Model"] = i
                elif "rated" in name_clean and "power" in name_clean:
                    col_map["Rated Power"] = i
                elif name_clean == "phase" or "phase" in name_clean:
                    col_map["Phase"] = i
                elif "rated" in name_clean or "voltage" in name_clean or "ac" in name_clean:
                    if "voltage" in name_clean:
                        col_map["AC Rated Voltage (v)"] = i
                elif "issue" in name_clean:
                    col_map["Issue Date"] = i
                elif "expiry" in name_clean or ("date" in name_clean and "issue" not in name_clean):
                    col_map["Expiry Date"] = i
            
            # Handle case where Expiry and Date are separate columns
            # Look for "Expiry" or "Date" in the column names
            expiry_col = None
            date_col = None
            for i, name in enumerate(col_names):
                name_clean = name.lower().strip()
                if "expiry" in name_clean:
                    expiry_col = i
                if name_clean == "date":
                    date_col = i
            
            if "Expiry Date" not in col_map and expiry_col is not None:
                if date_col is not None and date_col != expiry_col:
                    # "Expiry" and "Date" are separate columns
                    col_map["Expiry"] = expiry_col
                    col_map["Date"] = date_col
            
            print(f"  map = {col_map}")
            
            for row_idx, row in enumerate(data[1:]):
                if not any(str(cell).strip() for cell in row):
                    continue
                
                def get(col_name):
                    if col_name in col_map and col_map[col_name] < len(row):
                        return str(row[col_map[col_name]]).strip()
                    return ""
                
                no_val = get("No")
                brand = get("Brand")
                model = get("Model")
                rated_power = get("Rated Power")
                phase = get("Phase")
                voltage = get("AC Rated Voltage (v)")
                
                # Handle split Expiry/Date columns
                if "Expiry Date" in col_map:
                    expiry_date = get("Expiry Date")
                elif "Expiry" in col_map and "Date" in col_map:
                    e = get("Expiry")
                    d = get("Date")
                    expiry_date = f"{e} {d}".strip()
                else:
                    expiry_date = ""
                
                issue_date = get("Issue Date")
                
                # Clean Rated Power: extract just the number
                power_match = re.search(r'(\d+\.?\d*)', rated_power)
                if power_match:
                    rated_power = power_match.group(1)
                
                # Remove Thai characters from Brand, Model, Voltage (they leaked from adjacent date cells)
                thai_re = re.compile(r'[\u0E00-\u0E7F]')
                brand = thai_re.sub('', brand).strip()
                model = thai_re.sub('', model).strip()
                voltage = thai_re.sub('', voltage).strip()
                
                # Clean Phase: extract only digit 1 or 3, remove all Thai/newlines
                phase_clean = re.sub(r'[\u0E00-\u0E7F\s]', '', phase).strip()
                phase_match = re.search(r'^(\d+)$', phase_clean)
                phase = phase_match.group(1) if phase_match else phase_clean
                
                # Clean Issue Date: remove leading garbage, keep only the date part
                # Thai date format: D MMM YY (e.g. "5 ส.ค. 68")
                def clean_date(date_str):
                    if not date_str:
                        return ""
                    # Remove leading Thai chars and newlines before the date
                    d = re.sub(r'^[\s\u0E00-\u0E7F\n]+', '', date_str)
                    # Remove trailing Thai chars and newlines after the date
                    d = re.sub(r'[\s\u0E00-\u0E7F\n]+$', '', d)
                    # Remove any remaining newlines/extra spaces in the middle
                    d = ' '.join(d.split())
                    return d
                
                issue_date = clean_date(issue_date)
                expiry_date = clean_date(expiry_date)
                
                # Also clean up leftover whitespace/newlines
                brand = ' '.join(brand.split())
                model = ' '.join(model.split())
                voltage = ' '.join(voltage.split())
                
                all_rows.append([no_val, brand, model, rated_power, phase, voltage, issue_date, expiry_date])

print(f"\nTotal rows before dedup: {len(all_rows)}")

# Deduplicate by "No" column
seen_nos = set()
unique_rows = []
for row in all_rows:
    if row[0] and row[0] not in seen_nos:
        seen_nos.add(row[0])
        unique_rows.append(row)

print(f"Unique rows: {len(unique_rows)}")

# Write to Excel
wb = Workbook()
ws = wb.active
ws.title = "Inverter List"

headers = ["No", "Brand", "Model", "Rated Power", "Phase", "AC Rated Voltage (v)", "Issue Date", "Expiry Date"]
ws.append(headers)

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
print("Done!")
