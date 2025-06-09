# from doctr.io import DocumentFile
# from doctr.models import ocr_predictor
# import re
# import pandas as pd
# from openpyxl import Workbook, load_workbook
# from openpyxl.styles import PatternFill
# import os

# # ========== Step 1: Load OCR Model ==========
# model = ocr_predictor(pretrained=True)

# # ========== Step 2: Setup Excel Workbook ==========
# excel_filename = "vtu_structured_results.xlsx"
# if os.path.exists(excel_filename):
#     wb = load_workbook(excel_filename)
#     ws = wb.active
# else:
#     wb = Workbook()
#     ws = wb.active
#     ws.title = "Results"
#     ws.append(["University Seat Number"])  # Initial header

# # ========== Step 3: Build Header Index ==========
# header_cells = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
# existing_usns = {ws.cell(row=i, column=header_cells["University Seat Number"]).value for i in range(2, ws.max_row + 1)}

# # ========== Step 4: Define Folder Path ==========
# image_folder = "screenshots"  # change this to your folder name
# image_extensions = (".png", ".jpg", ".jpeg")

# for filename in os.listdir(image_folder):
#     if not filename.lower().endswith(image_extensions):
#         continue

#     image_path = os.path.join(image_folder, filename)
#     doc = DocumentFile.from_images(image_path)
#     result = model(doc)
#     json_output = result.export()

#     # Extract all text lines
#     all_lines = []
#     for page in json_output['pages']:
#         for block in page['blocks']:
#             for line in block['lines']:
#                 line_text = ' '.join(word['value'] for word in line['words'])
#                 all_lines.append(line_text.strip())

#     full_text = '\n'.join(all_lines)

#     # Extract USN
#     usn_match = re.search(r'University\s*Seat\s*Number\s*[:\-]?\s*([0-9A-Z]{8,12})', full_text, re.IGNORECASE)
#     usn = usn_match.group(1).strip() if usn_match else "Not Found"

#     if usn in existing_usns or usn == "Not Found":
#         print(f"⏭️ Skipping {filename} (Already processed or USN not found)")
#         continue

#     # Extract subject marks
#     subjects = []
#     i = 0
#     while i < len(all_lines):
#         line = all_lines[i].strip()
#         code_match = re.match(r'^([A-Z]{2,}\d{2,}[A-Z]?\d*|[0-9]{2}[A-Z]{2,}\d{1,})', line)

#         if code_match:
#             code = code_match.group(1)
#             internal = external = total = None
#             result_status = ""

#             int_line = all_lines[i + 1] if i + 1 < len(all_lines) else ''
#             ext_line = all_lines[i + 2] if i + 2 < len(all_lines) else ''
#             next_lines = all_lines[i:i + 5]

#             int_vals = re.findall(r'\d+', int_line)
#             if int_vals:
#                 internal = int(int_vals[0])

#             ext_vals = re.findall(r'\d+', ext_line)
#             if len(ext_vals) >= 2:
#                 external = int(ext_vals[0])
#                 total = int(ext_vals[1])
#             elif len(ext_vals) == 1:
#                 external = int(ext_vals[0])
#                 total = internal + external if internal is not None else external
#             else:
#                 external = 0
#                 total = internal if internal is not None else 0

#             for l in next_lines:
#                 if 'F' in l and len(l.strip()) <= 3:
#                     result_status = 'F'
#                     break
#                 elif 'P' in l and len(l.strip()) <= 3:
#                     result_status = 'P'

#             if total > 200:
#                 if internal is not None and external is not None:
#                     total = internal + external

#             if internal is not None:
#                 subjects.append({
#                     "Subject Code": code,
#                     "Total": total,
#                     "Result": result_status
#                 })
#                 i += 3
#                 continue
#         i += 1

#     # Deduplicate
#     subjects_cleaned = {}
#     for sub in subjects:
#         code = sub["Subject Code"]
#         if code not in subjects_cleaned or sub["Total"] > subjects_cleaned[code]["Total"]:
#             subjects_cleaned[code] = sub

#     # Add new subject columns if any
#     for code in subjects_cleaned.keys():
#         if code not in header_cells:
#             new_col = ws.max_column + 1
#             ws.cell(row=1, column=new_col).value = code
#             header_cells[code] = new_col

#     # Build row
#     row = [None] * len(header_cells)
#     row[header_cells["University Seat Number"] - 1] = usn
#     for code, sub_info in subjects_cleaned.items():
#         row[header_cells[code] - 1] = sub_info["Total"]

#     # Append row
#     ws.append(row)
#     last_row = ws.max_row

#     # Mark failed subjects
#     fill_red = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
#     for code, sub_info in subjects_cleaned.items():
#         if sub_info["Result"] == "F":
#             col_idx = header_cells[code]
#             ws.cell(row=last_row, column=col_idx).fill = fill_red

#     print(f"✅ Processed: {filename} - {usn}")

# # Save workbook
# wb.save(excel_filename)
# print("📘 Excel file saved:", excel_filename)

import sys
import os
import re
import pandas as pd
from doctr.io import DocumentFile
from doctr.models import ocr_predictor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

# ========== Step 0: Handle CLI Argument ==========
if len(sys.argv) != 2:
    print("❌ Usage: python marks.py <screenshot_path>")
    sys.exit(1)

image_path = sys.argv[1]
if not os.path.exists(image_path):
    print(f"❌ File not found: {image_path}")
    sys.exit(1)

filename = os.path.basename(image_path)

# ========== Step 1: Load OCR Model ==========
model = ocr_predictor(pretrained=True)

# ========== Step 2: Setup Excel Workbook ==========
excel_filename = "vtu_structured_results.xlsx"
if os.path.exists(excel_filename):
    wb = load_workbook(excel_filename)
    ws = wb.active
else:
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(["University Seat Number"])  # Initial header

# ========== Step 3: Build Header Index ==========
header_cells = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
existing_usns = {ws.cell(row=i, column=header_cells["University Seat Number"]).value for i in range(2, ws.max_row + 1)}

# ========== Step 4: OCR Processing ==========
doc = DocumentFile.from_images(image_path)
result = model(doc)
json_output = result.export()

# Extract all text lines
all_lines = []
for page in json_output['pages']:
    for block in page['blocks']:
        for line in block['lines']:
            line_text = ' '.join(word['value'] for word in line['words'])
            all_lines.append(line_text.strip())

full_text = '\n'.join(all_lines)

# Extract USN
usn_match = re.search(r'University\s*Seat\s*Number\s*[:\-]?\s*([0-9A-Z]{8,12})', full_text, re.IGNORECASE)
usn = usn_match.group(1).strip() if usn_match else "Not Found"

if usn in existing_usns or usn == "Not Found":
    print(f"⏭️ Skipping {filename} (Already processed or USN not found)")
    sys.exit(0)

# ========== Step 5: Extract subject marks ==========
subjects = []
i = 0
while i < len(all_lines):
    line = all_lines[i].strip()
    code_match = re.match(r'^([A-Z]{2,}\d{2,}[A-Z]?\d*|[0-9]{2}[A-Z]{2,}\d{1,})', line)

    if code_match:
        code = code_match.group(1)
        internal = external = total = None
        result_status = ""

        int_line = all_lines[i + 1] if i + 1 < len(all_lines) else ''
        ext_line = all_lines[i + 2] if i + 2 < len(all_lines) else ''
        next_lines = all_lines[i:i + 5]

        int_vals = re.findall(r'\d+', int_line)
        if int_vals:
            internal = int(int_vals[0])

        ext_vals = re.findall(r'\d+', ext_line)
        if len(ext_vals) >= 2:
            external = int(ext_vals[0])
            total = int(ext_vals[1])
        elif len(ext_vals) == 1:
            external = int(ext_vals[0])
            total = internal + external if internal is not None else external
        else:
            external = 0
            total = internal if internal is not None else 0

        for l in next_lines:
            if 'F' in l and len(l.strip()) <= 3:
                result_status = 'F'
                break
            elif 'P' in l and len(l.strip()) <= 3:
                result_status = 'P'

        if total > 200:
            if internal is not None and external is not None:
                total = internal + external

        if internal is not None:
            subjects.append({
                "Subject Code": code,
                "Total": total,
                "Result": result_status
            })
            i += 3
            continue
    i += 1

# Deduplicate
subjects_cleaned = {}
for sub in subjects:
    code = sub["Subject Code"]
    if code not in subjects_cleaned or sub["Total"] > subjects_cleaned[code]["Total"]:
        subjects_cleaned[code] = sub

# Add new subject columns if any
for code in subjects_cleaned.keys():
    if code not in header_cells:
        new_col = ws.max_column + 1
        ws.cell(row=1, column=new_col).value = code
        header_cells[code] = new_col

# Build row
row = [None] * len(header_cells)
row[header_cells["University Seat Number"] - 1] = usn
for code, sub_info in subjects_cleaned.items():
    row[header_cells[code] - 1] = sub_info["Total"]

# Append row
ws.append(row)
last_row = ws.max_row

# Mark failed subjects
fill_red = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
for code, sub_info in subjects_cleaned.items():
    if sub_info["Result"] == "F":
        col_idx = header_cells[code]
        ws.cell(row=last_row, column=col_idx).fill = fill_red

print(f"✅ Processed: {filename} - {usn}")

# Save workbook
wb.save(excel_filename)
print("📘 Excel file saved:", excel_filename)
