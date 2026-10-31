"""
Excel and CSV import/export service for Exam Finder.
Provides intelligent column mapping for English & Arabic sheets,
robust data sanitization, and beautifully formatted reconciliation exports.
"""
import io
import re
from typing import List, Dict, Any, Tuple, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from matcher import extract_last_6_digits

# Mapping dictionary for auto-detecting common column variations
HEADER_ALIASES = {
    "student_number": [
        "student number", "student_number", "student code", "student_code", "code", "id", "student id", 
        "seat number", "seat_number", "roll no", "phone", "mobile", "رقم الطالب", "كود الطالب", 
        "كود", "رقم الجلوس", "جلوس", "الهاتف", "موبايل", "رقم الهاتف", "الرقم التعريفي", "باركود", "barcode"
    ],
    "full_name": [
        "full name", "full_name", "student name", "student_name", "name", 
        "اسم الطالب", "الاسم", "اسم الطالب بالكامل", "الاسم ثلاثي", "الاسم رباعي"
    ],
    "center_name": [
        "center", "center name", "center_name", "branch", "exam center", "location",
        "السنتر", "اسم السنتر", "الفرع", "المقر", "لجنة", "اسم اللجنة"
    ],
    "phone": [
        "phone", "mobile", "tel", "phone number", "parent phone", "contact",
        "تليفون", "هاتف", "رقم التليفون", "موبايل ولي الامر", "تليفون ولي الامر", "رقم الموبايل"
    ],
    "seat_number": [
        "seat", "seat number", "seat_no", "رقم الجلوس", "جلوس", "المقعد", "رقم المقعد"
    ],
    "group_name": [
        "group", "group name", "batch", "class", "المجموعة", "الفصل", "الجروب", "الميعاد"
    ],
    "score": [
        "score", "mark", "grade", "degree", "result", "total", "percentage",
        "الدرجة", "العلامة", "النتيجة", "المجموع", "درجة الامتحان"
    ]
}

def clean_cell(val: Any) -> str:
    """Format cell value cleanly removing excessive whitespace and trailing .0."""
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

def identify_columns(headers: List[str]) -> Dict[str, Optional[int]]:
    """
    Given a list of header strings from an Excel/CSV file,
    match them against known semantic fields.
    Returns mapping of field_name -> col_index.
    """
    mapping: Dict[str, Optional[int]] = {
        "student_number": None,
        "full_name": None,
        "center_name": None,
        "phone": None,
        "seat_number": None,
        "group_name": None,
        "score": None
    }
    
    cleaned_headers = [str(h).strip().lower() for h in headers]
    
    for field, aliases in HEADER_ALIASES.items():
        for col_idx, h in enumerate(cleaned_headers):
            if not h:
                continue
            # Exact match
            if h in aliases:
                mapping[field] = col_idx
                break
            # Partial match
            if any(alias in h for alias in aliases):
                mapping[field] = col_idx
                break
                
    return mapping

def read_excel_or_csv(file_bytes: bytes, filename: str) -> Tuple[List[str], List[List[Any]]]:
    """
    Reads an uploaded Excel (.xlsx, .xls) or CSV file and returns (headers, rows).
    """
    filename_lower = filename.lower()
    
    if filename_lower.endswith(".csv"):
        # Try UTF-8 first, then windows-1256 (Arabic) or latin1
        for encoding in ["utf-8-sig", "utf-8", "cp1256", "latin1"]:
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding=encoding, dtype=str)
                headers = [str(c) for c in df.columns]
                rows = df.fillna("").values.tolist()
                return headers, rows
            except UnicodeDecodeError:
                continue
        # Fallback
        df = pd.read_csv(io.BytesIO(file_bytes), dtype=str)
        return [str(c) for c in df.columns], df.fillna("").values.tolist()
    else:
        # Excel (.xlsx, .xls)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet = wb.active
        all_rows = list(sheet.iter_rows(values_only=True))
        if not all_rows:
            return [], []
            
        # Locate the header row (first non-empty row)
        header_row_idx = 0
        for i, row in enumerate(all_rows):
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                header_row_idx = i
                break
                
        headers = [clean_cell(c) for c in all_rows[header_row_idx]]
        data_rows = []
        for row in all_rows[header_row_idx + 1:]:
            # Ignore empty rows
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                data_rows.append([clean_cell(c) for c in row])
                
        return headers, data_rows

def parse_master_student_sheet(file_bytes: bytes, filename: str, custom_mapping: Optional[Dict[str, int]] = None) -> List[Dict[str, Any]]:
    """
    Parses the Master Student Excel sheet and returns standardized student dictionaries.
    """
    headers, rows = read_excel_or_csv(file_bytes, filename)
    col_map = custom_mapping or identify_columns(headers)
    
    # Fallback heuristics if column detection missed something
    if col_map["student_number"] is None:
        # Default to first column with numeric values or column 0
        col_map["student_number"] = 0
    if col_map["full_name"] is None and len(headers) > 1:
        col_map["full_name"] = 1
    if col_map["center_name"] is None and len(headers) > 2:
        col_map["center_name"] = 2

    students = []
    for row in rows:
        def get_val(key):
            idx = col_map.get(key)
            if idx is not None and idx < len(row):
                return clean_cell(row[idx])
            return ""

        raw_num = get_val("student_number")
        if not raw_num:
            continue
            
        last_6 = extract_last_6_digits(raw_num)
        if not last_6:
            continue
            
        full_name = get_val("full_name") or f"Student {raw_num}"
        center_name = get_val("center_name") or "Main Center"
        phone = get_val("phone")
        seat_number = get_val("seat_number")
        group_name = get_val("group_name")
        
        students.append({
            "student_id": raw_num,
            "student_number": raw_num,
            "last_6_digits": last_6,
            "full_name": full_name,
            "center_name": center_name,
            "phone": phone,
            "seat_number": seat_number,
            "group_name": group_name,
            "notes": ""
        })
        
    return students

def parse_scan_sheet(file_bytes: bytes, filename: str, custom_mapping: Optional[Dict[str, int]] = None) -> List[Dict[str, Any]]:
    """
    Parses an uploaded scan sheet from an exam center.
    Extracts the 6-digit student code / number, score, and any optional info.
    """
    headers, rows = read_excel_or_csv(file_bytes, filename)
    col_map = custom_mapping or identify_columns(headers)
    
    if col_map["student_number"] is None:
        col_map["student_number"] = 0

    scanned_rows = []
    for row in rows:
        def get_val(key):
            idx = col_map.get(key)
            if idx is not None and idx < len(row):
                return clean_cell(row[idx])
            return ""

        raw_code = get_val("student_number")
        if not raw_code:
            continue
            
        name = get_val("full_name")
        score = get_val("score")
        
        # Capture any extra columns
        extra = {}
        for idx, val in enumerate(row):
            if idx not in [col_map.get("student_number"), col_map.get("full_name"), col_map.get("score")]:
                header_title = headers[idx] if idx < len(headers) and headers[idx] else f"Col_{idx+1}"
                extra[header_title] = clean_cell(val)

        scanned_rows.append({
            "raw_code": raw_code,
            "student_code": raw_code,
            "name": name,
            "score": score,
            **extra
        })

    return scanned_rows

def generate_reconciliation_excel(match_result: Dict[str, Any], batch_name: str) -> io.BytesIO:
    """
    Creates a styled, professional multi-sheet Excel workbook containing:
    1. Summary Dashboard
    2. Matched Students (Full profiles + Scan data)
    3. Discrepancies / Unmatched (Students not found in DB)
    4. Wrong Center Discrepancies
    5. Absent Students
    """
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    # Styles
    title_font = Font(name="Calibri", size=14, bold=True, color="1F2937")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    regular_font = Font(name="Calibri", size=10)
    bold_font = Font(name="Calibri", size=10, bold=True)
    
    header_fill_blue = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
    header_fill_green = PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")
    header_fill_red = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")
    header_fill_orange = PatternFill(start_color="9A3412", end_color="9A3412", fill_type="solid")
    header_fill_slate = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    
    card_fill = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="E5E7EB"),
        right=Side(style="thin", color="E5E7EB"),
        top=Side(style="thin", color="E5E7EB"),
        bottom=Side(style="thin", color="E5E7EB")
    )
    
    # ----------------- SHEET 1: Summary Dashboard -----------------
    ws_sum = wb.create_sheet(title="Summary")
    ws_sum.views.sheetView[0].showGridLines = True
    
    ws_sum["A1"] = f"Exam Reconciliation Report - {batch_name}"
    ws_sum["A1"].font = title_font
    
    summary = match_result.get("summary", {})
    metrics = [
        ("Selected Exam Center", summary.get("selected_center", "All")),
        ("Total Scanned in Sheet", summary.get("total_scanned", 0)),
        ("Successfully Matched Students", summary.get("matched_count", 0)),
        ("DISCREPANCY: Not in Master Database", summary.get("unmatched_count", 0)),
        ("DISCREPANCY: Wrong Center Registered", summary.get("wrong_center_count", 0)),
        ("Registered Students Absent from Scan", summary.get("absent_count", 0)),
    ]
    
    for row_idx, (label, val) in enumerate(metrics, start=3):
        ws_sum.cell(row=row_idx, column=1, value=label).font = bold_font
        ws_sum.cell(row=row_idx, column=1).fill = card_fill
        ws_sum.cell(row=row_idx, column=1).border = thin_border
        
        val_cell = ws_sum.cell(row=row_idx, column=2, value=val)
        val_cell.font = bold_font
        val_cell.border = thin_border
        val_cell.alignment = Alignment(horizontal="center")
        
        if "DISCREPANCY: Not" in label and val > 0:
            val_cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            val_cell.font = Font(name="Calibri", size=10, bold=True, color="991B1B")
        elif "Matched" in label:
            val_cell.fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
            val_cell.font = Font(name="Calibri", size=10, bold=True, color="065F46")
    
    ws_sum.column_dimensions["A"].width = 38
    ws_sum.column_dimensions["B"].width = 25

    # ----------------- SHEET 2: Matched Students -----------------
    ws_matched = wb.create_sheet(title="Matched Students")
    ws_matched.views.sheetView[0].showGridLines = True
    
    matched_headers = [
        "Row", "Last 6 Digits", "Full Student Number", "Full Name", 
        "Registered Center", "Score / Mark", "Phone Number", "Seat Number", "Group"
    ]
    for col_num, h_text in enumerate(matched_headers, 1):
        cell = ws_matched.cell(row=1, column=col_num, value=h_text)
        cell.font = header_font
        cell.fill = header_fill_green
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for r_idx, item in enumerate(match_result.get("matched", []), start=2):
        m = item.get("master_student", {})
        row_vals = [
            item.get("row_index"),
            item.get("extracted_6_digits"),
            m.get("student_number"),
            m.get("full_name"),
            m.get("center_name"),
            item.get("score"),
            m.get("phone"),
            m.get("seat_number"),
            m.get("group_name")
        ]
        for c_idx, val in enumerate(row_vals, 1):
            c = ws_matched.cell(row=r_idx, column=c_idx, value=val)
            c.font = regular_font
            c.border = thin_border
            if c_idx in [1, 2, 6, 8]:
                c.alignment = Alignment(horizontal="center")

    # ----------------- SHEET 3: Discrepancies - Unmatched -----------------
    ws_unmatched = wb.create_sheet(title="NOT in Database")
    ws_unmatched.views.sheetView[0].showGridLines = True
    
    unmatched_headers = [
        "Sheet Row", "Raw Code from Scan", "Extracted 6 Digits", 
        "Scanned Name (if any)", "Scanned Score", "Issue / Warning"
    ]
    for col_num, h_text in enumerate(unmatched_headers, 1):
        cell = ws_unmatched.cell(row=1, column=col_num, value=h_text)
        cell.font = header_font
        cell.fill = header_fill_red
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for r_idx, item in enumerate(match_result.get("unmatched", []), start=2):
        row_vals = [
            item.get("row_index"),
            item.get("raw_code"),
            item.get("extracted_6_digits"),
            item.get("scanned_name"),
            item.get("score"),
            item.get("reason", "Student code does not exist in master database")
        ]
        for c_idx, val in enumerate(row_vals, 1):
            c = ws_unmatched.cell(row=r_idx, column=c_idx, value=val)
            c.font = regular_font
            c.border = thin_border
            if c_idx in [1, 2, 3, 5]:
                c.alignment = Alignment(horizontal="center")
            if c_idx == 6:
                c.font = bold_font
                c.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    # ----------------- SHEET 4: Wrong Center -----------------
    ws_wrong = wb.create_sheet(title="Wrong Center")
    ws_wrong.views.sheetView[0].showGridLines = True
    
    wrong_headers = [
        "Sheet Row", "Last 6 Digits", "Student Name", "Uploaded Under Center", 
        "Official Registered Center", "Full Student Number", "Phone"
    ]
    for col_num, h_text in enumerate(wrong_headers, 1):
        cell = ws_wrong.cell(row=1, column=col_num, value=h_text)
        cell.font = header_font
        cell.fill = header_fill_orange
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for r_idx, item in enumerate(match_result.get("wrong_center", []), start=2):
        m = item.get("master_student", {})
        row_vals = [
            item.get("row_index"),
            item.get("extracted_6_digits"),
            m.get("full_name"),
            summary.get("selected_center"),
            m.get("center_name"),
            m.get("student_number"),
            m.get("phone")
        ]
        for c_idx, val in enumerate(row_vals, 1):
            c = ws_wrong.cell(row=r_idx, column=c_idx, value=val)
            c.font = regular_font
            c.border = thin_border
            if c_idx in [1, 2]:
                c.alignment = Alignment(horizontal="center")

    # ----------------- SHEET 5: Absent Students -----------------
    ws_absent = wb.create_sheet(title="Absent Students")
    ws_absent.views.sheetView[0].showGridLines = True
    
    absent_headers = ["Student Number", "Last 6 Digits", "Full Name", "Center", "Phone", "Seat", "Group"]
    for col_num, h_text in enumerate(absent_headers, 1):
        cell = ws_absent.cell(row=1, column=col_num, value=h_text)
        cell.font = header_font
        cell.fill = header_fill_slate
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for r_idx, s in enumerate(match_result.get("absent", []), start=2):
        row_vals = [
            s.get("student_number"),
            s.get("last_6_digits"),
            s.get("full_name"),
            s.get("center_name"),
            s.get("phone"),
            s.get("seat_number"),
            s.get("group_name")
        ]
        for c_idx, val in enumerate(row_vals, 1):
            c = ws_absent.cell(row=r_idx, column=c_idx, value=val)
            c.font = regular_font
            c.border = thin_border

    # Adjust column widths for all sheets
    for ws in [ws_matched, ws_unmatched, ws_wrong, ws_absent]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
