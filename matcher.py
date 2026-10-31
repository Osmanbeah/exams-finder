"""
Matching engine for Exam Finder.
Handles extraction of last 6 digits, center filtering,
and categorization (Matched, Unmatched, Wrong Center, Absent).
"""
import re
from typing import List, Dict, Any, Optional, Tuple

def extract_last_6_digits(value: Any) -> Optional[str]:
    """
    Cleans any input (float, int, string, phone number, barcode)
    and extracts the last 6 digits.
    Handles values like:
      - 10987654321 -> '765432'
      - 'ST-998877' -> '998877'
      - '012345'    -> '012345'
      - 123456.0    -> '123456'
    """
    if value is None:
        return None
        
    s = str(value).strip()
    
    # Handle float representation from Excel like "123456.0"
    if s.endswith(".0"):
        s = s[:-2]
        
    # Extract only digits
    digits = re.sub(r"\D", "", s)
    
    if len(digits) >= 6:
        return digits[-6:]
    elif len(digits) > 0:
        # If less than 6 digits, pad with leading zeros (e.g. 12345 -> 012345)
        return digits.zfill(6)
    
    return None

def normalize_text(text: Any) -> str:
    """Normalize text for consistent comparison (case, whitespace)."""
    if text is None:
        return ""
    return str(text).strip()

def match_scan_batch(
    scanned_rows: List[Dict[str, Any]], 
    master_students: List[Dict[str, Any]],
    selected_center: Optional[str] = None
) -> Dict[str, Any]:
    """
    Matches uploaded scanned records against the master students database.
    
    scanned_rows: List of dicts representing each row in the uploaded scan sheet.
                  Must have a key identified for student number/code, and optional keys for name, score, etc.
    master_students: Full list of master student records from the database.
    selected_center: The center chosen by the team member (optional or specific center).
    """
    # Build fast lookup map from master students by last_6_digits
    # A map of last_6_digits -> list of student records (to handle rare collisions)
    master_map: Dict[str, List[Dict[str, Any]]] = {}
    master_by_id: Dict[int, Dict[str, Any]] = {}
    
    for student in master_students:
        s_id = student.get("id")
        if s_id is not None:
            master_by_id[s_id] = student
            
        last6 = student.get("last_6_digits")
        if last6:
            if last6 not in master_map:
                master_map[last6] = []
            master_map[last6].append(student)
            
    matched_list: List[Dict[str, Any]] = []
    unmatched_list: List[Dict[str, Any]] = []
    wrong_center_list: List[Dict[str, Any]] = []
    scanned_master_ids = set()
    
    clean_selected_center = normalize_text(selected_center) if selected_center and selected_center.lower() != "all" else None

    for idx, row in enumerate(scanned_rows, start=1):
        raw_code = row.get("raw_code") or row.get("student_code") or row.get("number") or ""
        extracted_6 = extract_last_6_digits(raw_code)
        
        scanned_record = {
            "row_index": idx,
            "raw_code": str(raw_code),
            "roll_no": str(raw_code),
            "extracted_6_digits": extracted_6,
            "exam_model": row.get("exam_model") or "Exam",
            "score": str(row.get("score") or row.get("mark") or ""),
            "scanned_name": row.get("name") or row.get("student_name") or "",
            "extra_data": {k: v for k, v in row.items() if k not in ["raw_code", "student_code", "number", "name", "student_name", "score", "mark", "exam_model", "roll_no"]}
        }
        
        if not extracted_6 or extracted_6 not in master_map:
            # Not found in the master database at all!
            scanned_record["status"] = "unmatched"
            scanned_record["reason"] = f"Code/Number '{raw_code}' (last 6: '{extracted_6 or 'N/A'}') not found in database."
            unmatched_list.append(scanned_record)
            continue
            
        candidates = master_map[extracted_6]
        
        # If there are candidates, find the best match based on center if applicable
        chosen_student = None
        if clean_selected_center:
            for cand in candidates:
                if normalize_text(cand.get("center_name", "")).lower() == clean_selected_center.lower():
                    chosen_student = cand
                    break
        
        if not chosen_student:
            # Pick first candidate
            chosen_student = candidates[0]
            
        scanned_master_ids.add(chosen_student["id"])
        
        record_with_master = {
            **scanned_record,
            "matched_student_id": chosen_student["id"],
            "master_student": chosen_student
        }
        
        # Check center consistency
        cand_center = normalize_text(chosen_student.get("center_name", ""))
        if clean_selected_center and cand_center.lower() != clean_selected_center.lower():
            record_with_master["status"] = "wrong_center"
            record_with_master["warning"] = f"Student is registered at '{cand_center}', but uploaded under '{selected_center}'."
            wrong_center_list.append(record_with_master)
        else:
            record_with_master["status"] = "matched"
            matched_list.append(record_with_master)

    # Calculate absent students (Registered in this center, but not found in the scan)
    absent_list: List[Dict[str, Any]] = []
    if clean_selected_center:
        for student in master_students:
            if normalize_text(student.get("center_name", "")).lower() == clean_selected_center.lower():
                if student["id"] not in scanned_master_ids:
                    absent_list.append(student)

    return {
        "summary": {
            "total_scanned": len(scanned_rows),
            "matched_count": len(matched_list),
            "unmatched_count": len(unmatched_list),
            "wrong_center_count": len(wrong_center_list),
            "absent_count": len(absent_list),
            "selected_center": selected_center or "All Centers"
        },
        "matched": matched_list,
        "unmatched": unmatched_list,
        "wrong_center": wrong_center_list,
        "absent": absent_list
    }
