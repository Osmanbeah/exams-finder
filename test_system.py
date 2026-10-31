"""
Automated end-to-end verification tests for Exam Finder.
"""
import os
from database import init_db, get_connection
from matcher import extract_last_6_digits, match_scan_batch
from excel_service import (
    parse_master_student_sheet,
    parse_scan_sheet,
    generate_reconciliation_excel
)

def test_extract_last_6_digits():
    assert extract_last_6_digits("202610112345") == "112345"
    assert extract_last_6_digits("112345") == "112345"
    assert extract_last_6_digits(112345) == "112345"
    assert extract_last_6_digits(112345.0) == "112345"
    assert extract_last_6_digits("ST-998877") == "998877"
    assert extract_last_6_digits("54321") == "054321" # 5 digits padded to 6
    print("[PASS] test_extract_last_6_digits passed")

def test_full_reconciliation_flow():
    # 1. Init DB
    init_db()
    
    # 2. Parse Master File
    with open("sample_master_students.xlsx", "rb") as f:
        master_bytes = f.read()
    master_students = parse_master_student_sheet(master_bytes, "sample_master_students.xlsx")
    assert len(master_students) == 13, f"Expected 13 students, got {len(master_students)}"
    
    # Save into DB
    conn = get_connection()
    c = conn.cursor()
    c.execute("DELETE FROM master_students;")
    c.executemany("""
        INSERT INTO master_students (
            student_id, student_number, last_6_digits, full_name, 
            center_name, phone, seat_number, group_name, notes
        ) VALUES (
            :student_id, :student_number, :last_6_digits, :full_name,
            :center_name, :phone, :seat_number, :group_name, :notes
        );
    """, master_students)
    conn.commit()
    
    c.execute("SELECT * FROM master_students;")
    db_students = [dict(r) for r in c.fetchall()]
    conn.close()
    
    # 3. Parse Team Member's Scan File
    with open("sample_center_scan.xlsx", "rb") as f:
        scan_bytes = f.read()
    scanned_rows = parse_scan_sheet(scan_bytes, "sample_center_scan.xlsx")
    assert len(scanned_rows) == 9, f"Expected 9 scanned rows, got {len(scanned_rows)}"
    
    # 4. Run Matching Engine for Center 'Nasr City'
    result = match_scan_batch(
        scanned_rows=scanned_rows,
        master_students=db_students,
        selected_center="Nasr City"
    )
    
    summary = result["summary"]
    print("\nMatching Summary:", summary)
    
    # Verify counts
    assert summary["total_scanned"] == 9
    assert summary["matched_count"] == 6, f"Expected 6 matched, got {summary['matched_count']}"
    assert summary["unmatched_count"] == 2, f"Expected 2 unmatched, got {summary['unmatched_count']}"
    assert summary["wrong_center_count"] == 1, f"Expected 1 wrong center, got {summary['wrong_center_count']}"
    assert summary["absent_count"] == 2, f"Expected 2 absent, got {summary['absent_count']}"
    
    # Check Unmatched (Not in DB)
    unmatched_codes = [u["extracted_6_digits"] for u in result["unmatched"]]
    assert "999888" in unmatched_codes
    assert "888777" in unmatched_codes
    print("[PASS] Unmatched students correctly identified:", unmatched_codes)
    
    # Check Wrong Center
    wrong_center_student = result["wrong_center"][0]
    assert wrong_center_student["extracted_6_digits"] == "111222"
    assert wrong_center_student["master_student"]["center_name"] == "Dokki"
    print("[PASS] Wrong center correctly identified:", wrong_center_student["master_student"]["full_name"], "registered at Dokki")
    
    # Check Absent
    absent_names = [a["full_name"] for a in result["absent"]]
    assert "Karim Sherif Nabil" in absent_names
    assert "Hoda Amr Samir" in absent_names
    print("[PASS] Absent students correctly identified:", absent_names)
    
    # 5. Generate Excel Reconciliation Report
    excel_stream = generate_reconciliation_excel(result, "Nasr_City_Batch_1")
    assert excel_stream.getbuffer().nbytes > 1000
    print(f"[PASS] Excel report generated successfully ({excel_stream.getbuffer().nbytes} bytes)")
    
    print("\nALL VERIFICATION TESTS PASSED SUCCESSFULLY! [OK]\n")

if __name__ == "__main__":
    test_extract_last_6_digits()
    test_full_reconciliation_flow()
