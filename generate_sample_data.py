"""
Generates realistic sample Excel sheets for testing:
1. sample_master_students.xlsx: Master database of students across centers
2. sample_center_scan.xlsx: Scanned batch uploaded by a team member for a specific center
"""
import openpyxl

def generate_samples():
    # 1. Master Students File
    wb_master = openpyxl.Workbook()
    ws_master = wb_master.active
    ws_master.title = "MasterStudents"
    
    headers_master = ["Student ID", "Full Name", "Center Name", "Phone", "Seat Number", "Group"]
    ws_master.append(headers_master)
    
    # 20 students in Nasr City, Dokki, Maadi
    master_records = [
        # Nasr City students
        ("202610112345", "Ahmed Mohamed Ali", "Nasr City", "01011123456", "101", "Group A"),
        ("202610223456", "Sara Mahmoud Hassan", "Nasr City", "01022234567", "102", "Group A"),
        ("202610334567", "Omar Khaled Youssef", "Nasr City", "01033345678", "103", "Group A"),
        ("202610445678", "Nour Tarek Ibrahim", "Nasr City", "01044456789", "104", "Group B"),
        ("202610556789", "Youssef Hany Mostafa", "Nasr City", "01055567890", "105", "Group B"),
        ("202610667890", "Mariam Adel Fathy", "Nasr City", "01066678901", "106", "Group B"),
        ("202610778901", "Karim Sherif Nabil", "Nasr City", "01077789012", "107", "Group C"),
        ("202610889012", "Hoda Amr Samir", "Nasr City", "01088890123", "108", "Group C"),
        
        # Dokki students
        ("202620111222", "Mona Said El-Sayed", "Dokki", "01111112223", "201", "Group D"),
        ("202620222333", "Hassan Zaki Radwan", "Dokki", "01122223334", "202", "Group D"),
        ("202620333444", "Rana Wael Gamal", "Dokki", "01133334445", "203", "Group D"),
        
        # Maadi students
        ("202630444555", "Tamer Magdy Fawzy", "Maadi", "01244445556", "301", "Group M"),
        ("202630555666", "Salma Essam Kamal", "Maadi", "01255556667", "302", "Group M"),
    ]
    
    for rec in master_records:
        ws_master.append(rec)
        
    wb_master.save("sample_master_students.xlsx")
    print("Created sample_master_students.xlsx")
    
    # 2. Scanned Exam Results File (uploaded by team member for "Nasr City")
    # Note: Only has the LAST 6 DIGITS of the student's code/number!
    wb_scan = openpyxl.Workbook()
    ws_scan = wb_scan.active
    ws_scan.title = "ExamScans"
    
    headers_scan = ["Scanned Code (Last 6 Digits)", "Score / Grade", "Scanned Name", "Paper Serial"]
    ws_scan.append(headers_scan)
    
    scan_records = [
        # Matched Nasr City students
        ("112345", 48, "Ahmed Mohamed", "P-001"),
        ("223456", 50, "Sara Mahmoud", "P-002"),
        ("334567", 42, "Omar Khaled", "P-003"),
        ("445678", 39, "Nour Tarek", "P-004"),
        ("556789", 45, "Youssef Hany", "P-005"),
        ("667890", 47, "Mariam Adel", "P-006"),
        
        # Discrepancy 1: Wrong center (belongs to Dokki in master DB, but took exam in Nasr City!)
        ("111222", 44, "Mona Said", "P-007"),
        
        # Discrepancy 2: UNREGISTERED / NOT in Master DB!
        ("999888", 35, "Unknown New Student 1", "P-008"),
        ("888777", 41, "Unknown New Student 2", "P-009"),
        
        # Note: "Karim Sherif Nabil" (778901) and "Hoda Amr Samir" (889012) were NOT scanned (Absent!)
    ]
    
    for rec in scan_records:
        ws_scan.append(rec)
        
    wb_scan.save("sample_center_scan.xlsx")
    print("Created sample_center_scan.xlsx")

if __name__ == "__main__":
    generate_samples()
