# Exam Student Matcher & Discrepancy Finder 🎯

A high-performance web application designed for exam marking teams to reconcile student exam scans against master student records using **6-digit student code matching**, center verification, and automated discrepancy detection.

---

## 🌟 Key Features

1. **Master Database Management (Admin)**:
   - Upload master student lists in Excel (`.xlsx`, `.xls`) or CSV.
   - Automatic column detection for both **English** and **Arabic** headers (Student ID, Name, Center, Phone, Seat Number, Group).
   - Extracts and indexes the **last 6 digits** of each student number for instant $O(1)$ lookups.
   - Persistent SQLite database storage.

2. **Team Member Scan Upload & Center Selection**:
   - Team members pick the **Exam Center** from a dropdown.
   - Upload scanned exam marks or attendance sheets (`.xlsx`, `.xls`, `.csv`).
   - The system matches records using the **last 6 digits** of the student code.

3. **Intelligent Discrepancy Detection**:
   - **Matched Students (Green)**: Full master record displayed (Full Name, Full ID, Registered Center, Seat Number, Phone) alongside scanned score.
   - **NOT in Database (Red Alert 🚨)**: Identifies students who attended the exam / were scanned but **do not exist** in the master database. Highlights their row number, raw scanned code, and 6 digits so supervisors can immediately verify physical papers.
   - **Wrong Center (Orange 🔄)**: Flags students who attended this exam center but are officially registered at a different center.
   - **Absent Students (Slate 💤)**: Identifies registered students who did not have a scan for this exam session.

4. **Multi-User Collaboration over Local WiFi / LAN**:
   - The server can be hosted on one laptop or PC, and any team member on the same WiFi can open `http://<your-ip>:8000` to upload and match their center sheets simultaneously.

5. **One-Click Multi-Sheet Excel Export**:
   - Downloads a complete styled `.xlsx` workbook with separate sheets for:
     - Summary Dashboard
     - Matched Students
     - Discrepancies (Not in Database)
     - Wrong Center Attendees
     - Absent Students

6. **Bilingual UI**:
   - Built-in toggle between **English** and **Arabic (عربي)** with full RTL layout support.

---

## 🚀 How to Run the Application

### Option 1: Double-click the Batch file
Simply double-click:
```
start_system.bat
```
This starts the local web server and automatically opens your web browser.

### Option 2: Run via Command Line
```powershell
python run_server.py
```

---

## 🧪 Testing with Sample Data

The project includes two pre-built sample files:
- `sample_master_students.xlsx`: Sample master database with 13 students across 3 centers (Nasr City, Dokki, Maadi).
- `sample_center_scan.xlsx`: Sample scan sheet for "Nasr City" containing 9 records (6 matches, 2 unregistered students, and 1 student registered at Dokki).

You can run automated verification anytime:
```powershell
python test_system.py
```

---

## 📁 Project Structure

```
exams finder/
├── database.py              # SQLite storage for master students and scan batches
├── matcher.py               # 6-digit matching engine and discrepancy logic
├── excel_service.py         # Openpyxl parser, Arabic/English alias detector, and exporter
├── main.py                  # FastAPI REST API & static server
├── run_server.py            # Launcher script with auto-browser opening and LAN IP detection
├── start_system.bat         # 1-click Windows launcher
├── test_system.py           # End-to-end automated tests
├── generate_sample_data.py  # Sample generator for quick demonstrations
├── sample_master_students.xlsx
├── sample_center_scan.xlsx
└── static/
    └── index.html           # Modern responsive Single Page App (Vue 3 + Tailwind CSS)
```
