"""
FastAPI application for Exam Finder.
Provides REST APIs for database management, center selection,
scan matching, and downloading discrepancy reports.
"""
import os
import io
import json
import socket
from typing import Optional, List, Dict, Any
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from database import init_db, get_connection
from matcher import match_scan_batch
from excel_service import (
    parse_master_student_sheet,
    parse_scan_sheet,
    generate_reconciliation_excel,
    read_excel_or_csv,
    identify_columns
)

app = FastAPI(title="Exam Student Matcher & Discrepancy Finder", version="1.0.0")

# Enable CORS for local network access
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Initialize Database on startup
init_db()

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
os.makedirs(STATIC_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

def get_lan_ip():
    """Get the local network IP so team members can connect."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

@app.get("/api/system-info")
def get_system_info():
    """Returns local network access URL and system stats."""
    lan_ip = get_lan_ip()
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) as count FROM master_students;")
    master_count = c.fetchone()["count"]
    
    c.execute("SELECT DISTINCT center_name FROM master_students WHERE center_name != '' ORDER BY center_name ASC;")
    centers = [r["center_name"] for r in c.fetchall()]
    
    conn.close()
    return {
        "local_url": "http://localhost:8000",
        "lan_url": f"http://{lan_ip}:8000",
        "lan_ip": lan_ip,
        "master_students_count": master_count,
        "centers": centers
    }

@app.get("/api/centers")
def get_centers():
    """Get list of centers in the master database."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT DISTINCT center_name FROM master_students WHERE center_name != '' ORDER BY center_name ASC;")
    centers = [r["center_name"] for r in c.fetchall()]
    conn.close()
    return {"centers": centers}

@app.post("/api/preview-sheet")
async def preview_sheet(file: UploadFile = File(...)):
    """Previews columns of an uploaded file to help verify column mapping."""
    content = await file.read()
    try:
        headers, rows = read_excel_or_csv(content, file.filename)
        detected_map = identify_columns(headers)
        return {
            "filename": file.filename,
            "headers": headers,
            "sample_rows": rows[:5],
            "total_rows": len(rows),
            "detected_mapping": detected_map
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {str(e)}")

@app.post("/api/upload-master")
async def upload_master_database(
    file: UploadFile = File(...),
    replace_existing: bool = Form(True)
):
    """Uploads and saves the master student list into SQLite."""
    content = await file.read()
    try:
        students = parse_master_student_sheet(content, file.filename)
        if not students:
            raise HTTPException(status_code=400, detail="No student records could be parsed. Check column headers.")
            
        conn = get_connection()
        try:
            c = conn.cursor()
            if replace_existing:
                c.execute("DELETE FROM master_students;")
                
            c.executemany("""
                INSERT INTO master_students (
                    student_id, student_number, last_6_digits, full_name, 
                    center_name, phone, seat_number, group_name, notes
                ) VALUES (
                    :student_id, :student_number, :last_6_digits, :full_name,
                    :center_name, :phone, :seat_number, :group_name, :notes
                );
            """, students)
            
            conn.commit()
            
            c.execute("SELECT COUNT(*) as total FROM master_students;")
            total_in_db = c.fetchone()["total"]
            
            c.execute("SELECT DISTINCT center_name FROM master_students WHERE center_name != '' ORDER BY center_name ASC;")
            centers = [r["center_name"] for r in c.fetchall()]
        finally:
            conn.close()
            
        return {
            "success": True,
            "message": f"Successfully loaded {len(students)} student records.",
            "total_students": total_in_db,
            "centers": centers
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error importing master database: {str(e)}")

@app.delete("/api/clear-master")
def clear_master():
    """Clear master database."""
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("DELETE FROM master_students;")
        conn.commit()
        return {"success": True, "message": "Master database cleared."}
    finally:
        conn.close()

@app.post("/api/match-scan")
async def match_scan(
    file: UploadFile = File(...),
    center_name: str = Form(...),
    uploader_name: Optional[str] = Form("Team Member")
):
    """
    Main matching endpoint:
    Team member selects the center, uploads the scan sheet.
    The system matches records using the last 6 digits.
    """
    content = await file.read()
    try:
        scanned_rows = parse_scan_sheet(content, file.filename)
        if not scanned_rows:
            raise HTTPException(status_code=400, detail="Uploaded scan sheet contains no readable rows.")
            
        conn = get_connection()
        try:
            c = conn.cursor()
            c.execute("SELECT * FROM master_students;")
            master_rows = [dict(r) for r in c.fetchall()]
            
            if not master_rows:
                raise HTTPException(status_code=400, detail="Master database is empty! Please upload the Master Student Database first.")
                
            # Execute Matching Engine
            match_result = match_scan_batch(
                scanned_rows=scanned_rows,
                master_students=master_rows,
                selected_center=center_name
            )
            
            # Save batch info in history
            summary = match_result["summary"]
            c.execute("""
                INSERT INTO scan_batches (
                    batch_name, center_name, uploaded_by, filename,
                    total_scanned, total_matched, total_unmatched, total_wrong_center
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                file.filename,
                center_name,
                uploader_name,
                file.filename,
                summary["total_scanned"],
                summary["matched_count"],
                summary["unmatched_count"],
                summary["wrong_center_count"]
            ))
            conn.commit()
            batch_id = c.lastrowid
        finally:
            conn.close()
            
        match_result["batch_id"] = batch_id
        match_result["filename"] = file.filename
        
        return match_result
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing scan matching: {str(e)}")

@app.post("/api/export-reconciliation")
async def export_reconciliation(payload: Dict[str, Any]):
    """Generates an Excel reconciliation report from match results."""
    try:
        match_result = payload.get("match_result")
        batch_name = payload.get("batch_name", "Reconciliation_Report")
        
        if not match_result:
            raise HTTPException(status_code=400, detail="Missing match result data")
            
        excel_io = generate_reconciliation_excel(match_result, batch_name)
        
        filename = f"{batch_name.replace(' ', '_')}_reconciliation.xlsx"
        return StreamingResponse(
            excel_io,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Export failed: {str(e)}")

@app.get("/api/master-students")
def get_master_students(
    search: Optional[str] = None, 
    center: Optional[str] = None, 
    limit: int = 100, 
    offset: int = 0
):
    """Returns paginated and searchable list of master students."""
    conn = get_connection()
    c = conn.cursor()
    base_sql = "FROM master_students WHERE 1=1"
    params = []
    
    if center and center.lower() != "all":
        base_sql += " AND center_name = ?"
        params.append(center)
        
    if search and search.strip():
        term = f"%{search.strip()}%"
        base_sql += " AND (full_name LIKE ? OR student_number LIKE ? OR last_6_digits LIKE ? OR phone LIKE ?)"
        params.extend([term, term, term, term])
        
    c.execute(f"SELECT COUNT(*) as count {base_sql}", params)
    total_count = c.fetchone()["count"]
    
    query = f"SELECT * {base_sql} ORDER BY id DESC LIMIT ? OFFSET ?"
    params.extend([limit, offset])
    c.execute(query, params)
    students = [dict(r) for r in c.fetchall()]
    conn.close()
    
    return {"total": total_count, "students": students}

@app.get("/", response_class=HTMLResponse)
def index_page():
    """Serves the Team Member Scanner interface by default."""
    member_path = os.path.join(STATIC_DIR, "member.html")
    if os.path.exists(member_path):
        with open(member_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    index_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse("<h2>Exam Matcher Web App Initializing...</h2>")

@app.get("/member", response_class=HTMLResponse)
def member_page():
    """Dedicated Team Member Exam Upload & Matching page."""
    member_path = os.path.join(STATIC_DIR, "member.html")
    if os.path.exists(member_path):
        with open(member_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return index_page()

@app.get("/admin", response_class=HTMLResponse)
@app.get("/master", response_class=HTMLResponse)
def master_admin_page():
    """Dedicated Master Student Database Management page."""
    master_path = os.path.join(STATIC_DIR, "master.html")
    if os.path.exists(master_path):
        with open(master_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse("<h2>Master Admin Page Initializing...</h2>")
