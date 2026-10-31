"""
FastAPI application for Exam Finder.
Provides REST APIs for database management, center selection,
scan matching, and downloading discrepancy reports.
"""
import os
import io
import json
import socket
import urllib.parse
from typing import Optional, List, Dict, Any
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from database import init_db, get_connection
from matcher import match_scan_batch
from excel_service import (
    parse_master_student_sheet,
    parse_scan_sheet,
    generate_reconciliation_excel,
    generate_merged_center_excel,
    read_excel_or_csv,
    identify_columns
)

app = FastAPI(title="Exam Student Matcher & Discrepancy Finder", version="1.0.0")

# Enable CORS for local network and web access
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.middleware("http")
async def vercel_path_rewrite_middleware(request, call_next):
    """
    Normalizes Vercel serverless rewritten paths.
    Checks __vercel_route query param or x-matched-path headers.
    """
    route_param = request.query_params.get("__vercel_route")
    if route_param:
        request.scope["path"] = f"/api/{route_param.lstrip('/')}"
    else:
        orig_path = (
            request.headers.get("x-matched-path")
            or request.headers.get("x-forwarded-uri")
            or request.headers.get("x-invoke-path")
        )
        if orig_path:
            request.scope["path"] = orig_path.split("?")[0]
            
    return await call_next(request)

# Initialize Database on startup
init_db()

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
PUBLIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "public")
try:
    if not os.path.exists(STATIC_DIR):
        os.makedirs(STATIC_DIR, exist_ok=True)
except Exception:
    pass

if os.path.exists(STATIC_DIR):
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
            "centers": centers,
            "students": students
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error importing master database: {str(e)}")

@app.post("/api/restore-master-json")
async def restore_master_json(payload: Dict[str, Any]):
    """Restores master students from browser persistent backup into server DB."""
    students = payload.get("students", [])
    if not students:
        raise HTTPException(status_code=400, detail="No students data provided.")
        
    conn = get_connection()
    try:
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
        "message": f"Restored {total_in_db} students from persistent cache.",
        "total_students": total_in_db,
        "centers": centers
    }

@app.get("/api/all-master-students")
def get_all_master_students():
    """Returns all master students for browser backup."""
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM master_students ORDER BY id ASC;")
        students = [dict(r) for r in c.fetchall()]
        return {"students": students}
    finally:
        conn.close()

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
        
        # Auto-detect exam model from rows or filename
        detected_model = "Model"
        for r in scanned_rows:
            m_val = str(r.get("exam_model") or "").strip()
            if m_val and m_val.lower() not in ["exam", "none", "", "nan"]:
                detected_model = m_val
                break
        if detected_model == "Model":
            m_regex = re.search(r"(model\s*[a-zA-Z0-9]+|نموذج\s*[أ-يa-zA-Z0-9]+)", file.filename, re.IGNORECASE)
            if m_regex:
                detected_model = m_regex.group(1).title()
                
        match_result["detected_model"] = detected_model
        
        return match_result
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing scan matching: {str(e)}")

@app.post("/api/export-reconciliation")
async def export_reconciliation(payload: Dict[str, Any]):
    """Generates an Excel reconciliation report from match results with full Unicode filename support."""
    try:
        match_result = payload.get("match_result")
        batch_name = payload.get("batch_name", "Reconciliation_Report")
        
        if not match_result:
            raise HTTPException(status_code=400, detail="Missing match result data")
            
        excel_io = generate_reconciliation_excel(match_result, batch_name)
        
        safe_ascii = "reconciliation_report.xlsx"
        encoded_name = urllib.parse.quote(f"{batch_name}.xlsx")
        
        return Response(
            content=excel_io.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename=\"{safe_ascii}\"; filename*=UTF-8''{encoded_name}",
                "Access-Control-Expose-Headers": "Content-Disposition"
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Export failed: {str(e)}")

@app.post("/api/save-center-model")
async def save_center_model(payload: Dict[str, Any]):
    """
    Saves and merges a model scan submission under a center without overwriting other models.
    """
    center_name = payload.get("center_name")
    member_name = payload.get("member_name", "Team Member")
    exam_model = payload.get("exam_model", "Model")
    filename = payload.get("filename", "Scan.xlsx")
    records = payload.get("records", [])
    
    if not center_name or not records:
        raise HTTPException(status_code=400, detail="Center name and records are required.")
        
    conn = get_connection()
    try:
        c = conn.cursor()
        
        insert_rows = []
        for r in records:
            roll_no = r.get("roll_no") or r.get("raw_code") or r.get("extracted_6_digits")
            ext_6 = r.get("extracted_6_digits") or roll_no
            m_student = r.get("master_student") or {}
            
            insert_rows.append((
                center_name,
                exam_model,
                member_name,
                filename,
                str(roll_no),
                str(ext_6),
                str(r.get("score") or ""),
                m_student.get("full_name") or r.get("scanned_name") or "",
                r.get("status") or ("matched" if m_student else "unmatched"),
                m_student.get("id"),
                m_student.get("center_name") or "",
                m_student.get("phone") or ""
            ))
            
        c.executemany("""
            INSERT INTO center_submissions (
                center_name, exam_model, member_name, filename,
                roll_no, extracted_6_digits, score, student_name,
                status, matched_student_id, registered_center, phone
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, insert_rows)
        
        conn.commit()
        
        c.execute("""
            SELECT exam_model, member_name, COUNT(*) as count, MAX(created_at) as created_at
            FROM center_submissions
            WHERE center_name = ?
            GROUP BY exam_model, member_name
        """, (center_name,))
        models_summary = [dict(row) for row in c.fetchall()]
        
        c.execute("""
            SELECT COUNT(DISTINCT extracted_6_digits) as total_unique
            FROM center_submissions
            WHERE center_name = ?
        """, (center_name,))
        total_unique = c.fetchone()["total_unique"]
        
    finally:
        conn.close()
        
    return {
        "success": True,
        "message": f"Saved {len(records)} students for {exam_model} by {member_name} under {center_name}.",
        "models_summary": models_summary,
        "total_unique_students": total_unique
    }

@app.get("/api/center-models-status")
def get_center_models_status(center_name: Optional[str] = None):
    """Returns all models uploaded for a center or across all centers."""
    conn = get_connection()
    try:
        c = conn.cursor()
        if center_name and center_name.lower() != "all":
            c.execute("""
                SELECT center_name, exam_model, member_name, COUNT(*) as count, MAX(created_at) as created_at, MAX(filename) as filename
                FROM center_submissions
                WHERE center_name = ?
                GROUP BY center_name, exam_model, member_name
                ORDER BY created_at DESC
            """, (center_name,))
            rows = [dict(r) for r in c.fetchall()]
            
            c.execute("""
                SELECT COUNT(DISTINCT extracted_6_digits) as count 
                FROM center_submissions 
                WHERE center_name = ? AND status IN ('matched', 'wrong_center')
            """, (center_name,))
            unique_attended = c.fetchone()["count"]
            
            c.execute("SELECT COUNT(*) as count FROM master_students WHERE center_name = ?", (center_name,))
            total_registered = c.fetchone()["count"]
            
            c.execute("""
                SELECT COUNT(*) as count FROM master_students 
                WHERE center_name = ? AND last_6_digits IN (
                    SELECT DISTINCT extracted_6_digits FROM center_submissions WHERE center_name = ?
                )
            """, (center_name, center_name))
            registered_attended = c.fetchone()["count"]
            true_absent = max(0, total_registered - registered_attended)
            
            return {
                "center_name": center_name,
                "models": rows,
                "unique_attended": unique_attended,
                "total_registered": total_registered,
                "true_absent": true_absent
            }
        else:
            c.execute("""
                SELECT center_name, exam_model, member_name, COUNT(*) as count, MAX(created_at) as created_at
                FROM center_submissions
                GROUP BY center_name, exam_model, member_name
                ORDER BY center_name ASC, created_at DESC
            """)
            rows = [dict(r) for r in c.fetchall()]
            return {"centers_breakdown": rows}
    finally:
        conn.close()

@app.get("/api/export-center-final")
def export_center_final(center_name: str):
    """
    Generates and downloads the Consolidated Final Excel Workbook
    merging all models uploaded for this center.
    """
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""
            SELECT * FROM center_submissions 
            WHERE center_name = ? 
            ORDER BY id ASC
        """, (center_name,))
        submissions = [dict(r) for r in c.fetchall()]
        
        if not submissions:
            raise HTTPException(status_code=404, detail=f"No submissions found for center '{center_name}'")
            
        c.execute("""
            SELECT exam_model, member_name, COUNT(*) as count, MAX(created_at) as created_at, MAX(filename) as filename
            FROM center_submissions
            WHERE center_name = ?
            GROUP BY exam_model, member_name
        """, (center_name,))
        models_summary = [dict(r) for r in c.fetchall()]
        
        attended_last6 = {s["extracted_6_digits"] for s in submissions if s["extracted_6_digits"]}
        
        c.execute("SELECT * FROM master_students WHERE center_name = ?", (center_name,))
        master_for_center = [dict(r) for r in c.fetchall()]
        
        absent_students = [
            m for m in master_for_center
            if m["last_6_digits"] not in attended_last6
        ]
        
    finally:
        conn.close()
        
    excel_io = generate_merged_center_excel(
        center_name=center_name,
        submissions=submissions,
        absent_students=absent_students,
        models_summary=models_summary
    )
    
    safe_ascii = "center_final_merged.xlsx"
    encoded_name = urllib.parse.quote(f"{center_name}_Final_Merged.xlsx")
    
    return Response(
        content=excel_io.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename=\"{safe_ascii}\"; filename*=UTF-8''{encoded_name}",
            "Access-Control-Expose-Headers": "Content-Disposition"
        }
    )

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

def load_html(filename: str) -> str:
    for folder in [PUBLIC_DIR, STATIC_DIR]:
        path = os.path.join(folder, filename)
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return f.read()
    return f"<h2>Page {filename} not found</h2>"

@app.get("/", response_class=HTMLResponse)
@app.get("/member", response_class=HTMLResponse)
def member_page():
    """Serves the Team Member Scanner interface."""
    return HTMLResponse(content=load_html("member.html"))

@app.get("/master", response_class=HTMLResponse)
@app.get("/admin", response_class=HTMLResponse)
def master_admin_page():
    """Dedicated Master Student Database Management page."""
    return HTMLResponse(content=load_html("master.html"))
