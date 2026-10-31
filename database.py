"""
Database models and connection for Exam Matcher.
Supports local persistent SQLite and Vercel Serverless environment (/tmp).
"""
import sqlite3
import os
import tempfile
from typing import List, Dict, Any, Optional

def get_db_path() -> str:
    # If running on Vercel Serverless (or read-only filesystem), use /tmp
    if os.environ.get("VERCEL") or os.environ.get("AWS_LAMBDA_FUNCTION_NAME"):
        return "/tmp/exam_finder.db"
        
    local_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exam_finder.db")
    try:
        # Test if the directory is writable
        test_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".write_test")
        with open(test_file, "w") as f:
            f.write("1")
        if os.path.exists(test_file):
            os.remove(test_file)
        return local_path
    except Exception:
        # Fallback to system temp directory if local directory is read-only
        return os.path.join(tempfile.gettempdir(), "exam_finder.db")

DB_FILE = get_db_path()

def get_connection():
    """
    Returns a configured SQLite connection with a busy timeout
    to prevent file locking issues across environments.
    """
    global DB_FILE
    DB_FILE = get_db_path()
    
    conn = sqlite3.connect(DB_FILE, timeout=30.0, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA journal_mode = WAL;")
        conn.execute("PRAGMA busy_timeout = 30000;")
        conn.execute("PRAGMA synchronous = NORMAL;")
    except Exception:
        pass
    return conn

def init_db():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        
        # Master students table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS master_students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id TEXT,
            student_number TEXT NOT NULL,
            last_6_digits TEXT NOT NULL,
            full_name TEXT NOT NULL,
            center_name TEXT NOT NULL,
            phone TEXT,
            parent_phone TEXT,
            seat_number TEXT,
            group_name TEXT,
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)
        
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_master_last_6 ON master_students(last_6_digits);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_master_center ON master_students(center_name);")
        
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS scan_batches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_name TEXT NOT NULL,
            center_name TEXT NOT NULL,
            uploaded_by TEXT,
            filename TEXT NOT NULL,
            total_scanned INTEGER DEFAULT 0,
            total_matched INTEGER DEFAULT 0,
            total_unmatched INTEGER DEFAULT 0,
            total_wrong_center INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)
        
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS scan_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER,
            row_index INTEGER,
            raw_code TEXT,
            extracted_6_digits TEXT,
            scanned_name TEXT,
            score TEXT,
            status TEXT,
            matched_student_id INTEGER,
            extra_data TEXT,
            FOREIGN KEY(batch_id) REFERENCES scan_batches(id) ON DELETE CASCADE,
            FOREIGN KEY(matched_student_id) REFERENCES master_students(id) ON DELETE SET NULL
        );
        """)
        
        # Table for accumulating multiple models per center from different members
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS center_submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            center_name TEXT NOT NULL,
            exam_model TEXT NOT NULL,
            member_name TEXT NOT NULL,
            filename TEXT,
            roll_no TEXT NOT NULL,
            extracted_6_digits TEXT NOT NULL,
            score TEXT,
            student_name TEXT,
            status TEXT,
            matched_student_id INTEGER,
            registered_center TEXT,
            phone TEXT,
            parent_phone TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(matched_student_id) REFERENCES master_students(id) ON DELETE SET NULL
        );
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_sub_center ON center_submissions(center_name);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_sub_roll ON center_submissions(roll_no);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_sub_model ON center_submissions(exam_model);")
        
        # Ensure parent_phone column exists in existing databases
        try:
            cursor.execute("ALTER TABLE master_students ADD COLUMN parent_phone TEXT;")
        except Exception:
            pass
        try:
            cursor.execute("ALTER TABLE center_submissions ADD COLUMN parent_phone TEXT;")
        except Exception:
            pass
            
        conn.commit()
    finally:
        conn.close()

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully at", DB_FILE)
