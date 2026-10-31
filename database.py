"""
Database models and connection for Exam Matcher.
Uses SQLite for zero-config persistence of Master Students and Scan Batches.
"""
import sqlite3
import os
from typing import List, Dict, Any, Optional

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exam_finder.db")

def get_connection():
    """
    Returns a configured SQLite connection with WAL mode enabled
    and a 30-second busy timeout to prevent Windows file locking issues.
    """
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
        seat_number TEXT,
        group_name TEXT,
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)
    
    # Create indexes for blazing fast 6-digit lookup and center filtering
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_master_last_6 ON master_students(last_6_digits);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_master_center ON master_students(center_name);")
    
    # Upload batches table to keep history
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
    
    # Scan records associated with a batch
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS scan_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch_id INTEGER,
        row_index INTEGER,
        raw_code TEXT,
        extracted_6_digits TEXT,
        scanned_name TEXT,
        score TEXT,
        status TEXT, -- 'matched', 'unmatched', 'wrong_center'
        matched_student_id INTEGER,
        extra_data TEXT,
        FOREIGN KEY(batch_id) REFERENCES scan_batches(id) ON DELETE CASCADE,
        FOREIGN KEY(matched_student_id) REFERENCES master_students(id) ON DELETE SET NULL
    );
    """)
    
    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully at", DB_FILE)
