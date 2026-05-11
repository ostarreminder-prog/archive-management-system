#!/usr/bin/env python3
"""Database migration script - add missing columns"""
import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), 'najm.db')

def migrate():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Check current columns
    c.execute("PRAGMA table_info(users)")
    user_columns = [row[1] for row in c.fetchall()]
    
    # Add last_otp_sent_at if missing
    if 'last_otp_sent_at' not in user_columns:
        print("Adding last_otp_sent_at to users table...")
        c.execute("ALTER TABLE users ADD COLUMN last_otp_sent_at TEXT")
        conn.commit()
        print("✓ Added last_otp_sent_at")
    else:
        print("✓ last_otp_sent_at already exists")
    
    # Check documents table
    c.execute("PRAGMA table_info(documents)")
    doc_columns = [row[1] for row in c.fetchall()]
    
    if 'original_file_sha256' not in doc_columns:
        print("Adding original_file_sha256 to documents table...")
        c.execute("ALTER TABLE documents ADD COLUMN original_file_sha256 TEXT")
        conn.commit()
        print("✓ Added original_file_sha256")
    else:
        print("✓ original_file_sha256 already exists")
    
    conn.close()
    print("\n✅ Migration completed!")

if __name__ == '__main__':
    migrate()
