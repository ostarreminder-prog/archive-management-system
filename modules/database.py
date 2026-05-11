import sqlite3
import os
import re
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'najm.db')

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def parse_device(user_agent: str) -> dict:
    ua = user_agent or ""
    if "Windows NT 10" in ua:      os_name = "Windows 10/11"
    elif "Windows NT 6.1" in ua:   os_name = "Windows 7"
    elif "Windows" in ua:          os_name = "Windows"
    elif "iPhone" in ua:           os_name = "iPhone"
    elif "iPad" in ua:             os_name = "iPad"
    elif "Android" in ua:
        ver = re.search(r'Android [\d.]+', ua)
        os_name = ver.group() if ver else "Android"
    elif "Mac OS X" in ua:         os_name = "macOS"
    elif "Linux" in ua:            os_name = "Linux"
    else:                          os_name = "غير معروف"

    if "Edg/" in ua:               browser = "Edge"
    elif "Chrome/" in ua:          browser = "Chrome"
    elif "Firefox/" in ua:         browser = "Firefox"
    elif "Safari/" in ua:          browser = "Safari"
    else:                          browser = "غير معروف"

    if any(x in ua for x in ["iPhone","Android","Mobile"]):
        device_type = "📱 جوال"
    elif "iPad" in ua:             device_type = "📟 تابلت"
    else:                          device_type = "💻 حاسب"

    return {"os": os_name, "browser": browser, "device_type": device_type,
            "summary": f"{device_type} — {os_name} — {browser}"}

def init_db():
    conn = get_db()
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS archive_sections (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        section_name  TEXT UNIQUE NOT NULL,
        section_code  TEXT UNIQUE NOT NULL,
        is_active     INTEGER DEFAULT 1,
        created_by    INTEGER,
        created_at    TEXT DEFAULT (datetime('now'))
    )''')

    c.execute(
        "INSERT OR IGNORE INTO archive_sections (section_name, section_code, is_active) VALUES ('عام', 'GN', 1)"
    )

    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id                INTEGER PRIMARY KEY AUTOINCREMENT,
        name              TEXT NOT NULL,
        email             TEXT UNIQUE NOT NULL,
        phone             TEXT NOT NULL,
        job_title         TEXT,
        role              TEXT DEFAULT 'user',
        password_hash     TEXT NOT NULL,
        sign_password     TEXT,
        sign_pass_changed TEXT,
        signature_path    TEXT,
        stamp_text        TEXT,
        archive_section_id INTEGER,
        is_active         INTEGER DEFAULT 1,
        first_login       INTEGER DEFAULT 1,
        created_by        INTEGER,
        created_at        TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(created_by) REFERENCES users(id),
        FOREIGN KEY(archive_section_id) REFERENCES archive_sections(id)
    )''')

    user_columns = {row[1] for row in c.execute("PRAGMA table_info(users)").fetchall()}
    if 'archive_section_id' not in user_columns:
        c.execute("ALTER TABLE users ADD COLUMN archive_section_id INTEGER")
    if 'employee_id' not in user_columns:
        c.execute("ALTER TABLE users ADD COLUMN employee_id TEXT")
    if 'stamp_visibility_scope' not in user_columns:
        c.execute("ALTER TABLE users ADD COLUMN stamp_visibility_scope TEXT DEFAULT 'self'")
    if 'stamp_visible_to_user_id' not in user_columns:
        c.execute("ALTER TABLE users ADD COLUMN stamp_visible_to_user_id INTEGER")
    if 'last_otp_sent_at' not in user_columns:
        c.execute("ALTER TABLE users ADD COLUMN last_otp_sent_at TEXT")

    c.execute('''CREATE TABLE IF NOT EXISTS user_section_permissions (
        id               INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id          INTEGER NOT NULL,
        section_id       INTEGER NOT NULL,
        can_view_archive INTEGER DEFAULT 0,
        can_stamp        INTEGER DEFAULT 0,
        granted_by       INTEGER,
        created_at       TEXT DEFAULT (datetime('now')),
        UNIQUE(user_id, section_id),
        FOREIGN KEY(user_id)    REFERENCES users(id),
        FOREIGN KEY(section_id) REFERENCES archive_sections(id)
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS otp_codes (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id     INTEGER NOT NULL,
        code        TEXT NOT NULL,
        expires_at  TEXT NOT NULL,
        used        INTEGER DEFAULT 0,
        FOREIGN KEY(user_id) REFERENCES users(id)
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS trusted_devices (
        id                   INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id              INTEGER NOT NULL,
        device_hash          TEXT NOT NULL,
        device_name          TEXT,
        ip_address           TEXT,
        trusted              INTEGER DEFAULT 0,
        token                TEXT UNIQUE,
        last_otp_verified_at TEXT,
        created_at           TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(user_id) REFERENCES users(id)
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS documents (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        archive_number  TEXT UNIQUE,
        archive_section TEXT DEFAULT 'عام',
        archive_section_code TEXT DEFAULT 'GN',
        title           TEXT NOT NULL,
        template_name   TEXT,
        content_json    TEXT,
        file_path       TEXT,
        status          TEXT DEFAULT 'draft',
        notes           TEXT,
        created_by      INTEGER,
        approved_by     INTEGER,
        created_at      TEXT DEFAULT (datetime('now')),
        approved_at     TEXT,
        file_sha256           TEXT,
        archive_storage_path  TEXT,
        archived_at           TEXT,
        archived_by           INTEGER,
        original_file_sha256  TEXT,
        FOREIGN KEY(created_by) REFERENCES users(id),
        FOREIGN KEY(approved_by) REFERENCES users(id),
        FOREIGN KEY(archived_by) REFERENCES users(id)
    )''')

    doc_columns = {row[1] for row in c.execute("PRAGMA table_info(documents)").fetchall()}
    if 'archive_section' not in doc_columns:
        c.execute("ALTER TABLE documents ADD COLUMN archive_section TEXT DEFAULT 'عام'")
    if 'archive_section_code' not in doc_columns:
        c.execute("ALTER TABLE documents ADD COLUMN archive_section_code TEXT DEFAULT 'GN'")
    if 'archive_storage_path' not in doc_columns:
        c.execute("ALTER TABLE documents ADD COLUMN archive_storage_path TEXT")
    if 'archived_at' not in doc_columns:
        c.execute("ALTER TABLE documents ADD COLUMN archived_at TEXT")
    if 'archived_by' not in doc_columns:
        c.execute("ALTER TABLE documents ADD COLUMN archived_by INTEGER")
    if 'file_sha256' not in doc_columns:
        c.execute("ALTER TABLE documents ADD COLUMN file_sha256 TEXT")
    if 'original_file_sha256' not in doc_columns:
        c.execute("ALTER TABLE documents ADD COLUMN original_file_sha256 TEXT")

    c.execute("UPDATE documents SET archive_section='عام' WHERE archive_section IS NULL OR TRIM(archive_section)=''")
    c.execute("UPDATE documents SET archive_section_code=UPPER(TRIM(archive_section_code)) WHERE archive_section_code IS NOT NULL")
    c.execute("UPDATE documents SET archive_section_code='GN' WHERE UPPER(COALESCE(archive_section_code,'')) NOT GLOB '[A-Z][A-Z]'")

    existing_doc_sections = c.execute(
        """
        SELECT DISTINCT TRIM(COALESCE(archive_section, 'عام')) AS section_name,
               UPPER(TRIM(COALESCE(archive_section_code, 'GN'))) AS section_code
        FROM documents
        WHERE TRIM(COALESCE(archive_section_code, '')) != ''
        """
    ).fetchall()
    for row in existing_doc_sections:
        section_name = (row[0] or 'عام').strip() or 'عام'
        section_code = (row[1] or 'GN').strip().upper() or 'GN'
        if not re.fullmatch(r'[A-Z]{2}', section_code):
            continue
        c.execute(
            "INSERT OR IGNORE INTO archive_sections (section_name, section_code, is_active) VALUES (?,?,1)",
            (section_name, section_code)
        )

    default_section_row = c.execute(
        "SELECT id FROM archive_sections WHERE UPPER(section_code)='GN' LIMIT 1"
    ).fetchone()
    default_section_id = default_section_row[0] if default_section_row else None
    if default_section_id is not None:
        c.execute(
            "UPDATE users SET archive_section_id=? WHERE archive_section_id IS NULL",
            (default_section_id,)
        )

    # Migration: add last_otp_verified_at to trusted_devices
    trusted_devices_columns = {row[1] for row in c.execute("PRAGMA table_info(trusted_devices)").fetchall()}
    if 'last_otp_verified_at' not in trusted_devices_columns:
        c.execute("ALTER TABLE trusted_devices ADD COLUMN last_otp_verified_at TEXT")

    c.execute('''CREATE TABLE IF NOT EXISTS signature_requests (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id     INTEGER NOT NULL,
        requested_from  INTEGER NOT NULL,
        requested_by    INTEGER NOT NULL,
        signature_owner_id INTEGER,
        stamp_owner_id  INTEGER,
        signature_asset_id INTEGER,
        stamp_asset_id  INTEGER,
        positions_json  TEXT,
        include_qr      INTEGER DEFAULT 1,
        serial_number   TEXT UNIQUE,
        sign_type       TEXT DEFAULT 'signature',
        status          TEXT DEFAULT 'pending',
        sign_hash       TEXT,
        signed_at       TEXT,
        message         TEXT,
        FOREIGN KEY(document_id) REFERENCES documents(id),
        FOREIGN KEY(requested_from) REFERENCES users(id),
        FOREIGN KEY(requested_by) REFERENCES users(id)
    )''')

    sig_columns = {row[1] for row in c.execute("PRAGMA table_info(signature_requests)").fetchall()}
    if 'signature_owner_id' not in sig_columns:
        c.execute("ALTER TABLE signature_requests ADD COLUMN signature_owner_id INTEGER")
    if 'stamp_owner_id' not in sig_columns:
        c.execute("ALTER TABLE signature_requests ADD COLUMN stamp_owner_id INTEGER")
    if 'signature_asset_id' not in sig_columns:
        c.execute("ALTER TABLE signature_requests ADD COLUMN signature_asset_id INTEGER")
    if 'stamp_asset_id' not in sig_columns:
        c.execute("ALTER TABLE signature_requests ADD COLUMN stamp_asset_id INTEGER")
    if 'positions_json' not in sig_columns:
        c.execute("ALTER TABLE signature_requests ADD COLUMN positions_json TEXT")
    if 'include_qr' not in sig_columns:
        c.execute("ALTER TABLE signature_requests ADD COLUMN include_qr INTEGER DEFAULT 1")

    c.execute('''CREATE TABLE IF NOT EXISTS login_history (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id      INTEGER NOT NULL,
        event        TEXT NOT NULL,
        ip_address   TEXT,
        user_agent   TEXT,
        os_name      TEXT,
        browser      TEXT,
        device_type  TEXT,
        timestamp    TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(user_id) REFERENCES users(id)
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS audit_log (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id      INTEGER,
        document_id  INTEGER,
        action       TEXT NOT NULL,
        details      TEXT,
        ip_address   TEXT,
        user_agent   TEXT,
        os_name      TEXT,
        browser      TEXT,
        device_type  TEXT,
        timestamp    TEXT DEFAULT (datetime('now'))
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS stamp_assets (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id     INTEGER NOT NULL,
        stamp_name  TEXT,
        stamp_path  TEXT NOT NULL,
        visibility_scope TEXT DEFAULT 'self',
        visible_to_user_id INTEGER,
        is_active   INTEGER DEFAULT 1,
        created_at  TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(user_id) REFERENCES users(id)
    )''')

    stamp_asset_columns = {row[1] for row in c.execute("PRAGMA table_info(stamp_assets)").fetchall()}
    if 'visibility_scope' not in stamp_asset_columns:
        c.execute("ALTER TABLE stamp_assets ADD COLUMN visibility_scope TEXT DEFAULT 'self'")
    if 'visible_to_user_id' not in stamp_asset_columns:
        c.execute("ALTER TABLE stamp_assets ADD COLUMN visible_to_user_id INTEGER")

    # تصحيح السجلات القديمة التي كانت visibility_scope=NULL أو فارغة فقط (دون تغيير القيم الصحيحة)
    c.execute("UPDATE stamp_assets SET visibility_scope='self' WHERE TRIM(COALESCE(visibility_scope,''))='' OR visibility_scope IS NULL")
    c.execute("UPDATE stamp_assets SET visibility_scope='self' WHERE visibility_scope NOT IN ('all','managers','specific','self')")

    c.execute('''CREATE TABLE IF NOT EXISTS signature_assets (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id         INTEGER NOT NULL,
        signature_name  TEXT,
        signature_path  TEXT NOT NULL,
        visibility_scope TEXT DEFAULT 'self',
        visible_to_user_id INTEGER,
        is_active       INTEGER DEFAULT 1,
        created_at      TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(user_id) REFERENCES users(id)
    )''')

    signature_asset_columns = {row[1] for row in c.execute("PRAGMA table_info(signature_assets)").fetchall()}
    if 'visibility_scope' not in signature_asset_columns:
        c.execute("ALTER TABLE signature_assets ADD COLUMN visibility_scope TEXT DEFAULT 'self'")
    if 'visible_to_user_id' not in signature_asset_columns:
        c.execute("ALTER TABLE signature_assets ADD COLUMN visible_to_user_id INTEGER")

    # تصحيح السجلات القديمة التي كانت visibility_scope=NULL أو فارغة فقط
    c.execute("UPDATE signature_assets SET visibility_scope='self' WHERE TRIM(COALESCE(visibility_scope,''))='' OR visibility_scope IS NULL")
    c.execute("UPDATE signature_assets SET visibility_scope='self' WHERE visibility_scope NOT IN ('all','managers','specific','self')")

    c.execute('''CREATE TABLE IF NOT EXISTS app_settings (
        setting_key   TEXT PRIMARY KEY,
        setting_value TEXT,
        updated_at    TEXT DEFAULT (datetime('now'))
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS stamp_templates (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        name         TEXT NOT NULL,
        file_name    TEXT NOT NULL,
        text_x_ratio REAL DEFAULT 0.25,
        text_y_ratio REAL DEFAULT 0.08,
        is_default   INTEGER DEFAULT 0,
        created_at   TEXT DEFAULT (datetime('now'))
    )''')

    legacy_signature_rows = c.execute(
        """
        SELECT id, signature_path
        FROM users
        WHERE TRIM(COALESCE(signature_path, '')) <> ''
        """
    ).fetchall()
    for row in legacy_signature_rows:
        user_id = row[0]
        signature_path = (row[1] or '').strip()
        if not signature_path:
            continue

        existing = c.execute(
            "SELECT 1 FROM signature_assets WHERE user_id=? AND signature_path=? LIMIT 1",
            (user_id, signature_path)
        ).fetchone()
        if existing:
            continue

        c.execute(
            "INSERT INTO signature_assets (user_id, signature_name, signature_path, visibility_scope, is_active) VALUES (?,?,?,?,1)",
            (user_id, 'التوقيع الأساسي', signature_path, 'all')
        )

    conn.commit(); conn.close()
    print("[NAJM] OK - قاعدة البيانات جاهزة")

def log_login(user_id, event, ip="", user_agent=""):
    device = parse_device(user_agent)
    conn   = get_db()
    conn.execute(
        "INSERT INTO login_history (user_id,event,ip_address,user_agent,os_name,browser,device_type) VALUES (?,?,?,?,?,?,?)",
        (user_id, event, ip, user_agent, device['os'], device['browser'], device['device_type'])
    )
    conn.commit(); conn.close()

def log_action(user_id, action, details="", document_id=None, ip="", user_agent=""):
    device = parse_device(user_agent)
    conn   = get_db()
    conn.execute(
        "INSERT INTO audit_log (user_id,document_id,action,details,ip_address,user_agent,os_name,browser,device_type) VALUES (?,?,?,?,?,?,?,?,?)",
        (user_id, document_id, action, details, ip, user_agent, device['os'], device['browser'], device['device_type'])
    )
    conn.commit(); conn.close()

def get_login_history(user_id=None, limit=50):
    conn = get_db()
    if user_id:
        rows = conn.execute("SELECT l.*, u.name as user_name FROM login_history l LEFT JOIN users u ON l.user_id=u.id WHERE l.user_id=? ORDER BY l.id DESC LIMIT ?", (user_id, limit)).fetchall()
    else:
        rows = conn.execute("SELECT l.*, u.name as user_name FROM login_history l LEFT JOIN users u ON l.user_id=u.id ORDER BY l.id DESC LIMIT ?", (limit,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]