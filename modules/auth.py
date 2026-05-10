import bcrypt
import secrets
import hashlib
import re
from datetime import datetime, timedelta
from functools import wraps
from flask import session, redirect, url_for, request, jsonify
from .database import get_db, log_action, parse_device

# ─── باسورد ──────────────────────────────────
def hash_password(p): return bcrypt.hashpw(p.encode(), bcrypt.gensalt()).decode()
def check_password(p, h): return bcrypt.checkpw(p.encode(), h.encode())

# ─── OTP ─────────────────────────────────────
def generate_otp(user_id):
    code = str(secrets.randbelow(900000) + 100000)
    exp  = (datetime.utcnow() + timedelta(minutes=1)).isoformat()
    conn = get_db()
    conn.execute("UPDATE otp_codes SET used=1 WHERE user_id=? AND used=0", (user_id,))
    conn.execute("INSERT INTO otp_codes (user_id,code,expires_at) VALUES (?,?,?)", (user_id,code,exp))
    conn.commit(); conn.close()
    return code

def verify_otp(user_id, code):
    conn = get_db()
    row  = conn.execute(
        "SELECT * FROM otp_codes WHERE user_id=? AND code=? AND used=0 ORDER BY id DESC LIMIT 1",
        (user_id, code)
    ).fetchone()
    if not row: conn.close(); return False
    if datetime.utcnow() > datetime.fromisoformat(row['expires_at']):
        conn.close(); return False
    conn.execute("UPDATE otp_codes SET used=1 WHERE id=?", (row['id'],))
    conn.commit(); conn.close()
    return True

# ─── الأجهزة الموثوقة ────────────────────────
def get_device_hash(ip, user_agent):
    device = parse_device(user_agent)
    normalized = "|".join([
        (ip or "").strip().lower(),
        (device.get("device_type") or "").strip().lower(),
        (device.get("os") or "").strip().lower(),
        (device.get("browser") or "").strip().lower(),
    ])
    return hashlib.sha256(normalized.encode()).hexdigest()

def is_trusted_device(user_id, device_hash, ip="", user_agent=""):
    conn = get_db()
    device = parse_device(user_agent)
    row  = conn.execute(
        "SELECT * FROM trusted_devices WHERE user_id=? AND device_hash=? AND trusted=1",
        (user_id, device_hash)
    ).fetchone()

    if not row:
        legacy_hash = hashlib.sha256(f"{ip}:{user_agent}".encode()).hexdigest()
        if legacy_hash != device_hash:
            legacy_row = conn.execute(
                "SELECT * FROM trusted_devices WHERE user_id=? AND device_hash=? AND trusted=1",
                (user_id, legacy_hash)
            ).fetchone()
            if legacy_row:
                conn.execute(
                    "UPDATE trusted_devices SET device_hash=?, device_name=?, ip_address=? WHERE id=?",
                    (device_hash, device.get("summary"), ip, legacy_row['id'])
                )
                conn.commit()
                row = legacy_row

    if not row:
        compat_row = conn.execute(
            "SELECT * FROM trusted_devices WHERE user_id=? AND trusted=1 AND device_name=? AND ip_address=? ORDER BY id DESC LIMIT 1",
            (user_id, device.get("summary"), ip)
        ).fetchone()
        if compat_row:
            conn.execute(
                "UPDATE trusted_devices SET device_hash=? WHERE id=?",
                (device_hash, compat_row['id'])
            )
            conn.commit()
            row = compat_row

    conn.close()
    return bool(row)

def register_pending_device(user_id, device_hash, ip, user_agent):
    token  = secrets.token_urlsafe(32)
    device = parse_device(user_agent)
    conn   = get_db()
    conn.execute(
        "DELETE FROM trusted_devices WHERE user_id=? AND device_hash=? AND trusted=0",
        (user_id, device_hash)
    )
    conn.execute(
        "INSERT INTO trusted_devices (user_id,device_hash,device_name,ip_address,token) VALUES (?,?,?,?,?)",
        (user_id, device_hash, device['summary'], ip, token)
    )
    conn.commit(); conn.close()
    return token, device['summary']

def trust_device_by_token(token):
    conn = get_db()
    row  = conn.execute(
        "SELECT * FROM trusted_devices WHERE token=? AND trusted=0", (token,)
    ).fetchone()
    if not row: conn.close(); return None
    conn.execute("UPDATE trusted_devices SET trusted=1 WHERE token=?", (token,))
    conn.commit(); conn.close()
    return dict(row)

def block_user(user_id):
    conn = get_db()
    conn.execute("UPDATE users SET is_active=0 WHERE id=?", (user_id,))
    conn.commit(); conn.close()

def get_trusted_devices(user_id):
    conn  = get_db()
    rows  = conn.execute(
        "SELECT * FROM trusted_devices WHERE user_id=? ORDER BY id DESC", (user_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]

# ─── رقم تسلسلي ──────────────────────────────
def generate_serial(prefix, table, column, section_code=None):
    conn = get_db()
    now = datetime.utcnow()

    if str(prefix or '').upper() == 'DOC':
        normalized_section_code = str(section_code or '').strip().upper()
        if normalized_section_code:
            if not re.fullmatch(r'[A-Z]{2}', normalized_section_code):
                conn.close()
                raise ValueError('ARCHIVE_SECTION_CODE_INVALID')

            rows = conn.execute(
                f"SELECT {column} FROM {table} WHERE UPPER(COALESCE({column}, '')) LIKE ?",
                (f"{normalized_section_code}%",)
            ).fetchall()

            max_number = 0
            pattern = re.compile(rf"{re.escape(normalized_section_code)}(\d{{2,3}})")
            for row in rows:
                serial_value = str(row[0] or '').strip().upper()
                matched = pattern.fullmatch(serial_value)
                if not matched:
                    continue
                current_number = int(matched.group(1))
                if current_number > max_number:
                    max_number = current_number

            next_number = max_number + 1
            if next_number > 999:
                conn.close()
                raise ValueError('ARCHIVE_SECTION_SERIAL_LIMIT')

            number_part = str(next_number).zfill(2 if next_number < 100 else 3)
            conn.close()
            return f"{normalized_section_code}{number_part}"

        last_numeric = conn.execute(
            f"""
            SELECT MAX(CAST({column} AS INTEGER))
            FROM {table}
            WHERE {column} GLOB '[0-9][0-9][0-9][0-9][0-9][0-9]'
            """
        ).fetchone()[0]

        if last_numeric is None:
            last_numeric = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]

        next_number = int(last_numeric) + 1
        conn.close()
        return str(next_number).zfill(6)

    if str(prefix or '').upper() == 'SIG':
        day_key = now.strftime('%Y%m%d')
        count = conn.execute(
            f"SELECT COUNT(*) FROM {table} WHERE {column} LIKE ?",
            (f"SIG-{day_key}-%",)
        ).fetchone()[0]
        conn.close()
        return f"SIG-{day_key}-{now.strftime('%H%M%S')}-{str(count+1).zfill(4)}"

    count = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
    conn.close()
    return f"{str(prefix or 'SER').upper()}-{str(count+1).zfill(6)}"

def generate_temp_password():
    return secrets.token_hex(4).upper()

def verify_sign_password(user_id, password):
    user = get_user_by_id(user_id)
    if not user or not user['sign_password']: return False
    return check_password(password, user['sign_password'])

def generate_sign_hash(user_id, doc_id, serial, sign_password):
    raw = f"{user_id}:{doc_id}:{serial}:{sign_password}:{datetime.utcnow().date()}"
    return hashlib.sha256(raw.encode()).hexdigest()[:16].upper()

# ─── Decorators ──────────────────────────────
def _wants_json_response():
    accept = (request.headers.get('Accept') or '').lower()
    return request.path.startswith('/api/') or request.is_json or 'application/json' in accept

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return (jsonify({"error":"غير مصرح"}),401) if _wants_json_response() else redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if session.get('user_role') not in roles:
                return (jsonify({"error":"ليس لديك صلاحية"}),403) if _wants_json_response() else redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return decorated
    return decorator

# ─── CRUD ────────────────────────────────────
def get_all_archive_sections(include_inactive=False):
    conn = get_db()
    query = "SELECT id, section_name, section_code, is_active, created_by, created_at FROM archive_sections"
    params = []
    if not include_inactive:
        query += " WHERE is_active=1"
    query += " ORDER BY section_name ASC, section_code ASC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_archive_section_by_id(section_id):
    if not section_id:
        return None
    conn = get_db()
    row = conn.execute(
        "SELECT id, section_name, section_code, is_active, created_by, created_at FROM archive_sections WHERE id=? LIMIT 1",
        (section_id,)
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def get_default_archive_section():
    conn = get_db()
    row = conn.execute(
        """
        SELECT id, section_name, section_code, is_active, created_by, created_at
        FROM archive_sections
        WHERE is_active=1
        ORDER BY CASE WHEN UPPER(section_code)='GN' THEN 0 ELSE 1 END, id ASC
        LIMIT 1
        """
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def create_archive_section(section_name, section_code, created_by=None):
    normalized_name = str(section_name or '').strip()
    normalized_code = str(section_code or '').strip().upper()

    if not normalized_name:
        return False, None, "اسم القسم مطلوب"
    if not re.fullmatch(r'[A-Z]{2}', normalized_code):
        return False, None, "اختصار القسم يجب أن يكون حرفين إنجليزيين"

    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO archive_sections (section_name, section_code, created_by) VALUES (?,?,?)",
            (normalized_name, normalized_code, created_by)
        )
        conn.commit()
        row = conn.execute(
            "SELECT id, section_name, section_code, is_active, created_by, created_at FROM archive_sections WHERE id=last_insert_rowid()"
        ).fetchone()
        return True, dict(row) if row else None, ""
    except Exception as e:
        message = str(e)
        if 'UNIQUE constraint failed' in message and 'section_code' in message:
            return False, None, "اختصار القسم مستخدم مسبقاً"
        if 'UNIQUE constraint failed' in message and 'section_name' in message:
            return False, None, "اسم القسم موجود مسبقاً"
        return False, None, message
    finally:
        conn.close()


def create_user(name, email, phone, job_title, role, created_by, archive_section_id=None, employee_id=None):
    temp_pass = generate_temp_password()
    conn = get_db()
    try:
        resolved_section_id = archive_section_id
        if resolved_section_id:
            section_row = conn.execute(
                "SELECT id FROM archive_sections WHERE id=? AND is_active=1 LIMIT 1",
                (resolved_section_id,)
            ).fetchone()
            if not section_row:
                return False, "", "القسم المحدد غير موجود"
            resolved_section_id = section_row['id']
        else:
            section_row = conn.execute(
                "SELECT id FROM archive_sections WHERE is_active=1 ORDER BY CASE WHEN UPPER(section_code)='GN' THEN 0 ELSE 1 END, id ASC LIMIT 1"
            ).fetchone()
            resolved_section_id = section_row['id'] if section_row else None

        normalized_employee_id = str(employee_id or '').strip() or None

        conn.execute(
            "INSERT INTO users (name,email,phone,job_title,role,password_hash,first_login,created_by,archive_section_id,employee_id) VALUES (?,?,?,?,?,?,1,?,?,?)",
            (name, email, phone, job_title, role, hash_password(temp_pass), created_by, resolved_section_id, normalized_employee_id)
        )
        conn.commit()
        new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        return True, temp_pass, "", new_id
    except Exception as e:
        return False, "", str(e), None
    finally:
        conn.close()

def get_user_by_email(email):
    conn = get_db()
    u = conn.execute(
        """
        SELECT u.*, s.section_name AS archive_section, s.section_code AS archive_section_code
        FROM users u
        LEFT JOIN archive_sections s ON s.id = u.archive_section_id
        WHERE u.email=? AND u.is_active=1
        """,
        (email,)
    ).fetchone()
    conn.close(); return u

def get_user_by_id(uid):
    conn = get_db()
    u = conn.execute(
        """
        SELECT u.*, s.section_name AS archive_section, s.section_code AS archive_section_code
        FROM users u
        LEFT JOIN archive_sections s ON s.id = u.archive_section_id
        WHERE u.id=?
        """,
        (uid,)
    ).fetchone()
    conn.close(); return u

def get_all_users():
    conn = get_db()
    rows = conn.execute(
        """
        SELECT u.id, u.name, u.email, u.phone, u.job_title, u.role,
               u.signature_path, u.stamp_text, u.created_at, u.is_active,
               u.archive_section_id, u.employee_id, u.created_by,
               s.section_name AS archive_section,
               s.section_code AS archive_section_code
        FROM users u
        LEFT JOIN archive_sections s ON s.id = u.archive_section_id
        ORDER BY u.id DESC
        """
    ).fetchall()
    conn.close(); return [dict(r) for r in rows]

def update_password(user_id, new_password):
    conn = get_db()
    conn.execute("UPDATE users SET password_hash=?, first_login=0 WHERE id=?", (hash_password(new_password), user_id))
    conn.commit(); conn.close()

def update_sign_password(user_id, new_password):
    conn = get_db()
    conn.execute("UPDATE users SET sign_password=?, sign_pass_changed=? WHERE id=?",
                 (hash_password(new_password), datetime.utcnow().isoformat(), user_id))
    conn.commit(); conn.close()

def update_signature(user_id, path):
    conn = get_db()
    conn.execute("UPDATE users SET signature_path=? WHERE id=?", (path, user_id))
    conn.commit(); conn.close()

def can_use_stamp(user_id):
    user = get_user_by_id(user_id)
    if not user:
        return False
    if user['stamp_text']:
        return True
    conn = get_db()
    row = conn.execute(
        "SELECT 1 FROM stamp_assets WHERE user_id=? AND is_active=1 LIMIT 1",
        (user_id,)
    ).fetchone()
    conn.close()
    return bool(row)