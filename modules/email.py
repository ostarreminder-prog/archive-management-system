import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASSWORD")
FROM_NAME = os.getenv("FROM_NAME", "SignMy")

def send_email(to_email, subject, html_body):
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = f"{FROM_NAME} <{SMTP_USER}>"
        msg["To"]      = to_email
        msg.attach(MIMEText(html_body, "html", "utf-8"))
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.ehlo(); s.starttls(); s.ehlo()
            s.login(SMTP_USER, SMTP_PASS)
            s.sendmail(SMTP_USER, to_email, msg.as_string())
        print(f"[NAJM] OK email sent to {to_email}")
        return True
    except Exception as e:
        print(f"[NAJM] ERROR email failed: {e}")
        return False

def _wrap(content, title, color="#1e40af"):
    return f"""<!DOCTYPE html><html dir="rtl" lang="ar">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,sans-serif;direction:rtl">
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td align="center" style="padding:40px 20px">
<table width="560" cellpadding="0" cellspacing="0"
       style="background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08)">
  <tr><td style="background:linear-gradient(135deg,{color},{color}dd);padding:28px 32px;text-align:center">
    <div style="font-size:32px;margin-bottom:6px">⭐</div>
    <div style="color:#fff;font-size:20px;font-weight:700">SignMy</div>
    <div style="color:rgba(255,255,255,.7);font-size:13px;margin-top:4px">{title}</div>
  </td></tr>
  <tr><td style="padding:32px">{content}</td></tr>
  <tr><td style="background:#f9fafb;padding:18px 32px;border-top:1px solid #e5e7eb;text-align:center">
    <div style="color:#9ca3af;font-size:12px">
      هذا إيميل تلقائي من نظام SignMy — لا ترد عليه<br>
      {datetime.now().strftime('%Y-%m-%d %H:%M')}
    </div>
  </td></tr>
</table></td></tr></table></body></html>"""

# ─── 1. حساب جديد ────────────────────────────
def send_welcome(to_email, name, temp_password, role):
    role_ar = {"user": "موظف", "manager": "مدير", "admin": "مدير عام"}.get(role, role)
    content = f"""
    <p style="color:#374151;font-size:15px">مرحباً <strong>{name}</strong> 👋</p>
    <p style="color:#6b7280;font-size:14px">تم إنشاء حسابك في نظام <strong>SignMy</strong></p>
    <table width="100%" style="background:#eff6ff;border-radius:10px;padding:18px;margin:16px 0">
      <tr><td style="padding:6px 0">
        <span style="color:#6b7280;font-size:13px">📧 البريد الإلكتروني:</span>
        <strong style="color:#1e40af;font-size:14px;margin-right:8px">{to_email}</strong>
      </td></tr>
      <tr><td style="padding:6px 0">
        <span style="color:#6b7280;font-size:13px">🔑 الباسورد المؤقت:</span>
        <strong style="color:#1e40af;font-size:20px;margin-right:8px;font-family:monospace;letter-spacing:4px">{temp_password}</strong>
      </td></tr>
      <tr><td style="padding:6px 0">
        <span style="color:#6b7280;font-size:13px">👤 الدور:</span>
        <strong style="color:#374151;font-size:14px;margin-right:8px">{role_ar}</strong>
      </td></tr>
    </table>
    <div style="background:#fef2f2;border-radius:8px;padding:14px;border-right:4px solid #ef4444">
      <p style="color:#991b1b;font-size:13px;margin:0">
        ⚠️ هذا باسورد مؤقت — ستُطلب منك تغييره عند أول دخول<br>
        🔒 لا تشارك هذا الباسورد مع أحد
      </p>
    </div>"""
    return send_email(to_email, "مرحباً بك في SignMy — بيانات حسابك", _wrap(content, "حساب جديد"))

# ─── 2. OTP ──────────────────────────────────
def send_otp(to_email, name, code):
    content = f"""
    <p style="color:#374151;font-size:15px">مرحباً <strong>{name}</strong>،</p>
    <p style="color:#6b7280;font-size:14px">كود تسجيل الدخول:</p>
    <div style="background:#eff6ff;border-radius:12px;padding:28px;text-align:center;margin:20px 0">
      <div style="letter-spacing:12px;font-size:42px;font-weight:900;color:#1e40af;font-family:monospace">{code}</div>
    </div>
    <div style="background:#fef2f2;border-radius:8px;padding:14px;border-right:4px solid #ef4444">
      <p style="color:#991b1b;font-size:13px;margin:0">
        ⏱️ صالح دقيقة واحدة فقط<br>
        🔒 لا تشارك هذا الكود مع أحد
      </p>
    </div>"""
    return send_email(to_email, "🔐 كود تسجيل الدخول — SignMy", _wrap(content, "كود التحقق"))

# ─── 3. إشعار استخدام التوقيع ────────────────
def send_sign_used(to_email, name, doc_title, archive_num, serial, used_by,
                   sign_type="signature", device_name=None, used_at=None, file_name=None):
    type_map = {
        "signature": "توقيع",
        "stamp": "ختم",
        "both": "توقيع + ختم",
    }
    type_label = type_map.get(sign_type, sign_type)
    used_at_value = used_at or datetime.now().strftime('%Y-%m-%d %H:%M')
    extra_rows = ""
    if file_name:
        extra_rows += f"""<tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">📎 الملف المستخدم:</span>
        <strong style="color:#111827;margin-right:8px">{file_name}</strong></td></tr>"""
    if device_name:
        extra_rows += f"""<tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">💻 الجهاز:</span>
        <strong style="color:#374151;margin-right:8px">{device_name}</strong></td></tr>"""

    content = f"""
    <p style="color:#374151;font-size:15px">مرحباً <strong>{name}</strong>،</p>
    <p style="color:#6b7280;font-size:14px">تم استخدام <strong>{type_label}</strong> العائد لك على:</p>
    <table width="100%" style="background:#f9fafb;border-radius:10px;padding:18px;margin:16px 0">
      <tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">📄 الوثيقة:</span>
        <strong style="color:#111827;margin-right:8px">{doc_title}</strong></td></tr>
      <tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">🗂️ رقم الأرشيف:</span>
        <strong style="color:#1e40af;margin-right:8px">{archive_num}</strong></td></tr>
      <tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">🔖 الرقم التسلسلي:</span>
        <strong style="color:#065f46;margin-right:8px">{serial}</strong></td></tr>
      <tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">🧩 نوع الاستخدام:</span>
        <strong style="color:#374151;margin-right:8px">{type_label}</strong></td></tr>
      <tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">👤 استخدمه:</span>
        <strong style="color:#374151;margin-right:8px">{used_by}</strong></td></tr>
      <tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">🕒 وقت الاستخدام:</span>
        <strong style="color:#374151;margin-right:8px">{used_at_value}</strong></td></tr>
      {extra_rows}
    </table>
    <div style="background:#fef2f2;border-radius:8px;padding:14px;border-right:4px solid #ef4444">
      <p style="color:#991b1b;font-size:13px;margin:0">
        ⚠️ إذا لم تأذن بهذا الاستخدام تواصل مع المدير العام فوراً
      </p>
    </div>"""
    return send_email(to_email, f"⚠️ تم استخدام {type_label}ك — {doc_title}", _wrap(content, "إشعار التوقيع", "#b45309"))

# ─── 4. طلب توقيع ────────────────────────────
def send_sign_request(to_email, name, doc_title, archive_num, requester, message=""):
    msg_row = f"""<tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">💬 ملاحظة:</span>
      <span style="color:#374151;margin-right:8px">{message}</span></td></tr>""" if message else ""
    content = f"""
    <p style="color:#374151;font-size:15px">مرحباً <strong>{name}</strong>،</p>
    <p style="color:#6b7280;font-size:14px">الموظف <strong>{requester}</strong> أرسل لك مستنداً للتوقيع، يرجى توقيعه.</p>
    <table width="100%" style="background:#f9fafb;border-radius:10px;padding:18px;margin:16px 0">
      <tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">📄 الوثيقة:</span>
        <strong style="color:#111827;margin-right:8px">{doc_title}</strong></td></tr>
      <tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">🗂️ رقم الأرشيف:</span>
        <strong style="color:#1e40af;margin-right:8px">{archive_num}</strong></td></tr>
      {msg_row}
    </table>"""
    return send_email(to_email, f"🔔 {requester} أرسل لك مستنداً للتوقيع — {doc_title}", _wrap(content, "طلب توقيع"))

# ─── 5. اعتماد ✅ ─────────────────────────────
def send_approved(to_email, name, doc_title, archive_num):
    content = f"""
    <p style="color:#374151;font-size:15px">مرحباً <strong>{name}</strong>،</p>
    <div style="background:#ecfdf5;border-radius:12px;padding:24px;text-align:center;margin:20px 0">
      <div style="font-size:48px">✅</div>
      <div style="color:#065f46;font-size:18px;font-weight:700;margin-top:8px">تم اعتماد الوثيقة</div>
    </div>
    <table width="100%" style="background:#f9fafb;border-radius:10px;padding:18px">
      <tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">📄</span>
        <strong style="color:#111827;margin-right:8px">{doc_title}</strong></td></tr>
      <tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">🗂️ رقم الأرشيف:</span>
        <strong style="color:#1e40af;margin-right:8px">{archive_num}</strong></td></tr>
    </table>"""
    return send_email(to_email, f"✅ تم اعتماد الوثيقة — {doc_title}", _wrap(content, "اعتماد الوثيقة", "#065f46"))

# ─── 6. رفض ❌ ───────────────────────────────
def send_rejected(to_email, name, doc_title, archive_num, reason=""):
    reason_row = f"""<tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">📝 السبب:</span>
      <span style="color:#991b1b;margin-right:8px">{reason}</span></td></tr>""" if reason else ""
    content = f"""
    <p style="color:#374151;font-size:15px">مرحباً <strong>{name}</strong>،</p>
    <div style="background:#fef2f2;border-radius:12px;padding:24px;text-align:center;margin:20px 0">
      <div style="font-size:48px">❌</div>
      <div style="color:#991b1b;font-size:18px;font-weight:700;margin-top:8px">تم رفض الوثيقة</div>
    </div>
    <table width="100%" style="background:#f9fafb;border-radius:10px;padding:18px">
      <tr><td style="padding:5px 0"><span style="color:#6b7280;font-size:13px">📄</span>
        <strong style="color:#111827;margin-right:8px">{doc_title}</strong></td></tr>
      {reason_row}
    </table>"""
    return send_email(to_email, f"❌ تم رفض الوثيقة — {doc_title}", _wrap(content, "رفض الوثيقة", "#991b1b"))

def send_weekly_password(to_email, name, new_password):
    content = f"""
    <p style="color:#374151;font-size:15px">مرحباً <strong>{name}</strong>،</p>
    <p style="color:#6b7280;font-size:14px">
      تم تجديد الرقم السري الأسبوعي الخاص بك.
    </p>
    <div style="background:#eff6ff;border-radius:12px;padding:24px;text-align:center;margin:20px 0">
      <div style="color:#6b7280;font-size:13px;margin-bottom:8px">🔐 الرقم السري الجديد</div>
      <div style="letter-spacing:4px;font-size:32px;font-weight:800;color:#1e40af;font-family:monospace">{new_password}</div>
    </div>
    <div style="background:#fef2f2;border-radius:8px;padding:14px;border-right:4px solid #ef4444">
      <p style="color:#991b1b;font-size:13px;margin:0">
        ⚠️ هذا الرقم سري — لا تشاركه مع أحد
      </p>
    </div>"""
    return send_email(to_email, "🔁 تحديث الرقم السري الأسبوعي — SignMy", _wrap(content, "التحديث الأسبوعي", "#1e40af"))

# ─── 7. جهاز جديد غير مسجل ───────────────────
def send_new_device(to_email, name, device_name, ip, confirm_url, deny_url):
    content = f"""
    <p style="color:#374151;font-size:15px">مرحباً <strong>{name}</strong>،</p>
    <p style="color:#6b7280;font-size:14px">
      تم تسجيل دخول إلى حسابك من <strong>جهاز جديد غير مسجل</strong>:
    </p>

    <table width="100%" style="background:#fffbeb;border-radius:10px;
           padding:18px;margin:16px 0;border:2px solid #f59e0b">
      <tr><td style="padding:7px 0">
        <span style="color:#6b7280;font-size:13px">💻 الجهاز:</span>
        <strong style="color:#b45309;font-size:14px;margin-right:8px">{device_name}</strong>
      </td></tr>
      <tr><td style="padding:7px 0">
        <span style="color:#6b7280;font-size:13px">🌐 عنوان IP:</span>
        <strong style="color:#374151;font-size:14px;margin-right:8px">{ip}</strong>
      </td></tr>
      <tr><td style="padding:7px 0">
        <span style="color:#6b7280;font-size:13px">🕐 الوقت:</span>
        <strong style="color:#374151;font-size:14px;margin-right:8px">
          {datetime.now().strftime('%Y-%m-%d %H:%M')}
        </strong>
      </td></tr>
    </table>

    <p style="color:#111827;font-size:15px;font-weight:700;text-align:center">
      هل أنت من قام بتسجيل الدخول؟
    </p>

    <table width="100%" cellpadding="0" cellspacing="0" style="margin:20px 0">
      <tr>
        <td width="48%">
          <a href="{confirm_url}"
             style="display:block;background:#065f46;color:#fff;text-decoration:none;
                    padding:16px;border-radius:10px;text-align:center;
                    font-size:15px;font-weight:700">
            ✅ نعم، أنا<br>
            <span style="font-size:11px;opacity:.8">أضف هذا الجهاز للأجهزة الموثوقة</span>
          </a>
        </td>
        <td width="4%"></td>
        <td width="48%">
          <a href="{deny_url}"
             style="display:block;background:#991b1b;color:#fff;text-decoration:none;
                    padding:16px;border-radius:10px;text-align:center;
                    font-size:15px;font-weight:700">
            ❌ لا، ليس أنا<br>
            <span style="font-size:11px;opacity:.8">أوقف الحساب فوراً</span>
          </a>
        </td>
      </tr>
    </table>

    <div style="background:#fef2f2;border-radius:8px;padding:14px;border-right:4px solid #ef4444">
      <p style="color:#991b1b;font-size:13px;margin:0">
        ⚠️ إذا لم تكن أنت — اضغط "لا، ليس أنا" فوراً لحماية حسابك<br>
        🔒 لا تتجاهل هذا الإيميل
      </p>
    </div>"""
    return send_email(
        to_email,
        "⚠️ تسجيل دخول من جهاز جديد — SignMy",
        _wrap(content, "تنبيه أمني — جهاز جديد", "#b45309")
    )