#!/bin/bash
# Fix server script - run on the server

echo "=== Fixing Server ==="

# Pull latest code
cd /var/www/project
git pull origin main

# Kill existing gunicorn
pkill -f "gunicorn.*8000"
sleep 2

# Activate venv
source .venv/bin/activate

# Fix database - add missing columns
echo "Fixing database..."
python3 << 'PYEOF'
import sqlite3
conn = sqlite3.connect('najm.db')
c = conn.cursor()

# Fix users table
c.execute("PRAGMA table_info(users)")
if 'last_otp_sent_at' not in [r[1] for r in c.fetchall()]:
    c.execute("ALTER TABLE users ADD COLUMN last_otp_sent_at TEXT")
    print("Added last_otp_sent_at")

# Fix documents table
c.execute("PRAGMA table_info(documents)")
if 'original_file_sha256' not in [r[1] for r in c.fetchall()]:
    c.execute("ALTER TABLE documents ADD COLUMN original_file_sha256 TEXT")
    print("Added original_file_sha256")

conn.commit()
conn.close()
print("Database fixed!")
PYEOF

# Start gunicorn
echo "Starting Gunicorn..."
gunicorn -w 2 -b 127.0.0.1:8000 app:app --daemon --access-logfile logs/access.log --error-logfile logs/error.log

sleep 2

# Check if running
if pgrep -f "gunicorn.*8000" > /dev/null; then
    echo "✅ Server is running!"
    curl -s -o /dev/null -w "%{http_code}" http://127.0.0.1:8000/login
else
    echo "❌ Failed to start server"
fi

echo "=== Done ==="
