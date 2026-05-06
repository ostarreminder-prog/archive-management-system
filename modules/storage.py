"""
storage.py — إدارة الملفات عبر SFTP (paramiko)
يدعم: رفع، تنزيل، حذف، إعادة تسمية، قائمة الملفات
"""
import os
import stat

try:
    import paramiko
    PARAMIKO_AVAILABLE = True
except ImportError:
    PARAMIKO_AVAILABLE = False
    print("[STORAGE] ⚠️ paramiko غير مثبت — نفّذ: pip install paramiko")


class StorageManager:
    def __init__(self):
        self.host       = os.getenv("STORAGE_HOST",       "storage4000.is.cc")
        self.port       = int(os.getenv("STORAGE_PORT",   "22") or 22)
        self.username   = os.getenv("STORAGE_USERNAME",   "st72796")
        self.password   = os.getenv("STORAGE_PASSWORD",   "")
        self.remote_dir = os.getenv("STORAGE_REMOTE_DIR", "/archives").rstrip("/")

    # ── اتصال SFTP ───────────────────────────────────
    def _connect(self):
        if not PARAMIKO_AVAILABLE:
            raise RuntimeError("paramiko غير مثبت — نفّذ: pip install paramiko")

        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(
            hostname=self.host,
            port=self.port,
            username=self.username,
            password=self.password,
            timeout=30,
            banner_timeout=30,
            look_for_keys=False,
            allow_agent=False,
        )
        sftp = client.open_sftp()
        return client, sftp

    # ── مسار كامل ──────────────────────────────────
    def _full_path(self, sub="", name=""):
        parts = [self.remote_dir]
        if sub:  parts.append(sub.strip("/"))
        if name: parts.append(name)
        return "/".join(parts)

    # ── إنشاء مجلدات بشكل متكرر ───────────────────
    def _makedirs(self, sftp, remote_path):
        path = remote_path.rstrip("/")
        parts = [p for p in path.split("/") if p]
        current = ""
        for part in parts:
            current += f"/{part}"
            try:
                sftp.stat(current)
            except IOError:
                try:
                    sftp.mkdir(current)
                except Exception:
                    pass

    # ── رفع ملف ────────────────────────────────────
    def upload(self, local_path: str, remote_name: str, sub_path: str = "") -> bool:
        try:
            client, sftp = self._connect()
            remote_dir = self._full_path(sub_path)
            self._makedirs(sftp, remote_dir)
            remote_file = f"{remote_dir}/{remote_name}"
            sftp.put(local_path, remote_file)
            sftp.close()
            client.close()
            print(f"[STORAGE] ✅ رُفع: {remote_file}")
            return True
        except Exception as e:
            print(f"[STORAGE] ❌ upload: {e}")
            return False

    # ── رفع بيانات مباشرة (بدون ملف محلي) ─────────────
    def upload_bytes(self, blob: bytes, remote_name: str, sub_path: str = "") -> bool:
        if blob is None:
            return False
        client = None
        sftp = None
        try:
            client, sftp = self._connect()
            remote_dir = self._full_path(sub_path)
            self._makedirs(sftp, remote_dir)
            remote_file = f"{remote_dir}/{remote_name}"
            with sftp.open(remote_file, 'wb') as handle:
                handle.write(blob)
            print(f"[STORAGE] ✅ رُفع مباشر: {remote_file}")
            return True
        except Exception as e:
            print(f"[STORAGE] ❌ upload_bytes: {e}")
            return False
        finally:
            try:
                if sftp:
                    sftp.close()
            except Exception:
                pass
            try:
                if client:
                    client.close()
            except Exception:
                pass

    # ── تنزيل ملف ──────────────────────────────────
    def download(self, remote_path: str, local_path: str) -> bool:
        try:
            client, sftp = self._connect()
            full = self._full_path(remote_path) if not remote_path.startswith("/") else remote_path
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            sftp.get(full, local_path)
            sftp.close()
            client.close()
            return True
        except Exception as e:
            print(f"[STORAGE] ❌ download: {e}")
            return False

    # ── قائمة المجلد ───────────────────────────────
    def list_dir(self, sub_path: str = "") -> dict:
        client = None
        sftp = None
        try:
            client, sftp = self._connect()
            target = self._full_path(sub_path)
            try:
                items = sftp.listdir_attr(target)
            except Exception as e:
                message = str(e).lower()
                if "no such file" in message or "not found" in message:
                    self._makedirs(sftp, target)
                    items = []
                else:
                    raise

            result = {"folders": [], "files": []}
            for item in items:
                is_dir = stat.S_ISDIR(item.st_mode)
                if is_dir:
                    result["folders"].append({
                        "name":  item.filename,
                        "date":  "",
                        "count": 0,
                    })
                else:
                    size_kb = (item.st_size or 0) / 1024
                    result["files"].append({
                        "name": item.filename,
                        "size": f"{size_kb:.0f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB",
                        "date": str(item.st_mtime or ""),
                        "ts":   item.st_mtime or 0,
                    })

            return result
        except Exception as e:
            print(f"[STORAGE] ❌ list_dir: {e}")
            return {"folders": [], "files": []}
        finally:
            try:
                if sftp:
                    sftp.close()
            except Exception:
                pass
            try:
                if client:
                    client.close()
            except Exception:
                pass

    # ── إنشاء مجلد ─────────────────────────────────
    def create_folder(self, folder_name: str, sub_path: str = "") -> bool:
        try:
            client, sftp = self._connect()
            target = self._full_path(sub_path, folder_name)
            self._makedirs(sftp, target)
            sftp.close()
            client.close()
            return True
        except Exception as e:
            print(f"[STORAGE] ❌ create_folder: {e}")
            return False

    # ── حذف ملف أو مجلد ────────────────────────────
    def delete(self, remote_path: str) -> bool:
        try:
            client, sftp = self._connect()
            full = self._full_path(remote_path) if not remote_path.startswith("/") else remote_path
            try:
                sftp.remove(full)          # ملف
            except Exception:
                sftp.rmdir(full)           # مجلد فارغ
            sftp.close()
            client.close()
            return True
        except Exception as e:
            print(f"[STORAGE] ❌ delete: {e}")
            return False

    # ── إعادة تسمية ────────────────────────────────
    def rename(self, old_path: str, new_name: str) -> bool:
        try:
            client, sftp = self._connect()
            old_full = self._full_path(old_path) if not old_path.startswith("/") else old_path
            base     = old_full.rsplit("/", 1)[0]
            new_full = f"{base}/{new_name}"
            sftp.rename(old_full, new_full)
            sftp.close()
            client.close()
            return True
        except Exception as e:
            print(f"[STORAGE] ❌ rename: {e}")
            return False

    # ── معلومات التخزين ─────────────────────────────
    def get_storage_info(self) -> dict:
        client = None
        sftp = None
        default_info = {"used": "N/A", "total": "N/A", "free": "N/A", "percent": 0}
        try:
            client, sftp = self._connect()

            if not hasattr(sftp, "statvfs"):
                return default_info

            st = sftp.statvfs(self.remote_dir)
            total  = st.f_blocks * st.f_frsize
            free   = st.f_bavail * st.f_frsize
            used   = total - free
            pct    = round(used / total * 100) if total else 0

            def fmt(b):
                gb = b / 1024**3
                return f"{gb:.1f} GB" if gb >= 1 else f"{b/1024**2:.0f} MB"
            return {"used": fmt(used), "total": fmt(total), "free": fmt(free), "percent": pct}
        except Exception as e:
            message = str(e).lower()
            if "statvfs" in message or "unsupported" in message or "not implemented" in message:
                return default_info
            print(f"[STORAGE] ❌ storage_info: {e}")
            return default_info
        finally:
            try:
                if sftp:
                    sftp.close()
            except Exception:
                pass
            try:
                if client:
                    client.close()
            except Exception:
                pass

    # ── اختبار الاتصال ──────────────────────────────
    def test_connection(self) -> bool:
        try:
            client, sftp = self._connect()
            sftp.stat(self.remote_dir)
            sftp.close()
            client.close()
            print("[STORAGE] ✅ الاتصال ناجح")
            return True
        except Exception as e:
            print(f"[STORAGE] ❌ test: {e}")
            return False


# singleton
storage = StorageManager()
