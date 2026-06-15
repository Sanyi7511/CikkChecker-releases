"""
CikkChecker v3 – pywebview + Rust backend
Modern gradient web UI with native window.
"""
import json, os, queue, subprocess, sys, tempfile, threading, time
from collections import Counter
from datetime import datetime
from tkinter import filedialog
import tkinter as tk

import requests
import webview

# Version is read from version.txt (written by GitHub Actions at build time)
def _read_version():
    try:
        vpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "version.txt")
        with open(vpath, encoding="utf-8") as f:
            return f.read().strip().lstrip("v")
    except:
        return "0.0.0"

APP_VERSION = _read_version()
UPDATE_REPO  = "Sanyi7511/CikkChecker-releases"
UPDATE_ASSET = "CikkCheckerSetup.exe"
BINARY_NAME  = "checker_core.exe" if sys.platform == "win32" else "checker_core"
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
BINARY_PATH  = os.path.join(BASE_DIR, BINARY_NAME)
CONFIG_PATH  = os.path.join(BASE_DIR, "config.json")
UI_PATH      = os.path.join(BASE_DIR, "ui", "index.html")

DEFAULT_CONFIG = {
    "theme": "dark", "language": "hu",
    "auto_update": True, "update_interval": "on_start",
    "sleep_seconds": 0.8, "save_every": 100,
    "sort_order": "abc", "show_price": True,
    "auto_rm_dup": False, "dont_ask_dup": False, "save_dup_txt": True,
    "last_url": "", "last_xlsx": "", "last_txt": "",
}

def load_config():
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, encoding="utf-8") as f:
                return {**DEFAULT_CONFIG, **json.load(f)}
    except: pass
    return dict(DEFAULT_CONFIG)

def save_config_file(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except: pass


def export_excel(csv_path, excel_path):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        wb = Workbook(); ws = wb.active; ws.title = "CikkChecker"
        fills = {
            "Van":          PatternFill("solid", fgColor="0A2016"),
            "Nincs":        PatternFill("solid", fgColor="200A0A"),
            "Külső raktár": PatternFill("solid", fgColor="201508"),
        }
        fonts = {
            "Van":          Font(color="23C55E", bold=True, size=10),
            "Nincs":        Font(color="F04747", bold=True, size=10),
            "Külső raktár": Font(color="F5A623", bold=True, size=10),
        }
        thin = Side(style="thin", color="1F2535")
        brd  = Border(left=thin, right=thin, top=thin, bottom=thin)
        ws.append(["Cikkszám","Elérhetőség","Ár"])
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="0D1018")
            cell.font = Font(color="9BA3B4", bold=True, size=10)
            cell.alignment = Alignment(horizontal="center"); cell.border = brd
        import csv as _csv
        if os.path.exists(csv_path):
            with open(csv_path, encoding="utf-8-sig", newline="") as f:
                for row in _csv.DictReader(f):
                    # Try both accented and plain column names
                    c = (row.get("Cikkszám") or row.get("Cikkszam") or "").strip()
                    a = (row.get("Elérhetőség") or row.get("Elerhetoseg") or "").strip()
                    p = (row.get("Ár") or row.get("Ar") or "").strip()
                    if not c: continue
                    ws.append([c, a, p]); r = ws.max_row
                    for col in range(1, 4):
                        cell = ws.cell(r, col); cell.border = brd
                        cell.alignment = Alignment(horizontal="left" if col==1 else "center")
                        cell.fill = fills.get(a, PatternFill("solid", fgColor="13161F"))
                    ws.cell(r,2).font = fonts.get(a, Font(color="9BA3B4", size=10))
                    ws.cell(r,1).font = Font(color="E8EBF0", size=10)
                    ws.cell(r,3).font = Font(color="9BA3B4", size=10)
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 16
        wb.save(excel_path)
    except Exception as e:
        print(f"Excel export error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# Python API — exposed to JavaScript via pywebview
# ══════════════════════════════════════════════════════════════════════════════
class Api:
    def __init__(self):
        self._window     = None   # set after window creation
        self._cfg        = load_config()
        self._proc       = None
        self._stop_req   = False
        self._unk_mode   = None   # None | "skip_all" | "stop_all"
        self._unk_event  = None
        self._unk_result = None
        self._xlsx_path  = ""
        self._csv_path   = ""

    def set_window(self, w):
        self._window = w

    # ── Config ────────────────────────────────────────────────────────────
    def get_config(self):
        cfg = dict(self._cfg)
        cfg["app_version"] = APP_VERSION
        cfg["scheduler"]   = self.get_scheduler()
        return cfg

    def save_config(self, cfg):
        self._cfg.update(cfg)
        save_config_file(self._cfg)
        return True

    # ── File dialogs (must run on main thread via tk) ─────────────────────
    def _tk_dialog(self, fn):
        result = [None]
        ev = threading.Event()
        def _run():
            root = tk.Tk()
            root.withdraw()
            try: root.attributes('-topmost', True)
            except: pass
            result[0] = fn(root)
            root.destroy(); ev.set()
        threading.Thread(target=_run, daemon=True).start()
        ev.wait(timeout=30)
        return result[0]

    def browse_file(self, filetype: str):
        """Open file dialog for txt/csv/xlsx."""
        type_map = {
            "txt":  [("Text fájl", "*.txt"), ("Minden", "*.*")],
            "csv":  [("CSV fájl", "*.csv"), ("Minden", "*.*")],
            "xlsx": [("Excel fájl", "*.xlsx *.xls"), ("Minden", "*.*")],
        }
        filetypes = type_map.get(filetype, [("Minden", "*.*")])
        path = self._tk_dialog(lambda r: filedialog.askopenfilename(
            parent=r,
            title=f"Cikkszám forrás megnyitása ({filetype.upper()})",
            filetypes=filetypes))
        if path:
            self._cfg["last_txt"] = path
            self._cfg["last_txt_type"] = filetype
            save_config_file(self._cfg)
        return path or ""

    def parse_file(self, path: str, filetype: str, col_index: int = 0):
        """Parse codes from txt/csv/xlsx. Returns {codes, preview} or None."""
        try:
            codes = []
            preview = []

            if filetype == "txt":
                with open(path, encoding="utf-8", errors="replace") as f:
                    codes = [l.strip() for l in f if l.strip()]

            elif filetype == "csv":
                import csv as _csv
                with open(path, encoding="utf-8-sig", errors="replace") as f:
                    sample = f.read(2048)
                # Try to detect delimiter, fallback to semicolon then comma
                try:
                    dialect = _csv.Sniffer().sniff(sample, delimiters=",;	|")
                except _csv.Error:
                    dialect = _csv.excel  # default comma
                with open(path, encoding="utf-8-sig", errors="replace", newline="") as f:
                    reader = _csv.reader(f, dialect)
                    rows = list(reader)
                # Skip header if first cell looks like a header (non-numeric text)
                start = 0
                if rows and not any(c.isdigit() for c in (rows[0][col_index] if col_index < len(rows[0]) else "")):
                    start = 1
                codes   = [row[col_index].strip() for row in rows[start:] if col_index < len(row) and row[col_index].strip()]
                preview = [row[col_index].strip() for row in rows[start:start+5] if col_index < len(row)]

            elif filetype in ("xlsx", "xls"):
                try:
                    from openpyxl import load_workbook
                    wb = load_workbook(path, read_only=True, data_only=True)
                    ws = wb.active
                    rows = list(ws.iter_rows(values_only=True))
                    wb.close()
                    # Skip header if first cell is text (not a code)
                    start = 0
                    if rows:
                        first = str(rows[0][col_index] or "").strip()
                        if first and not any(c.isdigit() for c in first):
                            start = 1
                    codes   = [str(row[col_index]).strip() for row in rows[start:]
                               if col_index < len(row) and row[col_index] is not None
                               and str(row[col_index]).strip()]
                    preview = [str(row[col_index]).strip() for row in rows[start:start+5]
                               if col_index < len(row) and row[col_index] is not None]
                except ImportError:
                    return {"error": "openpyxl nincs telepítve"}

            # Deduplicate while preserving order
            seen_set = set()
            unique = [c for c in codes if not (c in seen_set or seen_set.add(c))]

            return {"codes": unique, "preview": preview, "total": len(codes), "unique": len(unique)}

        except Exception as e:
            self._js(f"onLog({json.dumps(f'[ERR] Fájl parse hiba: {e}')})")
            return None

    def drop_file(self, filename: str, filetype: str):
        """pywebview can't access drag-drop file paths directly — return None,
        user should use the browse buttons instead."""
        return None

    def browse_txt(self):
        """Legacy compat — redirect to browse_file."""
        return self.browse_file("txt")

    def export_excel_dialog(self, data):
        """Called from JS with table row data — shows save dialog and exports Excel."""
        path = self._tk_dialog(lambda r: filedialog.asksaveasfilename(
            parent=r, title="Excel mentési hely",
            initialfile=f"CikkChecker_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            defaultextension=".xlsx", filetypes=[("Excel fájl", "*.xlsx")]))
        if not path:
            return None
        try:
            from openpyxl import Workbook
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            wb = Workbook(); ws = wb.active; ws.title = "CikkChecker"
            fills = {
                "Van":          PatternFill("solid", fgColor="0A2016"),
                "Nincs":        PatternFill("solid", fgColor="200A0A"),
                "Külső raktár": PatternFill("solid", fgColor="201508"),
            }
            fonts = {
                "Van":          Font(color="23C55E", bold=True, size=10),
                "Nincs":        Font(color="F04747", bold=True, size=10),
                "Külső raktár": Font(color="F5A623", bold=True, size=10),
            }
            thin = Side(style="thin", color="2A3147")
            brd  = Border(left=thin, right=thin, top=thin, bottom=thin)
            ws.append(["Cikkszám", "Elérhetőség", "Ár"])
            for cell in ws[1]:
                cell.fill = PatternFill("solid", fgColor="0D1018")
                cell.font = Font(color="9BA3B4", bold=True, size=10)
                cell.alignment = Alignment(horizontal="center"); cell.border = brd
            for row in (data or []):
                c = str(row.get("code","")).strip()
                a = str(row.get("avail","")).strip()
                p = str(row.get("price","")).strip()
                if not c: continue
                ws.append([c, a, p]); r = ws.max_row
                for col in range(1,4):
                    cell = ws.cell(r, col); cell.border = brd
                    cell.alignment = Alignment(horizontal="left" if col==1 else "center")
                    cell.fill = fills.get(a, PatternFill("solid", fgColor="13161F"))
                ws.cell(r,2).font = fonts.get(a, Font(color="9BA3B4", size=10))
                ws.cell(r,1).font = Font(color="E8EBF0", size=10)
                ws.cell(r,3).font = Font(color="9BA3B4", size=10)
            ws.column_dimensions["A"].width = 28
            ws.column_dimensions["B"].width = 18
            ws.column_dimensions["C"].width = 16
            wb.save(path)
            return path
        except Exception as e:
            self._js(f"onLog({json.dumps(f'[ERR] Excel export hiba: {e}')})")
            return None

    # ── Login detection ───────────────────────────────────────────────────
    def detect_login(self, url):
        try:
            if not url.startswith("http"): url = "https://" + url
            resp = requests.get(f"{url}/product-search/test/0?1", timeout=8, allow_redirects=True)
            html = resp.text.lower()
            return ("/login" in resp.url or
                    any(k in html for k in ["bejelentkezés","belépés","felhasználónév","user-name","password"]))
        except:
            return False

    # ── Process control ───────────────────────────────────────────────────
    def start_process(self, config):
        if self._proc and self._proc.poll() is None:
            self._js("onLog('[!!] A folyamat már fut.')")
            return False

        url        = config.get("url","").strip()
        user       = config.get("username","").strip()
        pw         = config.get("password","").strip()
        req_login  = config.get("req_login", False)
        txt_path   = config.get("txt_path","").strip()
        sleep_s    = float(config.get("sleep_s", 0.8))
        save_every = int(config.get("save_every", 100))
        sort_order = config.get("sort_order","abc")
        auto_rm    = config.get("auto_rm", False)
        save_dup   = config.get("save_dup", True)
        manual     = config.get("manual","").strip()
        skip_dup   = config.get("skip_dup", False)

        if not url:
            self._js("onLog('[!!] Add meg az oldal URL-jét.')")
            return False
        if req_login and (not user or not pw):
            self._js("onLog('[!!] Felhasználónév és jelszó szükséges.')")
            return False

        # Auto-determine output folder:
        # 1. Same folder as input file (if given)
        # 2. Desktop
        # 3. Current directory
        if txt_path and os.path.exists(txt_path):
            out_folder = os.path.dirname(os.path.abspath(txt_path))
        else:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            out_folder = desktop if os.path.exists(desktop) else os.path.abspath(".")

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        xlsx_path = os.path.join(out_folder, f"CikkChecker_{ts}.xlsx")

        # Collect codes — preloaded takes priority over raw txt_path
        preloaded = config.get("preloaded", [])
        codes = []
        if preloaded:
            codes = [str(c).strip() for c in preloaded if str(c).strip()]
        elif txt_path and os.path.exists(txt_path):
            # Fallback: read raw txt
            with open(txt_path, encoding="utf-8", errors="replace") as f:
                codes += [l.strip() for l in f if l.strip()]
        if manual:
            codes += [l.strip() for l in manual.splitlines() if l.strip()]
        if not codes:
            self._js("onLog('[!!] Adj meg cikkszámokat.')")
            return False

        xlsx_folder = os.path.dirname(xlsx_path) or "."
        os.makedirs(xlsx_folder, exist_ok=True)
        csv_path = os.path.join(xlsx_folder, "cikkchecker_results.csv")

        # Duplicates
        counter   = Counter(codes)
        dup_items = {c:n for c,n in counter.items() if n>1}
        unique    = list(dict.fromkeys(codes))

        if dup_items and not skip_dup:
            dup_rows = sum(n-1 for n in counter.values() if n>1)
            if auto_rm:
                codes = unique
                self._js(f"onLog('[dup] {dup_rows} duplikált sor eltávolítva.')")
            if save_dup:
                dup_path = os.path.join(xlsx_folder, "duplicates.txt")
                with open(dup_path,"w",encoding="utf-8") as f:
                    for c,n in dup_items.items(): f.write(f"{c};{n}\n")

        # Total
        self._js(f"document.getElementById('s-total').textContent = '{len(set(codes))}'")

        # Write codes to temp file
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".txt",
                                          encoding="utf-8", delete=False)
        tmp.write("\n".join(codes)); tmp.close()

        self._cfg["last_url"] = url
        self._cfg["last_txt"] = txt_path
        save_config_file(self._cfg)

        if not os.path.exists(BINARY_PATH):
            self._js(f"onLog('[!!] checker_core nem található: {BINARY_PATH}')")
            return False

        cmd = [BINARY_PATH,
               "--base-url", url,
               "--user", user, "--password", pw,
               "--codes-file", tmp.name, "--csv-output", csv_path,
               "--sleep-seconds", str(sleep_s), "--save-every", str(save_every),
               "--sort-order", sort_order]
        if req_login: cmd.append("--requires-login")

        self._stop_req   = False
        self._unk_mode   = None
        self._xlsx_path  = xlsx_path
        self._csv_path   = csv_path

        self._js(f"onLog('[--] Indítás — {len(codes)} cikkszám')")
        self._js(f"onLog('[--] Kimenet: {xlsx_path}')")
        self._js("onStatus('Indítás...')")

        self._proc = subprocess.Popen(
            cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.PIPE, text=True, bufsize=1)

        threading.Thread(target=self._read_proc, args=(tmp.name,), daemon=True).start()
        return True

    def stop_process(self):
        self._stop_req = True
        if self._proc:
            try: self._proc.stdin.write("stop\n"); self._proc.stdin.flush()
            except: pass
        self._js("onLog('[--] Leállítás kérve...')")
        return True

    def restart_process(self):
        self.stop_process()
        self._stop_req  = False
        self._unk_mode  = None
        self._js("onLog('[--] Újraindítás — folyamat leállítva, újraindítható.')")
        self._js("onStatus('Várakozás')")
        return True

    # ── Process reader ────────────────────────────────────────────────────
    def _read_proc(self, tmp_path):
        try:
            for line in self._proc.stdout:
                line = line.strip()
                if not line: continue
                try: msg = json.loads(line)
                except: self._js(f"onLog({json.dumps(line)})"); continue

                kind = msg.get("kind","")

                if kind == "log":
                    ts = datetime.now().strftime("%H:%M:%S")
                    text = msg.get("msg", "")
                    self._js(f"onLog({json.dumps(f'[{ts}] {text}')})")

                elif kind == "status":
                    self._js(f"onStatus({json.dumps(msg['msg'])})")

                elif kind == "progress":
                    self._js(f"onProgress({msg['current']},{msg['total']})")

                elif kind == "login_detected":
                    req = str(msg.get("required",False)).lower()
                    self._js(f"onLoginDetected({req})")

                elif kind == "result":
                    code  = json.dumps(msg.get("cikkszam",""))
                    avail = json.dumps(msg.get("elerhetoseg",""))
                    price = json.dumps(msg.get("ar",""))
                    self._js(f"onResult({code},{avail},{price})")

                elif kind == "unknown":
                    code = msg["cikkszam"]
                    action = self._handle_unknown(code)
                    try: self._proc.stdin.write(action+"\n"); self._proc.stdin.flush()
                    except: pass

                elif kind == "error":
                    self._js(f"onLog({json.dumps('[ERR] ' + msg['msg'])})")

                elif kind == "done":
                    self._js("onLog('[--] Kész — Excel exportálás...')")
                    export_excel(self._csv_path, self._xlsx_path)
                    self._js(f"onLog({json.dumps(f'[--] Excel mentve: {self._xlsx_path}')})")
                    self._js("onDone()")

        except Exception as e:
            self._js(f"onLog({json.dumps(f'[ERR] Olvasási hiba: {e}')})")
        finally:
            try: os.unlink(tmp_path)
            except: pass

    def _handle_unknown(self, code):
        if self._unk_mode == "skip_all": return "skip"
        if self._unk_mode == "stop_all": return "stop"

        self._unk_event  = threading.Event()
        self._unk_result = None
        self._js(f"onUnknown({json.dumps(code)}).then(r => {{ window._unkResult = r; }})")

        # Poll for JS result
        for _ in range(300):  # 30 seconds max
            time.sleep(0.1)
            if self._unk_result is not None:
                break

        result = self._unk_result or {"action":"stop","applyAll":False}
        if result.get("applyAll"):
            if result["action"] == "skip": self._unk_mode = "skip_all"
            else:                          self._unk_mode = "stop_all"
        return result["action"]

    def set_unknown_result(self, action, apply_all):
        """Called from JS after user resolves unknown dialog."""
        self._unk_result = {"action": action, "applyAll": apply_all}
        return True

    # ── Update ────────────────────────────────────────────────────────────
    def check_update(self):
        threading.Thread(target=self._do_check_update, daemon=True).start()
        return True

    def _do_check_update(self):
        try:
            r = requests.get(
                f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest",
                timeout=10)
            r.raise_for_status(); data = r.json()
        except:
            self._js("onLog('[--] Frissítés ellenőrzése sikertelen.')")
            return

        latest = data.get("tag_name","").strip().lstrip("v")
        notes  = data.get("body","").strip()

        def _v(s):
            s = s.strip().lstrip("v")
            parts = s.split(".")
            return tuple(int(x) if x.isdigit() else 0 for x in parts)

        current = _v(APP_VERSION)
        remote  = _v(latest)

        self._js(f"onLog('[--] Jelenlegi verzió: {APP_VERSION} | GitHub: {latest}')")

        if not latest or remote <= current:
            self._js("onLog('[--] Nincs újabb verzió.')")
            return

        # Find installer asset — name includes version e.g. CikkCheckerSetup_3.1.0.exe
        asset_url = next((a["browser_download_url"] for a in data.get("assets",[])
                          if a.get("name","").startswith("CikkCheckerSetup")), None)
        if not asset_url:
            self._js("onLog('[--] Új verzió van, de a telepítő nem található.')")
            return

        self._js(f"onLog('[!!] Új verzió elérhető: v{latest}')")
        # Show banner — wait for user to approve before downloading
        self._js(f"onUpdateAvailable({json.dumps(latest)},{json.dumps(asset_url)})")

    def do_update(self, asset_url: str, latest: str):
        """Called from JS when user clicks the update button."""
        threading.Thread(target=self._dl_update,
                         args=(asset_url, latest), daemon=True).start()
        return True

    def _dl_update(self, url, latest):
        self._js(f"onLog('[--] Frissítő letöltése: v{latest}')")
        self._js("onStatus('Frissítés letöltése...')")
        self._js(f"onUpdateDownloading({json.dumps(latest)})")
        try:
            tmp_dir  = tempfile.gettempdir()
            # Versioned filename to avoid conflicts with running instance
            versioned_name = f"CikkCheckerSetup_{latest}.exe"
            tmp = os.path.join(tmp_dir, versioned_name)

            with requests.get(url, stream=True, timeout=120) as r:
                r.raise_for_status()
                total = int(r.headers.get("content-length", 0))
                downloaded = 0
                with open(tmp, "wb") as f:
                    for chunk in r.iter_content(8192):
                        if chunk:
                            f.write(chunk)
                            downloaded += len(chunk)
                            if total:
                                pct = int(downloaded / total * 100)
                                self._js(f"onUpdateProgress({pct})")

            self._js("onLog('[--] Letöltés kész — telepítő indul...')")
            self._js("onUpdateReady()")
            time.sleep(1.5)

            # Use a .bat launcher with delay so the app can fully close
            # before the installer touches any DLL files
            bat_path = os.path.join(tmp_dir, "cikkchecker_update.bat")
            # Find the installed CikkChecker.exe path to restart after update
            app_exe = os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__)
            # If running as PyInstaller bundle, exe is sys.executable
            # If running as script, find the CikkChecker.exe in same dir
            if not getattr(sys, "frozen", False):
                app_exe = os.path.join(os.path.dirname(os.path.abspath(__file__)), "CikkChecker.exe")

            bat = (
                "@echo off\n"
                ":: Varunk amig az app teljesen bezarul\n"
                "timeout /t 4 /nobreak > nul\n"
                ":: Csendes telepites\n"
                f'start "" /wait "{tmp}" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /CLOSEAPPLICATIONS\n'
                ":: Varunk hogy a telepito befejezze\n"
                "timeout /t 3 /nobreak > nul\n"
                ":: Ujrainditjuk az appot\n"
                f'if exist "{app_exe}" start "" "{app_exe}"\n'
                ":: Self-delete\n"
                "del \"%~f0\"\n"
            )
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(bat)

            import subprocess as sp
            sp.Popen(
                ["cmd", "/c", bat_path],
                creationflags=0x08000000,  # CREATE_NO_WINDOW
                close_fds=True
            )
            time.sleep(0.5)
            self._window.destroy()

        except Exception as e:
            self._js(f"onLog({json.dumps(f'[ERR] Frissítési hiba: {e}')})")
            self._js("onUpdateFailed()")

    # ── History ──────────────────────────────────────────────────────────
    def _history_path(self):
        hist_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local",
                                "CikkChecker") if sys.platform=="win32" else                    os.path.join(os.path.expanduser("~"), ".cikkchecker")
        os.makedirs(hist_dir, exist_ok=True)
        return os.path.join(hist_dir, "history.json")

    def get_history(self):
        try:
            p = self._history_path()
            if os.path.exists(p):
                with open(p, encoding="utf-8") as f:
                    return json.load(f)
        except: pass
        return []

    def save_history(self, entry):
        try:
            history = self.get_history()
            history.append(entry)
            history = history[-100:]  # keep last 100 runs
            with open(self._history_path(), "w", encoding="utf-8") as f:
                json.dump(history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self._js(f"onLog({json.dumps(f'[WARN] Előzmény mentés hiba: {e}')})")
        return True

    # ── Windows notification ──────────────────────────────────────────────
    def notify(self, message: str, subtitle: str = ""):
        if sys.platform != "win32":
            return True
        try:
            # Use PowerShell toast notification
            ps_script = f"""
Add-Type -AssemblyName System.Windows.Forms
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon = [System.Drawing.SystemIcons]::Information
$notify.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
$notify.BalloonTipText = '{message.replace("'", "")}'
$notify.BalloonTipTitle = 'CikkChecker'
$notify.Visible = $True
$notify.ShowBalloonTip(5000)
Start-Sleep -Milliseconds 5500
$notify.Dispose()
"""
            subprocess.Popen(
                ["powershell", "-NonInteractive", "-WindowStyle", "Hidden",
                 "-Command", ps_script],
                creationflags=0x08000000,
                close_fds=True
            )
        except Exception as e:
            pass  # Notification is optional, never crash for it
        return True

    # ── Scheduler config ──────────────────────────────────────────────────
    def save_scheduler(self, config):
        self._cfg["scheduler"] = config
        save_config_file(self._cfg)
        return True

    def get_scheduler(self):
        return self._cfg.get("scheduler", {"enabled": False, "time": "08:00", "days": [1,2,3,4,5]})

    # ── JS helper ─────────────────────────────────────────────────────────
    def _js(self, code):
        if self._window:
            try: self._window.evaluate_js(code)
            except: pass


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════
def main():
    api = Api()

    # Fix unknown dialog: JS calls back into Python
    original_handle = api._handle_unknown
    def patched_handle(code):
        if api._unk_mode == "skip_all": return "skip"
        if api._unk_mode == "stop_all": return "stop"

        api._unk_result = None
        api._js(f"""
            onUnknown({json.dumps(code)}).then(function(r) {{
                pywebview.api.set_unknown_result(r.action, r.applyAll);
            }});
        """)
        for _ in range(300):
            time.sleep(0.1)
            if api._unk_result is not None: break
        result = api._unk_result or {"action":"stop","applyAll":False}
        if result.get("applyAll"):
            if result["action"]=="skip": api._unk_mode="skip_all"
            else:                        api._unk_mode="stop_all"
        return result["action"]

    api._handle_unknown = patched_handle

    window = webview.create_window(
        title="CikkChecker",
        url=UI_PATH,
        js_api=api,
        width=1440, height=900,
        min_size=(1100, 700),
        background_color="#0D0821",
    )
    api.set_window(window)

    webview.start(debug=False)


if __name__ == "__main__":
    main()
