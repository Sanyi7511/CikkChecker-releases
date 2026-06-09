"""
CikkChecker – Modern Dark Dashboard UI
Python frontend that drives the Rust checker_core binary.
"""

import json
import os
import queue
import subprocess
import sys
import tempfile
import threading
from collections import Counter
from datetime import datetime
from tkinter import filedialog
import tkinter.messagebox as messagebox

import customtkinter as ctk
import requests
from PIL import Image, ImageDraw

# ── Constants ─────────────────────────────────────────────────────────────────
APP_VERSION   = "2.0.0"
UPDATE_REPO   = "Sanyi7511/CikkChecker-releases"
UPDATE_ASSET  = "CikkCheckerSetup.exe"
BINARY_NAME   = "checker_core.exe" if sys.platform == "win32" else "checker_core"
BINARY_PATH   = os.path.join(os.path.dirname(os.path.abspath(__file__)), BINARY_NAME)

# ── Theme tokens ──────────────────────────────────────────────────────────────
COLORS = {
    "bg":          "#0F1117",
    "panel":       "#161B27",
    "card":        "#1E2433",
    "border":      "#2A3147",
    "accent":      "#4F8EF7",
    "accent_dim":  "#2B4E8C",
    "green":       "#22C55E",
    "green_dim":   "#166534",
    "red":         "#EF4444",
    "red_dim":     "#7F1D1D",
    "amber":       "#F59E0B",
    "amber_dim":   "#78350F",
    "text":        "#E2E8F0",
    "text_dim":    "#64748B",
    "text_mid":    "#94A3B8",
}

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


# ── Utility: Excel export (openpyxl) ─────────────────────────────────────────
def export_to_excel(csv_path: str, excel_path: str):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment

        wb = Workbook()
        ws = wb.active
        ws.title = "Lista"

        hdr_fill = PatternFill(start_color="1E2433", end_color="1E2433", fill_type="solid")
        hdr_font = Font(bold=True, color="E2E8F0", size=11)
        fills = {
            "Van":          PatternFill(start_color="166534", end_color="166534", fill_type="solid"),
            "Nincs":        PatternFill(start_color="7F1D1D", end_color="7F1D1D", fill_type="solid"),
            "Külső raktár": PatternFill(start_color="78350F", end_color="78350F", fill_type="solid"),
        }
        fonts = {
            "Van":          Font(color="DCFCE7", bold=True),
            "Nincs":        Font(color="FEE2E2", bold=True),
            "Külső raktár": Font(color="FEF3C7", bold=True),
        }

        ws.append(["Cikkszám", "Elérhetőség"])
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center")

        import csv as _csv
        if os.path.exists(csv_path):
            with open(csv_path, encoding="utf-8-sig", newline="") as f:
                reader = _csv.DictReader(f)
                for row in reader:
                    c = (row.get("Cikkszám") or "").strip()
                    a = (row.get("Elérhetőség") or "").strip()
                    if c:
                        ws.append([c, a])
                        r = ws.max_row
                        if a in fills:
                            ws[f"B{r}"].fill = fills[a]
                            ws[f"B{r}"].font = fonts[a]
                        ws[f"B{r}"].alignment = Alignment(horizontal="center")

        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 20
        wb.save(excel_path)
    except Exception as e:
        print(f"[Excel export hiba] {e}")


# ── Duplicate dialog ──────────────────────────────────────────────────────────
class DuplicateDialog(ctk.CTkToplevel):
    def __init__(self, parent, duplicate_items, total, unique_count, duplicate_rows):
        super().__init__(parent)
        self.result_remove = False
        self.result_save   = False
        self.result_ok     = False

        self.title("Duplikátumok")
        self.geometry("660x500")
        self.minsize(580, 440)
        self.configure(fg_color=COLORS["bg"])
        self.transient(parent)
        self.grab_set()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header
        ctk.CTkLabel(self, text="⚠  Duplikátumok találva",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=COLORS["amber"]).grid(row=0, column=0, padx=24, pady=(22,10), sticky="w")

        # Summary card
        card = ctk.CTkFrame(self, fg_color=COLORS["card"], corner_radius=10)
        card.grid(row=1, column=0, padx=20, pady=(0,10), sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)

        summary = (f"Összes sor: {total}   •   Egyedi: {unique_count}   •   Duplikált: {duplicate_rows}\n\n"
                   "Cikkszámok több példányban:")
        ctk.CTkLabel(card, text=summary, justify="left",
                     text_color=COLORS["text_mid"]).grid(row=0, column=0, padx=14, pady=(14,8), sticky="w")

        tb = ctk.CTkTextbox(card, font=ctk.CTkFont(family="Consolas", size=12),
                            fg_color=COLORS["panel"], text_color=COLORS["text"])
        tb.grid(row=1, column=0, padx=14, pady=(0,14), sticky="nsew")
        for code, count in duplicate_items.items():
            tb.insert("end", f"  {code}   ×{count}\n")
        tb.configure(state="disabled")

        # Options
        opts = ctk.CTkFrame(self, fg_color=COLORS["card"], corner_radius=10)
        opts.grid(row=2, column=0, padx=20, pady=8, sticky="ew")
        opts.grid_columnconfigure(0, weight=1)

        self._rm_var   = ctk.BooleanVar(value=True)
        self._save_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(opts, text="Minden cikkszámból csak 1 maradjon",
                        variable=self._rm_var,
                        fg_color=COLORS["accent"]).grid(row=0, column=0, padx=14, pady=(14,6), sticky="w")
        ctk.CTkCheckBox(opts, text="Duplikátumok mentése külön TXT fájlba",
                        variable=self._save_var,
                        fg_color=COLORS["accent"]).grid(row=1, column=0, padx=14, pady=(6,14), sticky="w")

        # Buttons
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.grid(row=3, column=0, padx=20, pady=(0,20), sticky="ew")
        btn_row.grid_columnconfigure((0,1), weight=1)

        ctk.CTkButton(btn_row, text="Folytatás", height=42, corner_radius=8,
                      fg_color=COLORS["accent"], hover_color=COLORS["accent_dim"],
                      command=self._ok).grid(row=0, column=0, padx=8, pady=0, sticky="ew")
        ctk.CTkButton(btn_row, text="Mégse", height=42, corner_radius=8,
                      fg_color=COLORS["card"], hover_color=COLORS["border"],
                      command=self._cancel).grid(row=0, column=1, padx=8, pady=0, sticky="ew")

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_visibility(); self.focus()

    def _ok(self):
        self.result_remove = self._rm_var.get()
        self.result_save   = self._save_var.get()
        self.result_ok = True
        self.destroy()

    def _cancel(self):
        self.destroy()


# ── Unknown result dialog ─────────────────────────────────────────────────────
class UnknownDialog(ctk.CTkToplevel):
    def __init__(self, parent, cikkszam):
        super().__init__(parent)
        self.action = "stop"
        self.apply_all = False

        self.title("Ismeretlen találat")
        self.geometry("420x210")
        self.minsize(380, 200)
        self.configure(fg_color=COLORS["bg"])
        self.transient(parent)
        self.grab_set()
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="Nem értelmezhető találat",
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=COLORS["amber"]).grid(row=0, column=0, padx=20, pady=(20,8), sticky="w")
        ctk.CTkLabel(self, text=f"Cikkszám: {cikkszam}\nMit szeretnél tenni?",
                     justify="left", text_color=COLORS["text_mid"]).grid(row=1, column=0, padx=20, pady=(0,8), sticky="w")

        self._all_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(self, text="Alkalmazza a további ismeretlen találatokra is",
                        variable=self._all_var, fg_color=COLORS["accent"]
                        ).grid(row=2, column=0, padx=20, pady=(0,10), sticky="w")

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.grid(row=3, column=0, padx=20, pady=(0,20), sticky="ew")
        btns.grid_columnconfigure((0,1), weight=1)

        ctk.CTkButton(btns, text="Kihagyás", height=38, corner_radius=8,
                      fg_color=COLORS["card"], hover_color=COLORS["border"],
                      command=self._skip).grid(row=0, column=0, padx=6, sticky="ew")
        ctk.CTkButton(btns, text="Leállítás", height=38, corner_radius=8,
                      fg_color=COLORS["red_dim"], hover_color=COLORS["red"],
                      command=self._stop).grid(row=0, column=1, padx=6, sticky="ew")

        self.protocol("WM_DELETE_WINDOW", self._stop)
        self.wait_visibility(); self.focus()

    def _skip(self):
        self.action = "skip"; self.apply_all = self._all_var.get(); self.destroy()

    def _stop(self):
        self.action = "stop"; self.apply_all = self._all_var.get(); self.destroy()


# ── Update dialog ─────────────────────────────────────────────────────────────
class UpdateDialog(ctk.CTkToplevel):
    def __init__(self, parent, current, latest, notes):
        super().__init__(parent)
        self.result = "later"
        self.title("Frissítés")
        self.geometry("620x400")
        self.configure(fg_color=COLORS["bg"])
        self.transient(parent); self.grab_set()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self, text="🚀  Új verzió elérhető",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=COLORS["green"]).grid(row=0, column=0, padx=24, pady=(22,10), sticky="w")

        card = ctk.CTkFrame(self, fg_color=COLORS["card"], corner_radius=10)
        card.grid(row=1, column=0, padx=20, pady=(0,10), sticky="nsew")
        card.grid_columnconfigure(0, weight=1); card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(card, text=f"Jelenlegi: {current}   →   Elérhető: {latest}",
                     text_color=COLORS["text_mid"]).grid(row=0, column=0, padx=14, pady=(14,8), sticky="w")

        tb = ctk.CTkTextbox(card, fg_color=COLORS["panel"], text_color=COLORS["text"])
        tb.grid(row=1, column=0, padx=14, pady=(0,14), sticky="nsew")
        tb.insert("1.0", notes.strip() or "Nincs kiadási megjegyzés.")
        tb.configure(state="disabled")

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.grid(row=2, column=0, padx=20, pady=(0,20), sticky="ew")
        btn_row.grid_columnconfigure((0,1), weight=1)

        ctk.CTkButton(btn_row, text="Frissítés most", height=42, corner_radius=8,
                      fg_color=COLORS["green_dim"], hover_color=COLORS["green"],
                      command=self._update).grid(row=0, column=0, padx=8, sticky="ew")
        ctk.CTkButton(btn_row, text="Később", height=42, corner_radius=8,
                      fg_color=COLORS["card"], hover_color=COLORS["border"],
                      command=self._later).grid(row=0, column=1, padx=8, sticky="ew")

        self.protocol("WM_DELETE_WINDOW", self._later)
        self.wait_visibility(); self.focus()

    def _update(self): self.result = "update"; self.destroy()
    def _later(self): self.result = "later"; self.destroy()


# ── Stat badge widget ─────────────────────────────────────────────────────────
class StatBadge(ctk.CTkFrame):
    def __init__(self, parent, label, value="—", color=None):
        super().__init__(parent, fg_color=COLORS["card"], corner_radius=10)
        self._color = color or COLORS["text"]
        ctk.CTkLabel(self, text=label, font=ctk.CTkFont(size=11),
                     text_color=COLORS["text_dim"]).pack(pady=(10,2))
        self._val_lbl = ctk.CTkLabel(self, text=value,
                                     font=ctk.CTkFont(size=20, weight="bold"),
                                     text_color=self._color)
        self._val_lbl.pack(pady=(0,10))

    def set(self, value, color=None):
        self._val_lbl.configure(text=str(value), text_color=color or self._color)


# ── Main App ──────────────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CikkChecker")
        self.geometry("1380x960")
        self.minsize(1100, 760)
        self.configure(fg_color=COLORS["bg"])

        # State
        self._proc: subprocess.Popen | None = None
        self._stop_requested = False
        self._can_continue   = False
        self._unknown_mode   = None   # None | "skip_all" | "stop_all"
        self._log_q          = queue.Queue()
        self._stats          = {"van": 0, "nincs": 0, "kulso": 0, "ism": 0}

        self._build_ui()
        self._poll_log()
        self._update_start_btn()
        self.protocol("WM_DELETE_WINDOW", self._quit)
        self.after(3000, self._check_update_async)

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── Left sidebar ──────────────────────────────────────────────────────
        sidebar = ctk.CTkScrollableFrame(self, width=320, fg_color=COLORS["panel"],
                                         corner_radius=0, scrollbar_button_color=COLORS["border"])
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_columnconfigure(0, weight=1)

        # App title
        title_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        title_frame.grid(row=0, column=0, padx=18, pady=(22,18), sticky="ew")
        ctk.CTkLabel(title_frame, text="CikkChecker",
                     font=ctk.CTkFont(size=26, weight="bold"),
                     text_color=COLORS["accent"]).pack(side="left")
        ctk.CTkLabel(title_frame, text=f"v{APP_VERSION}",
                     font=ctk.CTkFont(size=12),
                     text_color=COLORS["text_dim"]).pack(side="left", padx=(8,0), pady=(6,0))

        def section(parent, label, row):
            ctk.CTkLabel(parent, text=label.upper(),
                         font=ctk.CTkFont(size=10, weight="bold"),
                         text_color=COLORS["text_dim"]
                         ).grid(row=row, column=0, padx=18, pady=(16,4), sticky="w")

        def card(parent, row):
            f = ctk.CTkFrame(parent, fg_color=COLORS["card"], corner_radius=10)
            f.grid(row=row, column=0, padx=12, pady=(0,4), sticky="ew")
            f.grid_columnconfigure(0, weight=1)
            return f

        def lbl(parent, text, row):
            ctk.CTkLabel(parent, text=text, font=ctk.CTkFont(size=12),
                         text_color=COLORS["text_mid"]).grid(row=row, column=0, padx=14, pady=(10,2), sticky="w")

        def entry(parent, row, placeholder="", show=""):
            e = ctk.CTkEntry(parent, placeholder_text=placeholder, show=show,
                             fg_color=COLORS["panel"], border_color=COLORS["border"],
                             text_color=COLORS["text"], corner_radius=7)
            e.grid(row=row, column=0, padx=14, pady=(0,10), sticky="ew")
            return e

        # ── Connection ────────────────────────────────────────────────────────
        section(sidebar, "Kapcsolat", 1)
        c1 = card(sidebar, 2)
        lbl(c1, "Oldal URL", 0)
        self._url_entry = entry(c1, 1, "https://szakalmetal.hu")
        self._login_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(c1, text="Bejelentkezés szükséges",
                        variable=self._login_var, fg_color=COLORS["accent"],
                        command=self._toggle_login).grid(row=2, column=0, padx=14, pady=(0,12), sticky="w")

        # ── Login ─────────────────────────────────────────────────────────────
        section(sidebar, "Bejelentkezés", 3)
        self._login_card = card(sidebar, 4)
        lbl(self._login_card, "Felhasználónév", 0)
        self._user_entry = entry(self._login_card, 1, "Felhasználónév")
        lbl(self._login_card, "Jelszó", 2)
        self._pass_entry = entry(self._login_card, 3, "Jelszó", show="*")

        # ── Files ─────────────────────────────────────────────────────────────
        section(sidebar, "Fájlok", 5)
        c3 = card(sidebar, 6)
        lbl(c3, "Cikkszám TXT", 0)
        self._txt_entry = entry(c3, 1, "Tallózás vagy drag & drop")
        ctk.CTkButton(c3, text="📂  TXT kiválasztása", height=34, corner_radius=7,
                      fg_color=COLORS["border"], hover_color=COLORS["accent_dim"],
                      command=self._browse_txt).grid(row=2, column=0, padx=14, pady=(0,10), sticky="ew")
        lbl(c3, "Excel kimeneti fájl", 3)
        self._xlsx_entry = entry(c3, 4, "pl. C:\\eredmeny.xlsx")
        ctk.CTkButton(c3, text="💾  Excel hely", height=34, corner_radius=7,
                      fg_color=COLORS["border"], hover_color=COLORS["accent_dim"],
                      command=self._browse_xlsx).grid(row=5, column=0, padx=14, pady=(0,12), sticky="ew")

        # ── Settings ──────────────────────────────────────────────────────────
        section(sidebar, "Beállítások", 7)
        c4 = card(sidebar, 8)
        lbl(c4, "Várakozás kérések közt (mp)", 0)
        self._sleep_entry = entry(c4, 1, "0.8")
        self._sleep_entry.insert(0, "0.8")
        lbl(c4, "Excel mentés ennyinként", 2)
        self._save_entry = entry(c4, 3, "100")
        self._save_entry.insert(0, "100")

        self._auto_rm_var   = ctk.BooleanVar(value=False)
        self._dont_ask_var  = ctk.BooleanVar(value=False)
        self._save_dup_var  = ctk.BooleanVar(value=True)

        def chk(parent, text, var, row):
            ctk.CTkCheckBox(parent, text=text, variable=var,
                            fg_color=COLORS["accent"], font=ctk.CTkFont(size=12),
                            ).grid(row=row, column=0, padx=14, pady=4, sticky="w")

        chk(c4, "Duplikátumok auto eltávolítása", self._auto_rm_var, 4)
        chk(c4, "Ne kérdezzen rá (auto alkalmaz)", self._dont_ask_var, 5)
        ctk.CTkFrame(c4, height=1, fg_color=COLORS["border"]).grid(row=6, column=0, padx=14, pady=6, sticky="ew")
        chk(c4, "Duplikátumok mentése TXT-be", self._save_dup_var, 7)

        # Pad at bottom
        ctk.CTkFrame(c4, height=10, fg_color="transparent").grid(row=8, column=0)

        # ── Action buttons ────────────────────────────────────────────────────
        section(sidebar, "Műveletek", 9)
        c5 = card(sidebar, 10)
        c5.grid_columnconfigure((0,1), weight=1)

        self._start_btn = ctk.CTkButton(
            c5, text="▶  Indítás", height=44, corner_radius=8,
            fg_color=COLORS["green_dim"], hover_color=COLORS["green"],
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._start_action)
        self._start_btn.grid(row=0, column=0, columnspan=2, padx=14, pady=(14,6), sticky="ew")

        self._stop_btn = ctk.CTkButton(
            c5, text="⏹  Stop", height=38, corner_radius=8,
            fg_color=COLORS["red_dim"], hover_color=COLORS["red"],
            command=self._stop)
        self._stop_btn.grid(row=1, column=0, padx=(14,6), pady=6, sticky="ew")

        self._restart_btn = ctk.CTkButton(
            c5, text="↺  Újra", height=38, corner_radius=8,
            fg_color=COLORS["amber_dim"], hover_color=COLORS["amber"],
            text_color="black", command=self._restart)
        self._restart_btn.grid(row=1, column=1, padx=(6,14), pady=6, sticky="ew")

        ctk.CTkButton(c5, text="🔄  Frissítés keresése", height=34, corner_radius=8,
                      fg_color=COLORS["border"], hover_color=COLORS["accent_dim"],
                      command=self._check_update_async
                      ).grid(row=2, column=0, columnspan=2, padx=14, pady=6, sticky="ew")
        ctk.CTkButton(c5, text="🗑  Log törlése", height=34, corner_radius=8,
                      fg_color=COLORS["border"], hover_color=COLORS["border"],
                      command=self._clear_log
                      ).grid(row=3, column=0, columnspan=2, padx=14, pady=(6,14), sticky="ew")

        self._toggle_login()

        # ── Right main area ───────────────────────────────────────────────────
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        # ── Stat badges row ───────────────────────────────────────────────────
        stats_row = ctk.CTkFrame(main, fg_color="transparent")
        stats_row.grid(row=0, column=0, sticky="ew", pady=(0,14))
        for i in range(5): stats_row.grid_columnconfigure(i, weight=1)

        self._badge_total  = StatBadge(stats_row, "Összes",   "—", COLORS["text"])
        self._badge_done   = StatBadge(stats_row, "Kész",     "0", COLORS["accent"])
        self._badge_van    = StatBadge(stats_row, "Van",      "0", COLORS["green"])
        self._badge_nincs  = StatBadge(stats_row, "Nincs",    "0", COLORS["red"])
        self._badge_kulso  = StatBadge(stats_row, "Külső raktár", "0", COLORS["amber"])

        for i, b in enumerate([self._badge_total, self._badge_done,
                                self._badge_van, self._badge_nincs, self._badge_kulso]):
            b.grid(row=0, column=i, padx=5, sticky="ew")

        # ── Progress + status ─────────────────────────────────────────────────
        prog_card = ctk.CTkFrame(main, fg_color=COLORS["card"], corner_radius=10)
        prog_card.grid(row=1, column=0, sticky="ew", pady=(0,14))
        prog_card.grid_columnconfigure(0, weight=1)

        status_row = ctk.CTkFrame(prog_card, fg_color="transparent")
        status_row.grid(row=0, column=0, padx=16, pady=(14,6), sticky="ew")
        status_row.grid_columnconfigure(0, weight=1)

        self._status_lbl = ctk.CTkLabel(status_row, text="Várakozás",
                                        font=ctk.CTkFont(size=16, weight="bold"),
                                        text_color=COLORS["text"])
        self._status_lbl.grid(row=0, column=0, sticky="w")

        self._pct_lbl = ctk.CTkLabel(status_row, text="0 / 0  (0%)",
                                     font=ctk.CTkFont(size=13),
                                     text_color=COLORS["text_dim"])
        self._pct_lbl.grid(row=0, column=1, sticky="e")

        self._progress = ctk.CTkProgressBar(prog_card, height=10, corner_radius=5,
                                            progress_color=COLORS["accent"],
                                            fg_color=COLORS["panel"])
        self._progress.grid(row=1, column=0, padx=16, pady=(0,14), sticky="ew")
        self._progress.set(0)

        # ── Log + manual codes (side by side) ────────────────────────────────
        bottom = ctk.CTkFrame(main, fg_color="transparent")
        bottom.grid(row=2, column=0, sticky="nsew")
        bottom.grid_columnconfigure(0, weight=3)
        bottom.grid_columnconfigure(1, weight=1)
        bottom.grid_rowconfigure(0, weight=1)

        log_card = ctk.CTkFrame(bottom, fg_color=COLORS["card"], corner_radius=10)
        log_card.grid(row=0, column=0, padx=(0,10), sticky="nsew")
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(log_card, text="Futási napló",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=COLORS["text_mid"]).grid(row=0, column=0, padx=14, pady=(12,6), sticky="w")

        self._log_box = ctk.CTkTextbox(log_card,
                                       font=ctk.CTkFont(family="Consolas", size=12),
                                       fg_color=COLORS["panel"],
                                       text_color=COLORS["text"],
                                       corner_radius=8)
        self._log_box.grid(row=1, column=0, padx=12, pady=(0,12), sticky="nsew")

        manual_card = ctk.CTkFrame(bottom, fg_color=COLORS["card"], corner_radius=10)
        manual_card.grid(row=0, column=1, sticky="nsew")
        manual_card.grid_columnconfigure(0, weight=1)
        manual_card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(manual_card, text="Kézi cikkszámok",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=COLORS["text_mid"]).grid(row=0, column=0, padx=14, pady=(12,2), sticky="w")
        ctk.CTkLabel(manual_card, text="Soronként 1 db",
                     font=ctk.CTkFont(size=11),
                     text_color=COLORS["text_dim"]).grid(row=1, column=0, padx=14, pady=(0,6), sticky="w")

        self._manual_box = ctk.CTkTextbox(manual_card,
                                          font=ctk.CTkFont(family="Consolas", size=12),
                                          fg_color=COLORS["panel"],
                                          text_color=COLORS["text"],
                                          corner_radius=8)
        self._manual_box.grid(row=2, column=0, padx=12, pady=(0,12), sticky="nsew")

    # ── UI helpers ────────────────────────────────────────────────────────────
    def _toggle_login(self):
        if self._login_var.get():
            self._login_card.grid(row=4, column=0, padx=12, pady=(0,4), sticky="ew")
        else:
            self._login_card.grid_remove()

    def _update_start_btn(self):
        if self._can_continue:
            self._start_btn.configure(text="▶  Folytatás")
        else:
            self._start_btn.configure(text="▶  Indítás")

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_q.put(f"[{ts}] {msg}")

    def _poll_log(self):
        lines = []
        try:
            while True: lines.append(self._log_q.get_nowait())
        except queue.Empty: pass
        if lines:
            self._log_box.insert("end", "\n".join(lines) + "\n")
            self._log_box.see("end")
        self.after(200, self._poll_log)

    def _set_status(self, msg):
        self.after(0, lambda: self._status_lbl.configure(text=msg))

    def _set_progress(self, cur, tot):
        pct = int(cur / tot * 100) if tot else 0
        def _u():
            self._progress.set(0 if not tot else cur / tot)
            self._pct_lbl.configure(text=f"{cur} / {tot}  ({pct}%)")
            self._badge_done.set(cur)
            done = self._stats["van"] + self._stats["nincs"] + self._stats["kulso"]
            # badge_done shows processed count
        self.after(0, _u)

    def _update_badges(self):
        self._badge_van.set(self._stats["van"])
        self._badge_nincs.set(self._stats["nincs"])
        self._badge_kulso.set(self._stats["kulso"])

    def _clear_log(self):
        self._log_box.delete("1.0", "end")

    def _browse_txt(self):
        p = filedialog.askopenfilename(title="Cikkszám TXT",
                                       filetypes=[("Text", "*.txt"), ("Minden", "*.*")])
        if p:
            self._txt_entry.delete(0, "end")
            self._txt_entry.insert(0, p)

    def _browse_xlsx(self):
        p = filedialog.asksaveasfilename(title="Excel kimeneti fájl",
                                         defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")])
        if p:
            self._xlsx_entry.delete(0, "end")
            self._xlsx_entry.insert(0, p)

    def _quit(self):
        self._stop()
        self.destroy()

    # ── Process control ───────────────────────────────────────────────────────
    def _start_action(self):
        if self._can_continue:
            self._run(skip_dup=True)
        else:
            self._run(skip_dup=False)

    def _stop(self):
        self._stop_requested = True
        if self._proc:
            try: self._proc.stdin.write("stop\n"); self._proc.stdin.flush()
            except: pass
        self._can_continue = True
        self._update_start_btn()
        self._log("Leállítás kérve...")

    def _restart(self):
        self._stop_requested = True
        if self._proc:
            try: self._proc.stdin.write("stop\n"); self._proc.stdin.flush()
            except: pass
        self._can_continue = False
        self._stats = {"van": 0, "nincs": 0, "kulso": 0, "ism": 0}
        self._update_start_btn()
        self._log("Újraindítás...")
        self.after(800, self._run_fresh)

    def _run_fresh(self):
        self._stop_requested = False
        self._run(skip_dup=False, fresh=True)

    def _collect_codes(self):
        codes = []
        txt = self._txt_entry.get().strip()
        manual = self._manual_box.get("1.0", "end").strip()
        if txt and os.path.exists(txt):
            with open(txt, encoding="utf-8") as f:
                codes += [l.strip() for l in f if l.strip()]
        if manual:
            codes += [l.strip() for l in manual.splitlines() if l.strip()]
        return codes

    def _run(self, skip_dup=False, fresh=False):
        if self._proc and self._proc.poll() is None:
            messagebox.showwarning("Figyelem", "A folyamat már fut.")
            return

        base_url = self._url_entry.get().strip()
        user     = self._user_entry.get().strip()
        pw       = self._pass_entry.get().strip()
        xlsx     = self._xlsx_entry.get().strip()
        req_login = self._login_var.get()

        if not base_url or not xlsx:
            messagebox.showerror("Hiba", "URL és Excel fájl megadása kötelező."); return
        if req_login and (not user or not pw):
            messagebox.showerror("Hiba", "Felhasználónév és jelszó szükséges."); return
        if not xlsx.lower().endswith(".xlsx"):
            messagebox.showerror("Hiba", "Az Excel fájlnak .xlsx kiterjesztésűnek kell lennie."); return

        try:
            sleep_s   = float(self._sleep_entry.get().strip())
            save_every = int(self._save_entry.get().strip())
        except ValueError:
            messagebox.showerror("Hiba", "Várakozás és mentés értéknek számnak kell lennie."); return

        codes = self._collect_codes()
        if not codes:
            messagebox.showerror("Hiba", "Adj meg cikkszámokat."); return

        xlsx_folder = os.path.dirname(xlsx) or "."
        os.makedirs(xlsx_folder, exist_ok=True)
        csv_path = os.path.join(xlsx_folder, "cikkchecker_results.csv")

        if fresh and os.path.exists(csv_path):
            try: os.remove(csv_path)
            except: pass

        # Duplicate handling
        counter = Counter(codes)
        dup_items = {c: n for c,n in counter.items() if n > 1}
        unique_codes = list(dict.fromkeys(codes))

        if dup_items and not skip_dup:
            dup_rows = sum(n-1 for n in counter.values() if n > 1)

            if self._dont_ask_var.get():
                if self._save_dup_var.get():
                    self._save_dup_txt(xlsx_folder, dup_items)
                if self._auto_rm_var.get():
                    codes = unique_codes
            else:
                dlg = DuplicateDialog(self, dup_items, len(codes),
                                      len(unique_codes), dup_rows)
                self.wait_window(dlg)
                if not dlg.result_ok: return
                if dlg.result_save: self._save_dup_txt(xlsx_folder, dup_items)
                if dlg.result_remove: codes = unique_codes

        # Write codes to temp file for Rust binary
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".txt",
                                          encoding="utf-8", delete=False)
        tmp.write("\n".join(codes))
        tmp.close()

        self._badge_total.set(len(set(codes)))
        self._stats = {"van": 0, "nincs": 0, "kulso": 0, "ism": 0}
        self._progress.set(0)
        self._stop_requested = False
        self._unknown_mode   = None

        if not os.path.exists(BINARY_PATH):
            messagebox.showerror("Hiba",
                f"A Rust bináris nem található:\n{BINARY_PATH}\n\n"
                "Futtasd: cargo build --release\n"
                "és másold a checker_core[.exe] fájlt az app mellé.")
            return

        cmd = [
            BINARY_PATH,
            "--base-url", base_url,
            "--user", user,
            "--password", pw,
            "--codes-file", tmp.name,
            "--csv-output", csv_path,
            "--sleep-seconds", str(sleep_s),
            "--save-every", str(save_every),
        ]
        if req_login:
            cmd.append("--requires-login")

        self._log("=" * 48)
        self._log(f"Folyamat indítása — {len(codes)} cikkszám")
        self._log(f"Bináris: {BINARY_PATH}")
        self._set_status("Indítás...")
        self._can_continue = True
        self._update_start_btn()

        self._proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True, bufsize=1
        )
        self._xlsx_path = xlsx
        self._csv_path  = csv_path

        threading.Thread(target=self._read_proc, args=(tmp.name,), daemon=True).start()

    def _save_dup_txt(self, folder, dup_items):
        p = os.path.join(folder, "duplicates.txt")
        with open(p, "w", encoding="utf-8") as f:
            for c, n in dup_items.items():
                f.write(f"{c};{n}\n")
        self._log(f"Duplikátumok mentve: {p}")

    def _read_proc(self, tmp_path: str):
        """Reads JSON-line output from the Rust binary and dispatches to UI."""
        try:
            for line in self._proc.stdout:
                line = line.strip()
                if not line: continue
                try:
                    msg = json.loads(line)
                except json.JSONDecodeError:
                    self._log(line)
                    continue

                kind = msg.get("kind", "")

                if kind == "log":
                    self._log(msg["msg"])

                elif kind == "status":
                    self._set_status(msg["msg"])

                elif kind == "progress":
                    self._set_progress(msg["current"], msg["total"])

                elif kind == "result":
                    avail = msg.get("elerhetoseg", "")
                    if avail == "Van":    self._stats["van"]   += 1
                    elif avail == "Nincs": self._stats["nincs"] += 1
                    elif avail == "Külső raktár": self._stats["kulso"] += 1
                    else: self._stats["ism"] += 1
                    self.after(0, self._update_badges)

                elif kind == "unknown":
                    code = msg["cikkszam"]
                    cmd  = self._handle_unknown(code)
                    try:
                        self._proc.stdin.write(cmd + "\n")
                        self._proc.stdin.flush()
                    except: pass

                elif kind == "error":
                    self._log(f"[HIBA] {msg['msg']}")

                elif kind == "done":
                    self._log("Folyamat kész – Excel exportálás...")
                    export_to_excel(self._csv_path, self._xlsx_path)
                    self._log(f"Excel mentve: {self._xlsx_path}")
                    self._set_status("Kész ✓")
                    self._can_continue = False
                    self.after(0, self._update_start_btn)

        except Exception as e:
            self._log(f"Olvasási hiba: {e}")
        finally:
            try: os.unlink(tmp_path)
            except: pass

    def _handle_unknown(self, cikkszam: str) -> str:
        if self._unknown_mode == "skip_all": return "skip"
        if self._unknown_mode == "stop_all": return "stop"

        result_box = {"action": "stop", "apply": False}
        ev = threading.Event()

        def _show():
            dlg = UnknownDialog(self, cikkszam)
            self.wait_window(dlg)
            result_box["action"] = dlg.action
            result_box["apply"]  = dlg.apply_all
            ev.set()

        self.after(0, _show)
        ev.wait()

        if result_box["apply"]:
            if result_box["action"] == "skip": self._unknown_mode = "skip_all"
            else:                              self._unknown_mode = "stop_all"

        return result_box["action"]

    # ── Update check ──────────────────────────────────────────────────────────
    def _check_update_async(self):
        threading.Thread(target=self._check_update, daemon=True).start()

    def _check_update(self):
        url = f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest"
        self._log("Frissítés ellenőrzése...")
        try:
            r = requests.get(url, timeout=15); r.raise_for_status()
            data = r.json()
        except Exception as e:
            self._log(f"Frissítés ellenőrzése sikertelen: {e}"); return

        latest = data.get("tag_name", "").strip()
        notes  = data.get("body", "").strip()

        def _parse(v):
            v = v.strip().lstrip("v")
            parts = []
            for p in v.split("."):
                try: parts.append(int(p))
                except: parts.append(0)
            return tuple(parts)

        if not latest or _parse(latest) <= _parse(APP_VERSION):
            self._log("Nem érhető el újabb verzió."); return

        asset_url = next((a["browser_download_url"] for a in data.get("assets", [])
                          if a.get("name") == UPDATE_ASSET), None)
        if not asset_url:
            self._log("Új verzió van, de a telepítő nem található."); return

        self._log(f"Új verzió: {latest}")
        self.after(0, lambda: self._prompt_update(latest, notes, asset_url))

    def _prompt_update(self, latest, notes, asset_url):
        dlg = UpdateDialog(self, APP_VERSION, latest, notes)
        self.wait_window(dlg)
        if dlg.result == "update":
            threading.Thread(target=self._download_update,
                             args=(asset_url, latest), daemon=True).start()
        else:
            self._log("Frissítés elhalasztva.")

    def _download_update(self, asset_url, latest):
        self._log(f"Letöltés: {latest}...")
        self._set_status(f"Frissítés letöltése ({latest})...")
        try:
            tmp = os.path.join(tempfile.gettempdir(), UPDATE_ASSET)
            with requests.get(asset_url, stream=True, timeout=60) as r:
                r.raise_for_status()
                with open(tmp, "wb") as f:
                    for chunk in r.iter_content(8192):
                        if chunk: f.write(chunk)
            self._log(f"Telepítő letöltve: {tmp}")
            import subprocess as sp
            sp.Popen([tmp], shell=True)
            self.after(1000, self._quit)
        except Exception as e:
            self._log(f"Letöltési hiba: {e}")
            self._set_status("Frissítés sikertelen")


if __name__ == "__main__":
    app = App()
    app.mainloop()
