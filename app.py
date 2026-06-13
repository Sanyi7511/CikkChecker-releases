"""
CikkChecker v3 – Ultra Modern Dark UI
"""
import json, os, queue, subprocess, sys, tempfile, threading, re
from collections import Counter
from datetime import datetime
from tkinter import filedialog
import tkinter as tk
import tkinter.messagebox as messagebox

import customtkinter as ctk
import requests

APP_VERSION  = "3.0.0"
UPDATE_REPO  = "Sanyi7511/CikkChecker-releases"
UPDATE_ASSET = "CikkCheckerSetup.exe"
BINARY_NAME  = "checker_core.exe" if sys.platform == "win32" else "checker_core"
BINARY_PATH  = os.path.join(os.path.dirname(os.path.abspath(__file__)), BINARY_NAME)

C = {
    "bg":         "#08090E",
    "sidebar":    "#0D1018",
    "card":       "#13161F",
    "card2":      "#181C28",
    "border":     "#1F2535",
    "border2":    "#2A3147",
    "accent":     "#5B8DEF",
    "accent2":    "#3B6DD4",
    "green":      "#23C55E",
    "green2":     "#16A34A",
    "green_bg":   "#0A2016",
    "red":        "#F04747",
    "red2":       "#C53030",
    "red_bg":     "#200A0A",
    "amber":      "#F5A623",
    "amber2":     "#D97706",
    "amber_bg":   "#201508",
    "purple":     "#A78BFA",
    "purple_bg":  "#16103A",
    "text":       "#E8EBF0",
    "text2":      "#9BA3B4",
    "text3":      "#4A5568",
}

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def export_excel(csv_path, excel_path):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, GradientFill
        wb = Workbook()
        ws = wb.active
        ws.title = "CikkChecker"

        hdr_fill = PatternFill("solid", fgColor="0D1018")
        fills = {
            "Van":           PatternFill("solid", fgColor="0A2016"),
            "Nincs":         PatternFill("solid", fgColor="200A0A"),
            "Külső raktár":  PatternFill("solid", fgColor="201508"),
            "Ismeretlen":    PatternFill("solid", fgColor="13161F"),
        }
        fonts = {
            "Van":           Font(color="23C55E", bold=True, size=10),
            "Nincs":         Font(color="F04747", bold=True, size=10),
            "Külső raktár":  Font(color="F5A623", bold=True, size=10),
            "Ismeretlen":    Font(color="9BA3B4", size=10),
        }
        thin = Side(style="thin", color="1F2535")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        headers = ["Cikkszám", "Elérhetőség", "Ár"]
        ws.append(headers)
        for i, cell in enumerate(ws[1], 1):
            cell.fill = hdr_fill
            cell.font = Font(color="9BA3B4", bold=True, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        import csv as _csv
        if os.path.exists(csv_path):
            with open(csv_path, encoding="utf-8-sig", newline="") as f:
                for row in _csv.DictReader(f):
                    c = (row.get("Cikkszam") or row.get("Cikkszám") or "").strip()
                    a = (row.get("Elerhetoseg") or row.get("Elérhetőség") or "").strip()
                    p = (row.get("Ar") or row.get("Ár") or "").strip()
                    if not c: continue
                    ws.append([c, a, p])
                    r = ws.max_row
                    for col in range(1, 4):
                        cell = ws.cell(r, col)
                        cell.border = border
                        cell.alignment = Alignment(horizontal="left" if col == 1 else "center", vertical="center")
                        cell.fill = fills.get(a, fills["Ismeretlen"])
                    ws.cell(r, 2).font = fonts.get(a, fonts["Ismeretlen"])
                    ws.cell(r, 1).font = Font(color="E8EBF0", size=10)
                    ws.cell(r, 3).font = Font(color="9BA3B4", size=10)

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 16
        ws.row_dimensions[1].height = 22
        wb.save(excel_path)
    except Exception as e:
        print(f"Excel export hiba: {e}")


# ── Dialogs ───────────────────────────────────────────────────────────────────
class ModernDialog(ctk.CTkToplevel):
    def __init__(self, parent, title, width=560, height=400):
        super().__init__(parent)
        self.title(title)
        self.geometry(f"{width}x{height}")
        self.configure(fg_color=C["bg"])
        self.transient(parent)
        self.grab_set()
        self.grid_columnconfigure(0, weight=1)
        self.wait_visibility()
        self.focus()

    def _header(self, icon, title, color, row=0):
        f = ctk.CTkFrame(self, fg_color=C["card"], corner_radius=0)
        f.grid(row=row, column=0, sticky="ew", padx=0, pady=0)
        f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(f, text=icon, font=ctk.CTkFont(size=22), text_color=color
                     ).grid(row=0, column=0, padx=(20,8), pady=16)
        ctk.CTkLabel(f, text=title, font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=C["text"]).grid(row=0, column=1, pady=16, sticky="w")

    def _btn_row(self, row, btns):
        f = ctk.CTkFrame(self, fg_color="transparent")
        f.grid(row=row, column=0, padx=20, pady=(0,20), sticky="ew")
        for i in range(len(btns)): f.grid_columnconfigure(i, weight=1)
        for i, (text, cmd, fg, hover) in enumerate(btns):
            ctk.CTkButton(f, text=text, command=cmd, height=40, corner_radius=8,
                          fg_color=fg, hover_color=hover,
                          font=ctk.CTkFont(size=13, weight="bold")
                          ).grid(row=0, column=i, padx=6, sticky="ew")


class DuplicateDialog(ModernDialog):
    def __init__(self, parent, dup_items, total, unique_count, dup_rows):
        super().__init__(parent, "Duplikátumok", 620, 500)
        self.ok = False; self.remove = False; self.save = False
        self.grid_rowconfigure(1, weight=1)

        self._header("⚠", f"Duplikátumok találva  —  {dup_rows} felesleges sor", C["amber"])

        card = ctk.CTkFrame(self, fg_color=C["card"], corner_radius=10)
        card.grid(row=1, column=0, padx=16, pady=(12,8), sticky="nsew")
        card.grid_columnconfigure(0, weight=1); card.grid_rowconfigure(1, weight=1)

        meta = f"Összes: {total}   •   Egyedi: {unique_count}   •   Duplikált sorok: {dup_rows}"
        ctk.CTkLabel(card, text=meta, text_color=C["text2"],
                     font=ctk.CTkFont(size=12)).grid(row=0, column=0, padx=14, pady=(12,6), sticky="w")

        tb = ctk.CTkTextbox(card, font=ctk.CTkFont(family="Consolas", size=11),
                            fg_color=C["card2"], text_color=C["text2"], corner_radius=6)
        tb.grid(row=1, column=0, padx=12, pady=(0,12), sticky="nsew")
        for code, count in dup_items.items():
            tb.insert("end", f"  {code}  ×{count}\n")
        tb.configure(state="disabled")

        opts = ctk.CTkFrame(self, fg_color=C["card"], corner_radius=10)
        opts.grid(row=2, column=0, padx=16, pady=4, sticky="ew")
        self._rm  = ctk.BooleanVar(value=True)
        self._sav = ctk.BooleanVar(value=True)
        for i, (text, var) in enumerate([
            ("Minden cikkszámból csak 1 maradjon", self._rm),
            ("Duplikátumok mentése TXT fájlba", self._sav),
        ]):
            ctk.CTkCheckBox(opts, text=text, variable=var, fg_color=C["accent"],
                            font=ctk.CTkFont(size=12)
                            ).grid(row=i, column=0, padx=14, pady=(10 if i==0 else 4, 10 if i==1 else 4), sticky="w")

        self._btn_row(3, [
            ("Folytatás", self._ok, C["accent2"], C["accent"]),
            ("Mégse",     self._cancel, C["card"], C["border2"]),
        ])
        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _ok(self):     self.ok=True; self.remove=self._rm.get(); self.save=self._sav.get(); self.destroy()
    def _cancel(self): self.destroy()


class UnknownDialog(ModernDialog):
    def __init__(self, parent, cikkszam):
        super().__init__(parent, "Ismeretlen találat", 440, 220)
        self.action="stop"; self.apply_all=False
        self._header("?", f"Nem értelmezhető:  {cikkszam}", C["amber"])

        ctk.CTkLabel(self, text="A cikkszám eredménye nem egyértelmű. Mit tegyek?",
                     text_color=C["text2"], font=ctk.CTkFont(size=12)
                     ).grid(row=1, column=0, padx=20, pady=(16,8), sticky="w")
        self._all = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(self, text="Alkalmazza a további ismeretlen találatokra is",
                        variable=self._all, fg_color=C["accent"],
                        font=ctk.CTkFont(size=12)
                        ).grid(row=2, column=0, padx=20, pady=(0,12), sticky="w")
        self._btn_row(3, [
            ("Kihagyás", self._skip, C["card"],  C["border2"]),
            ("Leállítás",self._stop, C["red2"],  C["red"]),
        ])
        self.protocol("WM_DELETE_WINDOW", self._stop)

    def _skip(self): self.action="skip"; self.apply_all=self._all.get(); self.destroy()
    def _stop(self): self.action="stop"; self.apply_all=self._all.get(); self.destroy()


class UpdateDialog(ModernDialog):
    def __init__(self, parent, current, latest, notes):
        super().__init__(parent, "Frissítés", 580, 420)
        self.result="later"
        self.grid_rowconfigure(1, weight=1)
        self._header("🚀", f"Új verzió elérhető:  {latest}  (jelenlegi: {current})", C["green"])

        card = ctk.CTkFrame(self, fg_color=C["card"], corner_radius=10)
        card.grid(row=1, column=0, padx=16, pady=(12,8), sticky="nsew")
        card.grid_columnconfigure(0, weight=1); card.grid_rowconfigure(0, weight=1)
        tb = ctk.CTkTextbox(card, fg_color=C["card2"], text_color=C["text2"])
        tb.grid(row=0, column=0, padx=12, pady=12, sticky="nsew")
        tb.insert("1.0", notes.strip() or "Nincs kiadási megjegyzés.")
        tb.configure(state="disabled")

        self._btn_row(2, [
            ("Frissítés most", self._upd, C["green2"],  C["green"]),
            ("Később",         self._later, C["card"],  C["border2"]),
        ])
        self.protocol("WM_DELETE_WINDOW", self._later)

    def _upd(self):   self.result="update"; self.destroy()
    def _later(self): self.result="later";  self.destroy()


# ── Stat card ─────────────────────────────────────────────────────────────────
class StatCard(ctk.CTkFrame):
    def __init__(self, parent, label, value="—", accent=None, icon=""):
        super().__init__(parent, fg_color=C["card"], corner_radius=12,
                         border_width=1, border_color=C["border"])
        self._accent = accent or C["text2"]
        self.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(self, fg_color="transparent")
        top.grid(row=0, column=0, padx=16, pady=(14,4), sticky="ew")
        ctk.CTkLabel(top, text=f"{icon}  {label}" if icon else label,
                     font=ctk.CTkFont(size=11), text_color=C["text3"]).pack(side="left")

        self._val = ctk.CTkLabel(self, text=str(value),
                                 font=ctk.CTkFont(size=26, weight="bold"),
                                 text_color=self._accent)
        self._val.grid(row=1, column=0, padx=16, pady=(0,14), sticky="w")

    def set(self, v, color=None):
        self._val.configure(text=str(v), text_color=color or self._accent)


# ── Result table row ──────────────────────────────────────────────────────────
class ResultTable(ctk.CTkScrollableFrame):
    AVAIL_STYLE = {
        "Van":          (C["green"],  C["green_bg"],  "●"),
        "Nincs":        (C["red"],    C["red_bg"],    "●"),
        "Külső raktár": (C["amber"],  C["amber_bg"],  "●"),
        "Ismeretlen":   (C["text3"],  C["card2"],     "○"),
    }

    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=2)
        self.grid_columnconfigure(2, weight=2)
        self._rows = 0
        self._draw_header()

    def _draw_header(self):
        for col, (text, anchor) in enumerate([("Cikkszám","w"),("Elérhetőség","center"),("Ár","center")]):
            ctk.CTkLabel(self, text=text, font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=C["text3"]
                         ).grid(row=0, column=col, padx=(14 if col==0 else 4), pady=(4,6),
                                sticky=anchor)

    def add_row(self, cikkszam, avail, price=""):
        r = self._rows + 1
        self._rows += 1
        color, bg, dot = self.AVAIL_STYLE.get(avail, (C["text2"], C["card2"], "○"))

        row_bg = C["card"] if r % 2 == 0 else C["card2"]

        ctk.CTkLabel(self, text=cikkszam, font=ctk.CTkFont(family="Consolas", size=12),
                     text_color=C["text"], fg_color=row_bg, anchor="w", corner_radius=0
                     ).grid(row=r, column=0, padx=(12,4), pady=1, sticky="ew")

        avail_frame = ctk.CTkFrame(self, fg_color=bg, corner_radius=6)
        avail_frame.grid(row=r, column=1, padx=4, pady=1, sticky="ew")
        ctk.CTkLabel(avail_frame, text=f"{dot}  {avail}",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=color).pack(pady=3, padx=8)

        ctk.CTkLabel(self, text=price if price else "—",
                     font=ctk.CTkFont(size=12), text_color=C["text2"],
                     fg_color=row_bg, anchor="center", corner_radius=0
                     ).grid(row=r, column=2, padx=(4,12), pady=1, sticky="ew")

    def clear(self):
        for w in self.winfo_children():
            if int(w.grid_info().get("row", 0)) > 0:
                w.destroy()
        self._rows = 0


# ── Sidebar section label ─────────────────────────────────────────────────────
def sidebar_section(parent, text, row):
    ctk.CTkLabel(parent, text=text.upper(),
                 font=ctk.CTkFont(size=9, weight="bold"),
                 text_color=C["text3"]
                 ).grid(row=row, column=0, padx=20, pady=(18,4), sticky="w")

def sidebar_card(parent, row):
    f = ctk.CTkFrame(parent, fg_color=C["card"], corner_radius=10,
                     border_width=1, border_color=C["border"])
    f.grid(row=row, column=0, padx=12, pady=(0,2), sticky="ew")
    f.grid_columnconfigure(0, weight=1)
    return f

def field_label(parent, text, row):
    ctk.CTkLabel(parent, text=text, font=ctk.CTkFont(size=11),
                 text_color=C["text3"]).grid(row=row, column=0, padx=14, pady=(10,2), sticky="w")

def field_entry(parent, row, placeholder="", show=""):
    e = ctk.CTkEntry(parent, placeholder_text=placeholder, show=show,
                     fg_color=C["card2"], border_color=C["border2"],
                     text_color=C["text"], placeholder_text_color=C["text3"],
                     corner_radius=7, height=34)
    e.grid(row=row, column=0, padx=14, pady=(0,10), sticky="ew")
    return e

def action_btn(parent, text, row, col, fg, hover, cmd, span=1, bold=False):
    b = ctk.CTkButton(parent, text=text, command=cmd, height=38, corner_radius=8,
                      fg_color=fg, hover_color=hover,
                      font=ctk.CTkFont(size=12, weight="bold" if bold else "normal"))
    b.grid(row=row, column=col, columnspan=span, padx=6, pady=4, sticky="ew")
    return b


# ── Main App ──────────────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CikkChecker")
        self.geometry("1440x900")
        self.minsize(1100, 700)
        self.configure(fg_color=C["bg"])

        self._proc = None
        self._stop_req = False
        self._can_continue = False
        self._unknown_mode = None
        self._log_q = queue.Queue()
        self._stats = {"van":0,"nincs":0,"kulso":0,"ism":0,"total":0}
        self._login_detected = False

        self._build_ui()
        self._poll_log()
        self._update_start_btn()
        self.protocol("WM_DELETE_WINDOW", self._quit)
        self.after(3000, lambda: threading.Thread(target=self._check_update, daemon=True).start())

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ════════════════════════════════════════════════════════════════════
        # LEFT SIDEBAR
        # ════════════════════════════════════════════════════════════════════
        sb = ctk.CTkScrollableFrame(self, width=300, fg_color=C["sidebar"],
                                    corner_radius=0, border_width=0,
                                    scrollbar_button_color=C["border"])
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_columnconfigure(0, weight=1)

        # Logo
        logo = ctk.CTkFrame(sb, fg_color="transparent")
        logo.grid(row=0, column=0, padx=20, pady=(24,4), sticky="ew")
        ctk.CTkLabel(logo, text="Cikk", font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=C["text"]).pack(side="left")
        ctk.CTkLabel(logo, text="Checker", font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=C["accent"]).pack(side="left")
        ctk.CTkLabel(logo, text=f"  v{APP_VERSION}",
                     font=ctk.CTkFont(size=11), text_color=C["text3"]).pack(side="left", pady=(4,0))

        # Divider
        ctk.CTkFrame(sb, height=1, fg_color=C["border"]).grid(
            row=1, column=0, padx=20, pady=(8,0), sticky="ew")

        # ── Connection ────────────────────────────────────────────────────
        sidebar_section(sb, "Kapcsolat", 2)
        c1 = sidebar_card(sb, 3)
        field_label(c1, "Weboldal URL", 0)
        self._url = field_entry(c1, 1, "https://")

        # Login auto-detect badge
        detect_row = ctk.CTkFrame(c1, fg_color="transparent")
        detect_row.grid(row=2, column=0, padx=14, pady=(0,4), sticky="ew")
        detect_row.grid_columnconfigure(0, weight=1)
        self._login_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(detect_row, text="Bejelentkezés szükséges",
                        variable=self._login_var, fg_color=C["accent"],
                        font=ctk.CTkFont(size=12),
                        command=self._toggle_login
                        ).grid(row=0, column=0, sticky="w")
        self._detect_lbl = ctk.CTkLabel(detect_row, text="",
                                        font=ctk.CTkFont(size=10),
                                        text_color=C["text3"])
        self._detect_lbl.grid(row=1, column=0, pady=(2,6), sticky="w")

        # URL blur → auto-detect
        self._url.bind("<FocusOut>", self._on_url_blur)

        # ── Login ─────────────────────────────────────────────────────────
        sidebar_section(sb, "Bejelentkezés", 4)
        self._login_card = sidebar_card(sb, 5)
        field_label(self._login_card, "Felhasználónév", 0)
        self._user = field_entry(self._login_card, 1)
        field_label(self._login_card, "Jelszó", 2)
        self._pw   = field_entry(self._login_card, 3, show="*")

        # ── Fájlok ────────────────────────────────────────────────────────
        sidebar_section(sb, "Fájlok", 6)
        c3 = sidebar_card(sb, 7)
        field_label(c3, "Cikkszám TXT", 0)
        self._txt  = field_entry(c3, 1)
        action_btn(c3, "📂  TXT kiválasztása", 2, 0, C["border2"], C["border"], self._browse_txt, span=1)
        field_label(c3, "Excel kimeneti fájl", 3)
        self._xlsx = field_entry(c3, 4)
        action_btn(c3, "💾  Excel mentési helye", 5, 0, C["border2"], C["border"], self._browse_xlsx, span=1)
        ctk.CTkFrame(c3, height=6, fg_color="transparent").grid(row=6, column=0)

        # ── Beállítások ───────────────────────────────────────────────────
        sidebar_section(sb, "Beállítások", 8)
        c4 = sidebar_card(sb, 9)
        c4.grid_columnconfigure((0,1), weight=1)

        field_label(c4, "Várakozás (mp)", 0)
        self._sleep = ctk.CTkEntry(c4, fg_color=C["card2"], border_color=C["border2"],
                                   text_color=C["text"], corner_radius=7, height=32)
        self._sleep.insert(0, "0.8")
        self._sleep.grid(row=1, column=0, padx=(14,6), pady=(0,10), sticky="ew")

        field_label(c4, "Mentés ennyinként", 0)
        self._save_ev = ctk.CTkEntry(c4, fg_color=C["card2"], border_color=C["border2"],
                                     text_color=C["text"], corner_radius=7, height=32)
        self._save_ev.insert(0, "100")
        self._save_ev.grid(row=1, column=1, padx=(6,14), pady=(0,10), sticky="ew")

        # Sort order
        field_label(c4, "Sorrend", 2)
        self._sort_var = ctk.StringVar(value="abc")
        sort_menu = ctk.CTkOptionMenu(
            c4, values=["ABC (A→Z)", "ABC (Z→A)", "Szám (0→9)", "Szám (9→0)", "Eredeti sorrend"],
            variable=self._sort_var,
            fg_color=C["card2"], button_color=C["border2"],
            button_hover_color=C["accent2"], text_color=C["text"],
            dropdown_fg_color=C["card"], dropdown_hover_color=C["border2"],
            font=ctk.CTkFont(size=12), corner_radius=7
        )
        sort_menu.grid(row=3, column=0, columnspan=2, padx=14, pady=(0,10), sticky="ew")

        # Dup handling
        self._auto_rm  = ctk.BooleanVar(value=False)
        self._dont_ask = ctk.BooleanVar(value=False)
        self._save_dup = ctk.BooleanVar(value=True)
        for i, (text, var) in enumerate([
            ("Auto duplikátum eltávolítás", self._auto_rm),
            ("Ne kérdezzen rá, automatikusan", self._dont_ask),
            ("Duplikátumok mentése TXT-be", self._save_dup),
        ]):
            ctk.CTkCheckBox(c4, text=text, variable=var, fg_color=C["accent"],
                            font=ctk.CTkFont(size=11)
                            ).grid(row=4+i, column=0, columnspan=2, padx=14,
                                   pady=(6 if i==0 else 2, 10 if i==2 else 2), sticky="w")

        # ── Műveletek ─────────────────────────────────────────────────────
        sidebar_section(sb, "Műveletek", 10)
        c5 = sidebar_card(sb, 11)
        c5.grid_columnconfigure((0,1), weight=1)

        self._start_btn = ctk.CTkButton(
            c5, text="▶  Indítás", height=44, corner_radius=8,
            fg_color=C["green2"], hover_color=C["green"],
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._start_action)
        self._start_btn.grid(row=0, column=0, columnspan=2, padx=12, pady=(14,6), sticky="ew")

        self._stop_btn = action_btn(c5, "⏹  Stop",   1, 0, C["red2"],   C["red"],   self._stop)
        self._rest_btn = action_btn(c5, "↺  Újra",   1, 1, C["amber2"], C["amber"], self._restart)
        action_btn(c5, "🔄  Frissítés keresése", 2, 0, C["border2"], C["border"],
                   lambda: threading.Thread(target=self._check_update, daemon=True).start(), span=2)
        action_btn(c5, "🗑  Log törlése", 3, 0, C["border2"], C["border"],
                   self._clear_log, span=2)
        ctk.CTkFrame(c5, height=8, fg_color="transparent").grid(row=4, column=0)

        self._toggle_login()

        # ════════════════════════════════════════════════════════════════════
        # RIGHT MAIN AREA
        # ════════════════════════════════════════════════════════════════════
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        # ── Stat cards row ────────────────────────────────────────────────
        stats = ctk.CTkFrame(main, fg_color="transparent")
        stats.grid(row=0, column=0, sticky="ew", pady=(0,16))
        for i in range(5): stats.grid_columnconfigure(i, weight=1)

        self._s_total  = StatCard(stats, "Összes cikkszám",  "—",  C["text2"],  "📋")
        self._s_done   = StatCard(stats, "Feldolgozva",      "0",  C["accent"], "✓")
        self._s_van    = StatCard(stats, "Elérhető",         "0",  C["green"],  "●")
        self._s_nincs  = StatCard(stats, "Nem elérhető",     "0",  C["red"],    "●")
        self._s_kulso  = StatCard(stats, "Külső raktár",     "0",  C["amber"],  "●")
        for i, w in enumerate([self._s_total,self._s_done,self._s_van,self._s_nincs,self._s_kulso]):
            w.grid(row=0, column=i, padx=5, sticky="ew")

        # ── Progress card ─────────────────────────────────────────────────
        prog = ctk.CTkFrame(main, fg_color=C["card"], corner_radius=12,
                            border_width=1, border_color=C["border"])
        prog.grid(row=1, column=0, sticky="ew", pady=(0,16))
        prog.grid_columnconfigure(0, weight=1)

        top_row = ctk.CTkFrame(prog, fg_color="transparent")
        top_row.grid(row=0, column=0, padx=18, pady=(14,6), sticky="ew")
        top_row.grid_columnconfigure(0, weight=1)

        self._status_lbl = ctk.CTkLabel(top_row, text="Várakozás",
                                        font=ctk.CTkFont(size=15, weight="bold"),
                                        text_color=C["text"])
        self._status_lbl.grid(row=0, column=0, sticky="w")

        self._pct_lbl = ctk.CTkLabel(top_row, text="0 / 0  (0%)",
                                     font=ctk.CTkFont(size=12), text_color=C["text3"])
        self._pct_lbl.grid(row=0, column=1, sticky="e")

        self._progress = ctk.CTkProgressBar(prog, height=6, corner_radius=3,
                                            progress_color=C["accent"], fg_color=C["border"])
        self._progress.grid(row=1, column=0, padx=18, pady=(0,14), sticky="ew")
        self._progress.set(0)

        # ── Bottom: result table + log + manual ───────────────────────────
        bottom = ctk.CTkFrame(main, fg_color="transparent")
        bottom.grid(row=2, column=0, sticky="nsew")
        bottom.grid_columnconfigure(0, weight=5)
        bottom.grid_columnconfigure(1, weight=2)
        bottom.grid_rowconfigure(0, weight=1)

        # Result table
        res_card = ctk.CTkFrame(bottom, fg_color=C["card"], corner_radius=12,
                                border_width=1, border_color=C["border"])
        res_card.grid(row=0, column=0, padx=(0,10), sticky="nsew")
        res_card.grid_columnconfigure(0, weight=1)
        res_card.grid_rowconfigure(1, weight=1)

        res_hdr = ctk.CTkFrame(res_card, fg_color="transparent")
        res_hdr.grid(row=0, column=0, padx=16, pady=(12,6), sticky="ew")
        res_hdr.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(res_hdr, text="Eredmények",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=C["text2"]).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(res_hdr, text="Törlés", width=60, height=26, corner_radius=6,
                      fg_color=C["border"], hover_color=C["border2"],
                      font=ctk.CTkFont(size=11), text_color=C["text3"],
                      command=self._clear_results
                      ).grid(row=0, column=1, sticky="e")

        self._table = ResultTable(res_card, fg_color="transparent")
        self._table.grid(row=1, column=0, padx=8, pady=(0,8), sticky="nsew")

        # Right side: log + manual
        right_side = ctk.CTkFrame(bottom, fg_color="transparent")
        right_side.grid(row=0, column=1, sticky="nsew")
        right_side.grid_columnconfigure(0, weight=1)
        right_side.grid_rowconfigure(0, weight=3)
        right_side.grid_rowconfigure(1, weight=2)

        # Log
        log_card = ctk.CTkFrame(right_side, fg_color=C["card"], corner_radius=12,
                                border_width=1, border_color=C["border"])
        log_card.grid(row=0, column=0, pady=(0,10), sticky="nsew")
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(log_card, text="Napló", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=C["text3"]).grid(row=0, column=0, padx=14, pady=(10,4), sticky="w")
        self._log_box = ctk.CTkTextbox(log_card, font=ctk.CTkFont(family="Consolas", size=10),
                                       fg_color=C["card2"], text_color=C["text2"], corner_radius=8)
        self._log_box.grid(row=1, column=0, padx=10, pady=(0,10), sticky="nsew")

        # Manual codes
        man_card = ctk.CTkFrame(right_side, fg_color=C["card"], corner_radius=12,
                                border_width=1, border_color=C["border"])
        man_card.grid(row=1, column=0, sticky="nsew")
        man_card.grid_columnconfigure(0, weight=1)
        man_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(man_card, text="Kézi cikkszámok",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=C["text3"]).grid(row=0, column=0, padx=14, pady=(10,2), sticky="w")
        self._manual = ctk.CTkTextbox(man_card, font=ctk.CTkFont(family="Consolas", size=11),
                                      fg_color=C["card2"], text_color=C["text2"], corner_radius=8)
        self._manual.grid(row=1, column=0, padx=10, pady=(0,10), sticky="nsew")

    # ── UI helpers ────────────────────────────────────────────────────────────
    def _toggle_login(self):
        if self._login_var.get():
            self._login_card.grid(row=5, column=0, padx=12, pady=(0,2), sticky="ew")
        else:
            self._login_card.grid_remove()

    def _on_url_blur(self, event=None):
        url = self._url.get().strip()
        if not url or url == "https://": return
        threading.Thread(target=self._detect_login, args=(url,), daemon=True).start()

    def _detect_login(self, url):
        self.after(0, lambda: self._detect_lbl.configure(text="⟳  Bejelentkezés detektálása...", text_color=C["text3"]))
        try:
            if not url.startswith("http"): url = "https://" + url
            resp = requests.get(f"{url}/product-search/test/0?1", timeout=8, allow_redirects=True)
            final_url = resp.url
            html = resp.text.lower()
            required = (
                "/login" in final_url or
                "bejelentkezés" in html or "belépés" in html or
                "felhasználónév" in html or "jelszó" in html or
                "user-name" in html or "password" in html
            )
            def _upd():
                self._login_var.set(required)
                self._toggle_login()
                if required:
                    self._detect_lbl.configure(text="🔒  Bejelentkezés szükséges (auto)", text_color=C["amber"])
                else:
                    self._detect_lbl.configure(text="🔓  Bejelentkezés nem szükséges (auto)", text_color=C["green"])
            self.after(0, _upd)
        except Exception:
            self.after(0, lambda: self._detect_lbl.configure(
                text="⚠  Nem sikerült ellenőrizni", text_color=C["text3"]))

    def _update_start_btn(self):
        self._start_btn.configure(text="▶  Folytatás" if self._can_continue else "▶  Indítás")

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_q.put(f"[{ts}]  {msg}")

    def _poll_log(self):
        lines = []
        try:
            while True: lines.append(self._log_q.get_nowait())
        except queue.Empty: pass
        if lines:
            self._log_box.insert("end", "\n".join(lines) + "\n")
            self._log_box.see("end")
        self.after(150, self._poll_log)

    def _set_status(self, msg): self.after(0, lambda: self._status_lbl.configure(text=msg))

    def _set_progress(self, cur, tot):
        pct = int(cur/tot*100) if tot else 0
        def _u():
            self._progress.set(0 if not tot else cur/tot)
            self._pct_lbl.configure(text=f"{cur} / {tot}  ({pct}%)")
            self._s_done.set(cur)
        self.after(0, _u)

    def _update_stat_cards(self):
        self._s_van.set(self._stats["van"])
        self._s_nincs.set(self._stats["nincs"])
        self._s_kulso.set(self._stats["kulso"])

    def _clear_log(self): self._log_box.delete("1.0","end")
    def _clear_results(self): self._table.clear()

    def _browse_txt(self):
        p = filedialog.askopenfilename(title="Cikkszám TXT",
                                       filetypes=[("Text","*.txt"),("Minden","*.*")])
        if p: self._txt.delete(0,"end"); self._txt.insert(0,p)

    def _browse_xlsx(self):
        p = filedialog.asksaveasfilename(title="Excel mentési hely",
                                         defaultextension=".xlsx",
                                         filetypes=[("Excel","*.xlsx")])
        if p: self._xlsx.delete(0,"end"); self._xlsx.insert(0,p)

    def _quit(self): self._stop(); self.destroy()

    def _get_sort_arg(self):
        m = {"ABC (A→Z)":"abc","ABC (Z→A)":"abc_desc",
             "Szám (0→9)":"num","Szám (9→0)":"num_desc","Eredeti sorrend":"none"}
        return m.get(self._sort_var.get(), "abc")

    # ── Process control ───────────────────────────────────────────────────────
    def _start_action(self):
        if self._can_continue: self._run(skip_dup=True)
        else: self._run(skip_dup=False)

    def _stop(self):
        self._stop_req = True
        if self._proc:
            try: self._proc.stdin.write("stop\n"); self._proc.stdin.flush()
            except: pass
        self._can_continue = True
        self._update_start_btn()
        self._log("Leállítás kérve...")

    def _restart(self):
        self._stop_req = True
        if self._proc:
            try: self._proc.stdin.write("stop\n"); self._proc.stdin.flush()
            except: pass
        self._can_continue = False
        self._stats = {"van":0,"nincs":0,"kulso":0,"ism":0,"total":0}
        self._update_start_btn()
        self._log("Újraindítás...")
        self.after(800, lambda: self._run(skip_dup=False, fresh=True))

    def _collect_codes(self):
        codes = []
        txt = self._txt.get().strip()
        manual = self._manual.get("1.0","end").strip()
        if txt and os.path.exists(txt):
            with open(txt, encoding="utf-8") as f:
                codes += [l.strip() for l in f if l.strip()]
        if manual:
            codes += [l.strip() for l in manual.splitlines() if l.strip()]
        return codes

    def _run(self, skip_dup=False, fresh=False):
        if self._proc and self._proc.poll() is None:
            messagebox.showwarning("Figyelem", "A folyamat már fut."); return

        base_url = self._url.get().strip()
        user     = self._user.get().strip()
        pw       = self._pw.get().strip()
        xlsx     = self._xlsx.get().strip()
        req_login = self._login_var.get()

        if not base_url or base_url == "https://":
            messagebox.showerror("Hiba", "Add meg az oldal URL-jét!"); return
        if not xlsx:
            messagebox.showerror("Hiba", "Add meg az Excel fájl mentési helyét!"); return
        if req_login and (not user or not pw):
            messagebox.showerror("Hiba", "Felhasználónév és jelszó szükséges!"); return
        if not xlsx.lower().endswith(".xlsx"):
            messagebox.showerror("Hiba", "Az Excel fájlnak .xlsx kiterjesztésűnek kell lennie!"); return

        try:
            sleep_s    = float(self._sleep.get().strip())
            save_every = int(self._save_ev.get().strip())
        except ValueError:
            messagebox.showerror("Hiba", "Várakozás és mentés értéknek számnak kell lennie!"); return

        codes = self._collect_codes()
        if not codes:
            messagebox.showerror("Hiba", "Adj meg cikkszámokat!"); return

        xlsx_folder = os.path.dirname(xlsx) or "."
        os.makedirs(xlsx_folder, exist_ok=True)
        csv_path = os.path.join(xlsx_folder, "cikkchecker_results.csv")

        if fresh and os.path.exists(csv_path):
            try: os.remove(csv_path)
            except: pass

        # Duplicate handling
        counter   = Counter(codes)
        dup_items = {c:n for c,n in counter.items() if n>1}
        unique    = list(dict.fromkeys(codes))

        if dup_items and not skip_dup:
            dup_rows = sum(n-1 for n in counter.values() if n>1)
            if self._dont_ask.get():
                if self._save_dup.get(): self._save_dup_txt(xlsx_folder, dup_items)
                if self._auto_rm.get():  codes = unique
            else:
                dlg = DuplicateDialog(self, dup_items, len(codes), len(unique), dup_rows)
                self.wait_window(dlg)
                if not dlg.ok: return
                if dlg.save:   self._save_dup_txt(xlsx_folder, dup_items)
                if dlg.remove: codes = unique

        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".txt",
                                          encoding="utf-8", delete=False)
        tmp.write("\n".join(codes)); tmp.close()

        self._stats = {"van":0,"nincs":0,"kulso":0,"ism":0,"total":len(set(codes))}
        self._s_total.set(len(set(codes)))
        self._s_done.set(0); self._s_van.set(0); self._s_nincs.set(0); self._s_kulso.set(0)
        self._progress.set(0)
        self._stop_req = False
        self._unknown_mode = None

        if not os.path.exists(BINARY_PATH):
            messagebox.showerror("Hiba",
                f"A Rust bináris nem található:\n{BINARY_PATH}\n\n"
                "Futtasd: cargo build --release\n"
                "és másold a checker_core[.exe] fájlt az app.py mellé.")
            return

        cmd = [BINARY_PATH,
               "--base-url", base_url,
               "--user", user, "--password", pw,
               "--codes-file", tmp.name,
               "--csv-output", csv_path,
               "--sleep-seconds", str(sleep_s),
               "--save-every", str(save_every),
               "--sort-order", self._get_sort_arg()]
        if req_login: cmd.append("--requires-login")

        self._log("═" * 42)
        self._log(f"Indítás — {len(codes)} cikkszám")
        self._set_status("Indítás...")
        self._can_continue = True
        self._update_start_btn()
        self._xlsx_path = xlsx
        self._csv_path  = csv_path

        self._proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                      stderr=subprocess.PIPE, text=True, bufsize=1)
        threading.Thread(target=self._read_proc, args=(tmp.name,), daemon=True).start()

    def _save_dup_txt(self, folder, dup_items):
        p = os.path.join(folder, "duplicates.txt")
        with open(p,"w",encoding="utf-8") as f:
            for c,n in dup_items.items(): f.write(f"{c};{n}\n")
        self._log(f"Duplikátumok mentve: {p}")

    def _read_proc(self, tmp_path):
        try:
            for line in self._proc.stdout:
                line = line.strip()
                if not line: continue
                try: msg = json.loads(line)
                except: self._log(line); continue

                kind = msg.get("kind","")

                if kind == "log":    self._log(msg["msg"])
                elif kind == "status": self._set_status(msg["msg"])
                elif kind == "progress":
                    self._set_progress(msg["current"], msg["total"])
                elif kind == "login_detected":
                    req = msg.get("required", False)
                    def _upd_login(r=req):
                        self._login_var.set(r)
                        self._toggle_login()
                        if r:
                            self._detect_lbl.configure(text="🔒  Bejelentkezés szükséges (auto)", text_color=C["amber"])
                        else:
                            self._detect_lbl.configure(text="🔓  Bejelentkezés nem szükséges (auto)", text_color=C["green"])
                    self.after(0, _upd_login)
                elif kind == "result":
                    avail = msg.get("elerhetoseg","")
                    price = msg.get("ar","")
                    code  = msg.get("cikkszam","")
                    if avail=="Van":           self._stats["van"]   += 1
                    elif avail=="Nincs":        self._stats["nincs"] += 1
                    elif avail=="Külső raktár": self._stats["kulso"] += 1
                    else:                       self._stats["ism"]   += 1
                    self.after(0, self._update_stat_cards)
                    self.after(0, lambda c=code, a=avail, p=price: self._table.add_row(c,a,p))
                elif kind == "unknown":
                    code = msg["cikkszam"]
                    cmd  = self._handle_unknown(code)
                    try: self._proc.stdin.write(cmd+"\n"); self._proc.stdin.flush()
                    except: pass
                elif kind == "error":
                    self._log(f"[HIBA] {msg['msg']}")
                elif kind == "done":
                    self._log("Folyamat kész — Excel exportálás...")
                    export_excel(self._csv_path, self._xlsx_path)
                    self._log(f"Excel mentve: {self._xlsx_path}")
                    self._set_status("Kész  ✓")
                    self._can_continue = False
                    self.after(0, self._update_start_btn)
        except Exception as e:
            self._log(f"Olvasási hiba: {e}")
        finally:
            try: os.unlink(tmp_path)
            except: pass

    def _handle_unknown(self, cikkszam):
        if self._unknown_mode == "skip_all": return "skip"
        if self._unknown_mode == "stop_all": return "stop"
        box = {"action":"stop","apply":False}
        ev  = threading.Event()
        def _show():
            dlg = UnknownDialog(self, cikkszam)
            self.wait_window(dlg)
            box["action"] = dlg.action; box["apply"] = dlg.apply_all; ev.set()
        self.after(0, _show); ev.wait()
        if box["apply"]:
            if box["action"]=="skip": self._unknown_mode="skip_all"
            else:                     self._unknown_mode="stop_all"
        return box["action"]

    # ── Update check ──────────────────────────────────────────────────────────
    def _check_update(self):
        try:
            r = requests.get(f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest", timeout=10)
            r.raise_for_status(); data = r.json()
        except: return

        latest = data.get("tag_name","").strip()
        notes  = data.get("body","").strip()

        def _v(s):
            s = s.strip().lstrip("v")
            return tuple(int(x) if x.isdigit() else 0 for x in s.split("."))

        if not latest or _v(latest) <= _v(APP_VERSION): return

        asset_url = next((a["browser_download_url"] for a in data.get("assets",[])
                          if a.get("name")==UPDATE_ASSET), None)
        if not asset_url: return

        self._log(f"Új verzió: {latest}")
        self.after(0, lambda: self._prompt_update(latest, notes, asset_url))

    def _prompt_update(self, latest, notes, asset_url):
        dlg = UpdateDialog(self, APP_VERSION, latest, notes)
        self.wait_window(dlg)
        if dlg.result == "update":
            threading.Thread(target=self._dl_update, args=(asset_url,latest), daemon=True).start()

    def _dl_update(self, url, latest):
        self._log(f"Letöltés: {latest}...")
        self._set_status(f"Frissítés letöltése ({latest})...")
        try:
            import tempfile as tf, subprocess as sp
            tmp = os.path.join(tf.gettempdir(), UPDATE_ASSET)
            with requests.get(url, stream=True, timeout=60) as r:
                r.raise_for_status()
                with open(tmp,"wb") as f:
                    for chunk in r.iter_content(8192):
                        if chunk: f.write(chunk)
            sp.Popen([tmp], shell=True)
            self.after(1000, self._quit)
        except Exception as e:
            self._log(f"Frissítési hiba: {e}")
            self._set_status("Frissítés sikertelen")


if __name__ == "__main__":
    App().mainloop()
