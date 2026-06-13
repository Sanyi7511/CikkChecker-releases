"""
CikkChecker v3 – Multi-language, theme-switching, settings panel
"""
import json, os, queue, subprocess, sys, tempfile, threading
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
CONFIG_PATH  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
ASSETS_DIR   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
LOGO_PATH    = os.path.join(ASSETS_DIR, "logo.png")
ICON_PATH    = os.path.join(ASSETS_DIR, "icon.ico")

# ══════════════════════════════════════════════════════════════════════════════
# TRANSLATIONS
# ══════════════════════════════════════════════════════════════════════════════
TRANSLATIONS = {
    "hu": {
        "_name": "Magyar",
        # Sidebar sections
        "s_connection":   "Kapcsolat",
        "s_login":        "Bejelentkezés",
        "s_files":        "Fájlok",
        "s_settings":     "Beállítások",
        "s_actions":      "Műveletek",
        # Connection
        "url_label":      "Weboldal URL",
        "login_required": "Bejelentkezés szükséges",
        "detecting":      "⟳  Detektálás...",
        "login_auto_yes": "🔒  Bejelentkezés szükséges (auto)",
        "login_auto_no":  "🔓  Nem szükséges (auto)",
        "login_check_fail":"⚠  Nem sikerült ellenőrizni",
        # Login
        "username":       "Felhasználónév",
        "password":       "Jelszó",
        # Files
        "codes_txt":      "Cikkszám TXT",
        "browse_txt":     "📂  TXT kiválasztása",
        "excel_output":   "Excel kimeneti fájl",
        "browse_xlsx":    "💾  Excel mentési helye",
        # Settings
        "wait_sec":       "Várakozás (mp)",
        "save_every":     "Mentés ennyinként",
        "sort_order":     "Sorrend",
        "sort_abc_asc":   "ABC (A→Z)",
        "sort_abc_desc":  "ABC (Z→A)",
        "sort_num_asc":   "Szám (0→9)",
        "sort_num_desc":  "Szám (9→0)",
        "sort_none":      "Eredeti sorrend",
        "auto_rm_dup":    "Auto duplikátum eltávolítás",
        "dont_ask_dup":   "Ne kérdezzen rá",
        "save_dup_txt":   "Duplikátumok TXT-be",
        # Actions
        "btn_start":      "▶  Indítás",
        "btn_continue":   "▶  Folytatás",
        "btn_stop":       "⏹  Stop",
        "btn_restart":    "↺  Újra",
        "btn_check_upd":  "🔄  Frissítés keresése",
        "btn_clear_log":  "🗑  Log törlése",
        # Main area
        "stat_total":     "Összes",
        "stat_done":      "Kész",
        "stat_van":       "Elérhető",
        "stat_nincs":     "Nem elérhető",
        "stat_kulso":     "Külső raktár",
        "status_waiting": "Várakozás",
        "progress_label": "0 / 0  (0%)",
        "results_title":  "Eredmények",
        "btn_clear_res":  "Törlés",
        "log_title":      "Napló",
        "manual_title":   "Kézi cikkszámok",
        # Settings dialog
        "settings_title":     "Beállítások",
        "section_appearance": "Megjelenés",
        "section_updates":    "Frissítések",
        "section_defaults":   "Alapértékek",
        "section_duplicates": "Duplikátumok",
        "section_language":   "Nyelv",
        "theme_label":        "Téma",
        "theme_dark":         "Sötét",
        "theme_light":        "Világos",
        "compact_log":        "Kompakt napló",
        "compact_log_desc":   "Kisebb betűméret a naplóban",
        "start_minimized":    "Minimalizálva indul",
        "auto_update":        "Automatikus frissítés",
        "update_interval":    "Ellenőrzés gyakorisága",
        "upd_on_start":       "Indításkor",
        "upd_daily":          "Naponta",
        "upd_never":          "Soha",
        "show_price":         "Ár lekérés",
        "show_price_desc":    "Árakat is letölti a keresés során",
        "auto_rm_label":      "Auto eltávolítás",
        "dont_ask_label":     "Ne kérdezzen rá",
        "save_dup_label":     "Mentés TXT-be",
        "language_label":     "Alkalmazás nyelve",
        "btn_save":           "Mentés",
        "btn_cancel":         "Mégse",
        # Update
        "upd_banner":         "Új verzió elérhető",
        "upd_btn_now":        "Frissítés",
        "upd_btn_later":      "Később",
        "upd_dialog_title":   "Frissítés elérhető",
        "upd_dialog_hdr":     "Új verzió",
        "upd_dialog_current": "jelenlegi",
        "upd_btn_install":    "⬇  Frissítés most",
        "upd_no_notes":       "Nincs kiadási megjegyzés.",
        "upd_downloading":    "Letöltés",
        "upd_done_export":    "Kész — Excel exportálás...",
        "upd_excel_saved":    "Excel mentve",
        "upd_status_ok":      "Kész  ✓",
        "upd_not_found":      "Nem érhető el újabb verzió.",
        "upd_failed":         "Frissítés sikertelen",
        "upd_error":          "Frissítési hiba",
        # Duplicate dialog
        "dup_title":          "Duplikátumok",
        "dup_hdr":            "Duplikátumok találva",
        "dup_total":          "Összes",
        "dup_unique":         "Egyedi",
        "dup_rows":           "Duplikált sorok",
        "dup_rm_opt":         "Minden cikkszámból csak 1 maradjon",
        "dup_save_opt":       "Duplikátumok mentése TXT fájlba",
        "btn_continue_dup":   "Folytatás",
        # Unknown dialog
        "unk_title":          "Ismeretlen találat",
        "unk_hdr":            "Nem értelmezhető",
        "unk_body":           "Mit tegyek a többi ismeretlen találattal?",
        "unk_apply_all":      "Alkalmazza a többi ismeretlen találatra is",
        "btn_skip":           "Kihagyás",
        "btn_stop_unk":       "Leállítás",
        # Errors / warnings
        "err_no_url":         "Add meg az oldal URL-jét!",
        "err_no_xlsx":        "Add meg az Excel fájl mentési helyét!",
        "err_no_creds":       "Felhasználónév és jelszó szükséges!",
        "err_xlsx_ext":       "Az Excel fájlnak .xlsx kiterjesztésűnek kell lennie!",
        "err_no_codes":       "Adj meg cikkszámokat!",
        "err_no_binary":      "checker_core nem található",
        "warn_already_running":"A folyamat már fut.",
        # Log messages
        "log_starting":       "Indítás",
        "log_codes":          "cikkszám",
        "log_stop_req":       "Leállítás kérve...",
        "log_restart_req":    "Újraindítás...",
        "log_dup_saved":      "Duplikátumok mentve",
        "log_read_err":       "Olvasási hiba",
        "status_starting":    "Indítás...",
    },

    "en": {
        "_name": "English",
        "s_connection":   "Connection",
        "s_login":        "Login",
        "s_files":        "Files",
        "s_settings":     "Settings",
        "s_actions":      "Actions",
        "url_label":      "Website URL",
        "login_required": "Login required",
        "detecting":      "⟳  Detecting...",
        "login_auto_yes": "🔒  Login required (auto)",
        "login_auto_no":  "🔓  No login needed (auto)",
        "login_check_fail":"⚠  Could not verify",
        "username":       "Username",
        "password":       "Password",
        "codes_txt":      "Article codes TXT",
        "browse_txt":     "📂  Browse TXT",
        "excel_output":   "Excel output file",
        "browse_xlsx":    "💾  Browse Excel location",
        "wait_sec":       "Wait between requests (s)",
        "save_every":     "Save Excel every N items",
        "sort_order":     "Sort order",
        "sort_abc_asc":   "ABC (A→Z)",
        "sort_abc_desc":  "ABC (Z→A)",
        "sort_num_asc":   "Number (0→9)",
        "sort_num_desc":  "Number (9→0)",
        "sort_none":      "Original order",
        "auto_rm_dup":    "Auto remove duplicates",
        "dont_ask_dup":   "Don't ask, apply automatically",
        "save_dup_txt":   "Save duplicates to TXT",
        "btn_start":      "▶  Start",
        "btn_continue":   "▶  Continue",
        "btn_stop":       "⏹  Stop",
        "btn_restart":    "↺  Restart",
        "btn_check_upd":  "🔄  Check for updates",
        "btn_clear_log":  "🗑  Clear log",
        "stat_total":     "Total",
        "stat_done":      "Done",
        "stat_van":       "In stock",
        "stat_nincs":     "Out of stock",
        "stat_kulso":     "External warehouse",
        "status_waiting": "Waiting",
        "progress_label": "0 / 0  (0%)",
        "results_title":  "Results",
        "btn_clear_res":  "Clear",
        "log_title":      "Log",
        "manual_title":   "Manual article codes",
        "settings_title":     "Settings",
        "section_appearance": "Appearance",
        "section_updates":    "Updates",
        "section_defaults":   "Defaults",
        "section_duplicates": "Duplicates",
        "section_language":   "Language",
        "theme_label":        "Theme",
        "theme_dark":         "Dark",
        "theme_light":        "Light",
        "compact_log":        "Compact log",
        "compact_log_desc":   "Smaller font size in log",
        "start_minimized":    "Start minimized",
        "auto_update":        "Automatic update check",
        "update_interval":    "Check frequency",
        "upd_on_start":       "On startup",
        "upd_daily":          "Daily",
        "upd_never":          "Never",
        "show_price":         "Fetch prices",
        "show_price_desc":    "Also downloads prices during search",
        "auto_rm_label":      "Auto remove",
        "dont_ask_label":     "Don't ask",
        "save_dup_label":     "Save to TXT",
        "language_label":     "Application language",
        "btn_save":           "Save",
        "btn_cancel":         "Cancel",
        "upd_banner":         "New version available",
        "upd_btn_now":        "Update",
        "upd_btn_later":      "Later",
        "upd_dialog_title":   "Update available",
        "upd_dialog_hdr":     "New version",
        "upd_dialog_current": "current",
        "upd_btn_install":    "⬇  Update now",
        "upd_no_notes":       "No release notes.",
        "upd_downloading":    "Downloading",
        "upd_done_export":    "Done — exporting Excel...",
        "upd_excel_saved":    "Excel saved",
        "upd_status_ok":      "Done  ✓",
        "upd_not_found":      "No newer version available.",
        "upd_failed":         "Update failed",
        "upd_error":          "Update error",
        "dup_title":          "Duplicates",
        "dup_hdr":            "Duplicates found",
        "dup_total":          "Total",
        "dup_unique":         "Unique",
        "dup_rows":           "Duplicate rows",
        "dup_rm_opt":         "Keep only 1 of each article code",
        "dup_save_opt":       "Save duplicates to TXT file",
        "btn_continue_dup":   "Continue",
        "unk_title":          "Unknown result",
        "unk_hdr":            "Unrecognized",
        "unk_body":           "What should I do with other unknown results?",
        "unk_apply_all":      "Apply to all other unknown results",
        "btn_skip":           "Skip",
        "btn_stop_unk":       "Stop",
        "err_no_url":         "Please enter the website URL!",
        "err_no_xlsx":        "Please enter the Excel output path!",
        "err_no_creds":       "Username and password are required!",
        "err_xlsx_ext":       "Excel file must have .xlsx extension!",
        "err_no_codes":       "Please enter article codes!",
        "err_no_binary":      "checker_core not found",
        "warn_already_running":"Process is already running.",
        "log_starting":       "Starting",
        "log_codes":          "article codes",
        "log_stop_req":       "Stop requested...",
        "log_restart_req":    "Restarting...",
        "log_dup_saved":      "Duplicates saved",
        "log_read_err":       "Read error",
        "status_starting":    "Starting...",
    },

    "de": {
        "_name": "Deutsch",
        "s_connection":   "Verbindung",
        "s_login":        "Anmeldung",
        "s_files":        "Dateien",
        "s_settings":     "Einstellungen",
        "s_actions":      "Aktionen",
        "url_label":      "Website URL",
        "login_required": "Anmeldung erforderlich",
        "detecting":      "⟳  Erkennung...",
        "login_auto_yes": "🔒  Anmeldung erforderlich (auto)",
        "login_auto_no":  "🔓  Keine Anmeldung nötig (auto)",
        "login_check_fail":"⚠  Konnte nicht prüfen",
        "username":       "Benutzername",
        "password":       "Passwort",
        "codes_txt":      "Artikelnummern TXT",
        "browse_txt":     "📂  TXT auswählen",
        "excel_output":   "Excel-Ausgabedatei",
        "browse_xlsx":    "💾  Excel-Speicherort",
        "wait_sec":       "Wartezeit (Sek.)",
        "save_every":     "Excel speichern alle N",
        "sort_order":     "Sortierung",
        "sort_abc_asc":   "ABC (A→Z)",
        "sort_abc_desc":  "ABC (Z→A)",
        "sort_num_asc":   "Nummer (0→9)",
        "sort_num_desc":  "Nummer (9→0)",
        "sort_none":      "Ursprüngliche Reihenfolge",
        "auto_rm_dup":    "Duplikate auto. entfernen",
        "dont_ask_dup":   "Nicht fragen, auto. anwenden",
        "save_dup_txt":   "Duplikate als TXT speichern",
        "btn_start":      "▶  Starten",
        "btn_continue":   "▶  Fortsetzen",
        "btn_stop":       "⏹  Stopp",
        "btn_restart":    "↺  Neustart",
        "btn_check_upd":  "🔄  Auf Updates prüfen",
        "btn_clear_log":  "🗑  Log leeren",
        "stat_total":     "Gesamt",
        "stat_done":      "Fertig",
        "stat_van":       "Verfügbar",
        "stat_nincs":     "Nicht verfügbar",
        "stat_kulso":     "Externes Lager",
        "status_waiting": "Warten",
        "progress_label": "0 / 0  (0%)",
        "results_title":  "Ergebnisse",
        "btn_clear_res":  "Löschen",
        "log_title":      "Protokoll",
        "manual_title":   "Manuelle Artikelnummern",
        "settings_title":     "Einstellungen",
        "section_appearance": "Darstellung",
        "section_updates":    "Updates",
        "section_defaults":   "Standardwerte",
        "section_duplicates": "Duplikate",
        "section_language":   "Sprache",
        "theme_label":        "Design",
        "theme_dark":         "Dunkel",
        "theme_light":        "Hell",
        "compact_log":        "Kompaktes Protokoll",
        "compact_log_desc":   "Kleinere Schriftgröße im Protokoll",
        "start_minimized":    "Minimiert starten",
        "auto_update":        "Automatische Update-Prüfung",
        "update_interval":    "Prüfhäufigkeit",
        "upd_on_start":       "Beim Start",
        "upd_daily":          "Täglich",
        "upd_never":          "Nie",
        "show_price":         "Preise abrufen",
        "show_price_desc":    "Lädt auch Preise herunter",
        "auto_rm_label":      "Auto entfernen",
        "dont_ask_label":     "Nicht fragen",
        "save_dup_label":     "Als TXT speichern",
        "language_label":     "Anwendungssprache",
        "btn_save":           "Speichern",
        "btn_cancel":         "Abbrechen",
        "upd_banner":         "Neue Version verfügbar",
        "upd_btn_now":        "Update",
        "upd_btn_later":      "Später",
        "upd_dialog_title":   "Update verfügbar",
        "upd_dialog_hdr":     "Neue Version",
        "upd_dialog_current": "aktuell",
        "upd_btn_install":    "⬇  Jetzt aktualisieren",
        "upd_no_notes":       "Keine Versionshinweise.",
        "upd_downloading":    "Herunterladen",
        "upd_done_export":    "Fertig — Excel exportieren...",
        "upd_excel_saved":    "Excel gespeichert",
        "upd_status_ok":      "Fertig  ✓",
        "upd_not_found":      "Keine neuere Version verfügbar.",
        "upd_failed":         "Update fehlgeschlagen",
        "upd_error":          "Update-Fehler",
        "dup_title":          "Duplikate",
        "dup_hdr":            "Duplikate gefunden",
        "dup_total":          "Gesamt",
        "dup_unique":         "Eindeutig",
        "dup_rows":           "Doppelte Zeilen",
        "dup_rm_opt":         "Von jeder Artikelnummer nur 1 behalten",
        "dup_save_opt":       "Duplikate als TXT speichern",
        "btn_continue_dup":   "Weiter",
        "unk_title":          "Unbekanntes Ergebnis",
        "unk_hdr":            "Nicht erkannt",
        "unk_body":           "Was soll ich mit anderen unbekannten Ergebnissen tun?",
        "unk_apply_all":      "Auf alle anderen unbekannten Ergebnisse anwenden",
        "btn_skip":           "Überspringen",
        "btn_stop_unk":       "Stoppen",
        "err_no_url":         "Bitte Website-URL eingeben!",
        "err_no_xlsx":        "Bitte Excel-Ausgabepfad eingeben!",
        "err_no_creds":       "Benutzername und Passwort erforderlich!",
        "err_xlsx_ext":       "Excel-Datei muss .xlsx Erweiterung haben!",
        "err_no_codes":       "Bitte Artikelnummern eingeben!",
        "err_no_binary":      "checker_core nicht gefunden",
        "warn_already_running":"Prozess läuft bereits.",
        "log_starting":       "Starten",
        "log_codes":          "Artikelnummern",
        "log_stop_req":       "Stopp angefordert...",
        "log_restart_req":    "Neustart...",
        "log_dup_saved":      "Duplikate gespeichert",
        "log_read_err":       "Lesefehler",
        "status_starting":    "Starte...",
    },

    "ro": {
        "_name": "Română",
        "s_connection":   "Conexiune",
        "s_login":        "Autentificare",
        "s_files":        "Fișiere",
        "s_settings":     "Setări",
        "s_actions":      "Acțiuni",
        "url_label":      "URL website",
        "login_required": "Autentificare necesară",
        "detecting":      "⟳  Detectare...",
        "login_auto_yes": "🔒  Autentificare necesară (auto)",
        "login_auto_no":  "🔓  Fără autentificare (auto)",
        "login_check_fail":"⚠  Verificare eșuată",
        "username":       "Utilizator",
        "password":       "Parolă",
        "codes_txt":      "Coduri articole TXT",
        "browse_txt":     "📂  Selectare TXT",
        "excel_output":   "Fișier Excel ieșire",
        "browse_xlsx":    "💾  Locație Excel",
        "wait_sec":       "Așteptare (sec.)",
        "save_every":     "Salvare Excel la N",
        "sort_order":     "Ordine",
        "sort_abc_asc":   "ABC (A→Z)",
        "sort_abc_desc":  "ABC (Z→A)",
        "sort_num_asc":   "Număr (0→9)",
        "sort_num_desc":  "Număr (9→0)",
        "sort_none":      "Ordine originală",
        "auto_rm_dup":    "Eliminare auto duplicate",
        "dont_ask_dup":   "Nu întreba, aplică automat",
        "save_dup_txt":   "Salvare duplicate TXT",
        "btn_start":      "▶  Start",
        "btn_continue":   "▶  Continuă",
        "btn_stop":       "⏹  Stop",
        "btn_restart":    "↺  Repornire",
        "btn_check_upd":  "🔄  Verificare actualizări",
        "btn_clear_log":  "🗑  Șterge log",
        "stat_total":     "Total",
        "stat_done":      "Finalizat",
        "stat_van":       "Disponibil",
        "stat_nincs":     "Indisponibil",
        "stat_kulso":     "Depozit extern",
        "status_waiting": "Așteptare",
        "progress_label": "0 / 0  (0%)",
        "results_title":  "Rezultate",
        "btn_clear_res":  "Șterge",
        "log_title":      "Jurnal",
        "manual_title":   "Coduri manuale",
        "settings_title":     "Setări",
        "section_appearance": "Aspect",
        "section_updates":    "Actualizări",
        "section_defaults":   "Valori implicite",
        "section_duplicates": "Duplicate",
        "section_language":   "Limbă",
        "theme_label":        "Temă",
        "theme_dark":         "Întunecat",
        "theme_light":        "Luminos",
        "compact_log":        "Jurnal compact",
        "compact_log_desc":   "Font mai mic în jurnal",
        "start_minimized":    "Pornire minimizată",
        "auto_update":        "Verificare automată actualizări",
        "update_interval":    "Frecvența verificării",
        "upd_on_start":       "La pornire",
        "upd_daily":          "Zilnic",
        "upd_never":          "Niciodată",
        "show_price":         "Descărcare prețuri",
        "show_price_desc":    "Descarcă și prețurile",
        "auto_rm_label":      "Eliminare auto",
        "dont_ask_label":     "Nu întreba",
        "save_dup_label":     "Salvare TXT",
        "language_label":     "Limba aplicației",
        "btn_save":           "Salvare",
        "btn_cancel":         "Anulare",
        "upd_banner":         "Versiune nouă disponibilă",
        "upd_btn_now":        "Actualizare",
        "upd_btn_later":      "Mai târziu",
        "upd_dialog_title":   "Actualizare disponibilă",
        "upd_dialog_hdr":     "Versiune nouă",
        "upd_dialog_current": "curent",
        "upd_btn_install":    "⬇  Actualizează acum",
        "upd_no_notes":       "Fără note de versiune.",
        "upd_downloading":    "Descărcare",
        "upd_done_export":    "Gata — export Excel...",
        "upd_excel_saved":    "Excel salvat",
        "upd_status_ok":      "Gata  ✓",
        "upd_not_found":      "Nicio versiune mai nouă.",
        "upd_failed":         "Actualizare eșuată",
        "upd_error":          "Eroare actualizare",
        "dup_title":          "Duplicate",
        "dup_hdr":            "Duplicate găsite",
        "dup_total":          "Total",
        "dup_unique":         "Unic",
        "dup_rows":           "Rânduri duplicate",
        "dup_rm_opt":         "Păstrează doar 1 din fiecare cod",
        "dup_save_opt":       "Salvare duplicate în TXT",
        "btn_continue_dup":   "Continuă",
        "unk_title":          "Rezultat necunoscut",
        "unk_hdr":            "Nerecunoscut",
        "unk_body":           "Ce fac cu alte rezultate necunoscute?",
        "unk_apply_all":      "Aplică la toate rezultatele necunoscute",
        "btn_skip":           "Omite",
        "btn_stop_unk":       "Oprire",
        "err_no_url":         "Introduceți URL-ul website-ului!",
        "err_no_xlsx":        "Introduceți calea fișierului Excel!",
        "err_no_creds":       "Utilizator și parolă necesare!",
        "err_xlsx_ext":       "Fișierul Excel trebuie să aibă extensia .xlsx!",
        "err_no_codes":       "Introduceți coduri de articole!",
        "err_no_binary":      "checker_core negăsit",
        "warn_already_running":"Procesul rulează deja.",
        "log_starting":       "Pornire",
        "log_codes":          "coduri articole",
        "log_stop_req":       "Oprire solicitată...",
        "log_restart_req":    "Repornire...",
        "log_dup_saved":      "Duplicate salvate",
        "log_read_err":       "Eroare citire",
        "status_starting":    "Pornire...",
    },

    "sk": {
        "_name": "Slovenčina",
        "s_connection":   "Pripojenie",
        "s_login":        "Prihlásenie",
        "s_files":        "Súbory",
        "s_settings":     "Nastavenia",
        "s_actions":      "Akcie",
        "url_label":      "URL webovej stránky",
        "login_required": "Prihlásenie je potrebné",
        "detecting":      "⟳  Zisťovanie...",
        "login_auto_yes": "🔒  Prihlásenie potrebné (auto)",
        "login_auto_no":  "🔓  Prihlásenie nepotrebné (auto)",
        "login_check_fail":"⚠  Overenie zlyhalo",
        "username":       "Používateľské meno",
        "password":       "Heslo",
        "codes_txt":      "Kódy článkov TXT",
        "browse_txt":     "📂  Vybrať TXT",
        "excel_output":   "Excel výstupný súbor",
        "browse_xlsx":    "💾  Uložiť Excel",
        "wait_sec":       "Čakanie (sek.)",
        "save_every":     "Uložiť Excel každých N",
        "sort_order":     "Poradie",
        "sort_abc_asc":   "ABC (A→Z)",
        "sort_abc_desc":  "ABC (Z→A)",
        "sort_num_asc":   "Číslo (0→9)",
        "sort_num_desc":  "Číslo (9→0)",
        "sort_none":      "Pôvodné poradie",
        "auto_rm_dup":    "Auto odstránenie duplikátov",
        "dont_ask_dup":   "Nepýtaj sa, aplikuj auto",
        "save_dup_txt":   "Uložiť duplikáty do TXT",
        "btn_start":      "▶  Spustiť",
        "btn_continue":   "▶  Pokračovať",
        "btn_stop":       "⏹  Zastaviť",
        "btn_restart":    "↺  Reštart",
        "btn_check_upd":  "🔄  Skontrolovať aktualizácie",
        "btn_clear_log":  "🗑  Vymazať log",
        "stat_total":     "Celkom",
        "stat_done":      "Hotovo",
        "stat_van":       "Dostupné",
        "stat_nincs":     "Nedostupné",
        "stat_kulso":     "Externý sklad",
        "status_waiting": "Čakanie",
        "progress_label": "0 / 0  (0%)",
        "results_title":  "Výsledky",
        "btn_clear_res":  "Vymazať",
        "log_title":      "Protokol",
        "manual_title":   "Manuálne kódy",
        "settings_title":     "Nastavenia",
        "section_appearance": "Vzhľad",
        "section_updates":    "Aktualizácie",
        "section_defaults":   "Predvolené",
        "section_duplicates": "Duplikáty",
        "section_language":   "Jazyk",
        "theme_label":        "Téma",
        "theme_dark":         "Tmavá",
        "theme_light":        "Svetlá",
        "compact_log":        "Kompaktný protokol",
        "compact_log_desc":   "Menší font v protokole",
        "start_minimized":    "Spustiť minimalizované",
        "auto_update":        "Automatická kontrola aktualizácií",
        "update_interval":    "Frekvencia kontroly",
        "upd_on_start":       "Pri spustení",
        "upd_daily":          "Denne",
        "upd_never":          "Nikdy",
        "show_price":         "Načítať ceny",
        "show_price_desc":    "Stiahne aj ceny",
        "auto_rm_label":      "Auto odstrániť",
        "dont_ask_label":     "Nepýtaj sa",
        "save_dup_label":     "Uložiť do TXT",
        "language_label":     "Jazyk aplikácie",
        "btn_save":           "Uložiť",
        "btn_cancel":         "Zrušiť",
        "upd_banner":         "Dostupná nová verzia",
        "upd_btn_now":        "Aktualizovať",
        "upd_btn_later":      "Neskôr",
        "upd_dialog_title":   "Dostupná aktualizácia",
        "upd_dialog_hdr":     "Nová verzia",
        "upd_dialog_current": "aktuálna",
        "upd_btn_install":    "⬇  Aktualizovať teraz",
        "upd_no_notes":       "Žiadne poznámky k vydaniu.",
        "upd_downloading":    "Sťahovanie",
        "upd_done_export":    "Hotovo — export Excel...",
        "upd_excel_saved":    "Excel uložený",
        "upd_status_ok":      "Hotovo  ✓",
        "upd_not_found":      "Žiadna novšia verzia.",
        "upd_failed":         "Aktualizácia zlyhala",
        "upd_error":          "Chyba aktualizácie",
        "dup_title":          "Duplikáty",
        "dup_hdr":            "Nájdené duplikáty",
        "dup_total":          "Celkom",
        "dup_unique":         "Jedinečné",
        "dup_rows":           "Duplicitné riadky",
        "dup_rm_opt":         "Ponechať iba 1 z každého kódu",
        "dup_save_opt":       "Uložiť duplikáty do TXT",
        "btn_continue_dup":   "Pokračovať",
        "unk_title":          "Neznámy výsledok",
        "unk_hdr":            "Nerozpoznané",
        "unk_body":           "Čo robiť s ďalšími neznámymi výsledkami?",
        "unk_apply_all":      "Aplikovať na všetky neznáme výsledky",
        "btn_skip":           "Preskočiť",
        "btn_stop_unk":       "Zastaviť",
        "err_no_url":         "Zadajte URL webovej stránky!",
        "err_no_xlsx":        "Zadajte cestu k Excel súboru!",
        "err_no_creds":       "Vyžaduje sa používateľské meno a heslo!",
        "err_xlsx_ext":       "Excel súbor musí mať príponu .xlsx!",
        "err_no_codes":       "Zadajte kódy článkov!",
        "err_no_binary":      "checker_core nenájdený",
        "warn_already_running":"Proces už beží.",
        "log_starting":       "Spúšťam",
        "log_codes":          "kódov článkov",
        "log_stop_req":       "Zastavenie požiadané...",
        "log_restart_req":    "Reštart...",
        "log_dup_saved":      "Duplikáty uložené",
        "log_read_err":       "Chyba čítania",
        "status_starting":    "Spúšťam...",
    },
}

LANG_NAMES  = {k: v["_name"] for k, v in TRANSLATIONS.items()}
LANG_CODES  = {v["_name"]: k for k, v in TRANSLATIONS.items()}

_current_lang = "hu"

def t(key: str) -> str:
    lang = TRANSLATIONS.get(_current_lang, TRANSLATIONS["hu"])
    return lang.get(key, TRANSLATIONS["hu"].get(key, key))

def set_language(lang_code: str):
    global _current_lang
    if lang_code in TRANSLATIONS:
        _current_lang = lang_code

# ══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════════════════
DEFAULT_CONFIG = {
    "theme": "dark", "language": "hu",
    "auto_update": True, "update_interval": "on_start",
    "sleep_seconds": 0.8, "save_every": 100,
    "sort_order": "abc", "show_price": True,
    "compact_log": False, "start_minimized": False,
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

def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except: pass

# ══════════════════════════════════════════════════════════════════════════════
# THEMES
# ══════════════════════════════════════════════════════════════════════════════
THEMES = {
    "dark": {
        "bg":"#08090E","sidebar":"#0D1018","card":"#13161F","card2":"#181C28",
        "border":"#1F2535","border2":"#2A3147","accent":"#5B8DEF","accent2":"#3B6DD4",
        "green":"#23C55E","green2":"#16A34A","green_bg":"#0A2016",
        "red":"#F04747","red2":"#C53030","red_bg":"#200A0A",
        "amber":"#F5A623","amber2":"#D97706","amber_bg":"#201508",
        "text":"#E8EBF0","text2":"#9BA3B4","text3":"#4A5568",
    },
    "light": {
        "bg":"#F0F2F7","sidebar":"#E4E8F2","card":"#FFFFFF","card2":"#F7F8FC",
        "border":"#D1D9E8","border2":"#B0BCDA","accent":"#3B6DD4","accent2":"#2754B8",
        "green":"#16A34A","green2":"#15803D","green_bg":"#DCFCE7",
        "red":"#DC2626","red2":"#B91C1C","red_bg":"#FEE2E2",
        "amber":"#D97706","amber2":"#B45309","amber_bg":"#FEF3C7",
        "text":"#111827","text2":"#374151","text3":"#9CA3AF",
    },
}
C = THEMES["dark"]
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ══════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════════════════════
def export_excel(csv_path, excel_path):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        wb = Workbook(); ws = wb.active; ws.title = "CikkChecker"
        fills = {
            "Van":          PatternFill("solid", fgColor="0A2016"),
            "Nincs":        PatternFill("solid", fgColor="200A0A"),
            "Külső raktár": PatternFill("solid", fgColor="201508"),
            "In stock":     PatternFill("solid", fgColor="0A2016"),
            "Out of stock": PatternFill("solid", fgColor="200A0A"),
        }
        fonts = {
            "Van":          Font(color="23C55E", bold=True, size=10),
            "Nincs":        Font(color="F04747", bold=True, size=10),
            "Külső raktár": Font(color="F5A623", bold=True, size=10),
            "In stock":     Font(color="23C55E", bold=True, size=10),
            "Out of stock": Font(color="F04747", bold=True, size=10),
        }
        thin = Side(style="thin", color="1F2535")
        brd  = Border(left=thin,right=thin,top=thin,bottom=thin)
        ws.append(["Cikkszám","Elérhetőség","Ár"])
        for cell in ws[1]:
            cell.fill=PatternFill("solid",fgColor="0D1018")
            cell.font=Font(color="9BA3B4",bold=True,size=10)
            cell.alignment=Alignment(horizontal="center"); cell.border=brd
        import csv as _csv
        if os.path.exists(csv_path):
            with open(csv_path,encoding="utf-8-sig",newline="") as f:
                for row in _csv.DictReader(f):
                    c=(row.get("Cikkszam") or row.get("Cikkszám") or "").strip()
                    a=(row.get("Elerhetoseg") or row.get("Elérhetőség") or "").strip()
                    p=(row.get("Ar") or row.get("Ár") or "").strip()
                    if not c: continue
                    ws.append([c,a,p]); r=ws.max_row
                    for col in range(1,4):
                        cell=ws.cell(r,col); cell.border=brd
                        cell.alignment=Alignment(horizontal="left" if col==1 else "center")
                        cell.fill=fills.get(a,PatternFill("solid",fgColor="13161F"))
                    ws.cell(r,2).font=fonts.get(a,Font(color="9BA3B4",size=10))
                    ws.cell(r,1).font=Font(color="E8EBF0",size=10)
                    ws.cell(r,3).font=Font(color="9BA3B4",size=10)
        ws.column_dimensions["A"].width=28
        ws.column_dimensions["B"].width=18
        ws.column_dimensions["C"].width=16
        wb.save(excel_path)
    except Exception as e:
        print(f"Excel export error: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# BASE DIALOG
# ══════════════════════════════════════════════════════════════════════════════
class D(ctk.CTkToplevel):
    def __init__(self, parent, title, w=560, h=400):
        super().__init__(parent)
        self.title(title); self.geometry(f"{w}x{h}")
        self.configure(fg_color=C["bg"])
        self.transient(parent); self.grab_set()
        self.grid_columnconfigure(0, weight=1)
        self.wait_visibility(); self.focus()

    def _hdr(self, icon, text, color, row=0):
        f=ctk.CTkFrame(self,fg_color=C["card"],corner_radius=0)
        f.grid(row=row,column=0,sticky="ew")
        f.grid_columnconfigure(1,weight=1)
        ctk.CTkLabel(f,text=icon,font=ctk.CTkFont(size=20),text_color=color
                     ).grid(row=0,column=0,padx=(20,8),pady=16)
        ctk.CTkLabel(f,text=text,font=ctk.CTkFont(size=15,weight="bold"),
                     text_color=C["text"]).grid(row=0,column=1,pady=16,sticky="w")

    def _btns(self, row, items):
        f=ctk.CTkFrame(self,fg_color="transparent")
        f.grid(row=row,column=0,padx=20,pady=(0,20),sticky="ew")
        for i in range(len(items)): f.grid_columnconfigure(i,weight=1)
        for i,(txt,cmd,fg,hov) in enumerate(items):
            ctk.CTkButton(f,text=txt,command=cmd,height=40,corner_radius=8,
                          fg_color=fg,hover_color=hov,
                          font=ctk.CTkFont(size=13,weight="bold")
                          ).grid(row=0,column=i,padx=6,sticky="ew")

# ══════════════════════════════════════════════════════════════════════════════
# SETTINGS DIALOG
# ══════════════════════════════════════════════════════════════════════════════
class SettingsDialog(D):
    def __init__(self, parent, config, on_save):
        super().__init__(parent, t("settings_title"), 580, 660)
        self._cfg=dict(config); self._on_save=on_save
        self.minsize(500,560)
        self.grid_rowconfigure(1,weight=1)
        self._hdr("⚙", t("settings_title"), C["accent"])

        sc=ctk.CTkScrollableFrame(self,fg_color="transparent")
        sc.grid(row=1,column=0,padx=16,pady=(12,8),sticky="nsew")
        sc.grid_columnconfigure(0,weight=1)

        def sec(text,row):
            ctk.CTkLabel(sc,text=text.upper(),font=ctk.CTkFont(size=9,weight="bold"),
                         text_color=C["text3"]).grid(row=row,column=0,padx=4,pady=(16,4),sticky="w")

        def card(row):
            f=ctk.CTkFrame(sc,fg_color=C["card"],corner_radius=10,
                           border_width=1,border_color=C["border"])
            f.grid(row=row,column=0,sticky="ew",pady=(0,4))
            f.grid_columnconfigure(0,weight=1); return f

        def row_sw(parent,label,var,row,desc=None):
            f=ctk.CTkFrame(parent,fg_color="transparent")
            f.grid(row=row,column=0,padx=14,pady=(10 if row==0 else 4, 2),sticky="ew")
            f.grid_columnconfigure(0,weight=1)
            ctk.CTkLabel(f,text=label,font=ctk.CTkFont(size=12),
                         text_color=C["text"]).grid(row=0,column=0,sticky="w")
            ctk.CTkSwitch(f,variable=var,text="",
                          fg_color=C["border2"],progress_color=C["accent"]
                          ).grid(row=0,column=1,sticky="e",padx=(8,0))
            if desc:
                ctk.CTkLabel(parent,text=desc,font=ctk.CTkFont(size=10),
                             text_color=C["text3"]
                             ).grid(row=row+1,column=0,padx=14,pady=(0,6),sticky="w")

        def row_menu(parent,label,var,values,row,w=150):
            f=ctk.CTkFrame(parent,fg_color="transparent")
            f.grid(row=row,column=0,padx=14,pady=(10 if row==0 else 4,4),sticky="ew")
            f.grid_columnconfigure(0,weight=1)
            ctk.CTkLabel(f,text=label,font=ctk.CTkFont(size=12),
                         text_color=C["text"]).grid(row=0,column=0,sticky="w")
            ctk.CTkOptionMenu(f,values=values,variable=var,width=w,
                              fg_color=C["card2"],button_color=C["border2"],
                              button_hover_color=C["accent2"],text_color=C["text"],
                              dropdown_fg_color=C["card"],dropdown_hover_color=C["border2"],
                              font=ctk.CTkFont(size=12)
                              ).grid(row=0,column=1,sticky="e",padx=(8,0))

        def row_entry(parent,label,var,row):
            f=ctk.CTkFrame(parent,fg_color="transparent")
            f.grid(row=row,column=0,padx=14,pady=(10 if row==0 else 4,4),sticky="ew")
            f.grid_columnconfigure(0,weight=1)
            ctk.CTkLabel(f,text=label,font=ctk.CTkFont(size=12),
                         text_color=C["text"]).grid(row=0,column=0,sticky="w")
            ctk.CTkEntry(f,textvariable=var,width=100,fg_color=C["card2"],
                         border_color=C["border2"],text_color=C["text"],
                         corner_radius=7,height=32
                         ).grid(row=0,column=1,sticky="e",padx=(8,0))

        # ── Language ──────────────────────────────────────────────────────
        sec(t("section_language"), 0)
        c0=card(1)
        lang_names=[v["_name"] for v in TRANSLATIONS.values()]
        cur_lang_name=TRANSLATIONS.get(self._cfg.get("language","hu"),TRANSLATIONS["hu"])["_name"]
        self._lang_var=ctk.StringVar(value=cur_lang_name)
        row_menu(c0,t("language_label"),self._lang_var,lang_names,0,w=160)

        # ── Appearance ────────────────────────────────────────────────────
        sec(t("section_appearance"), 2)
        c1=card(3)
        self._theme_var=ctk.StringVar(
            value=t("theme_dark") if self._cfg["theme"]=="dark" else t("theme_light"))
        row_menu(c1,t("theme_label"),self._theme_var,[t("theme_dark"),t("theme_light")],0)
        self._compact=ctk.BooleanVar(value=self._cfg.get("compact_log",False))
        row_sw(c1,t("compact_log"),self._compact,2,desc=t("compact_log_desc"))
        self._minimized=ctk.BooleanVar(value=self._cfg.get("start_minimized",False))
        row_sw(c1,t("start_minimized"),self._minimized,4)

        # ── Updates ───────────────────────────────────────────────────────
        sec(t("section_updates"), 4)
        c2=card(5)
        self._auto_upd=ctk.BooleanVar(value=self._cfg.get("auto_update",True))
        row_sw(c2,t("auto_update"),self._auto_upd,0)
        interval_map={
            "on_start": t("upd_on_start"),
            "daily":    t("upd_daily"),
            "never":    t("upd_never"),
        }
        rev_interval={v:k for k,v in interval_map.items()}
        cur_interval=interval_map.get(self._cfg.get("update_interval","on_start"),t("upd_on_start"))
        self._interval_var=ctk.StringVar(value=cur_interval)
        row_menu(c2,t("update_interval"),self._interval_var,
                 [t("upd_on_start"),t("upd_daily"),t("upd_never")],2)
        self._rev_interval=rev_interval

        # ── Defaults ──────────────────────────────────────────────────────
        sec(t("section_defaults"), 6)
        c3=card(7)
        self._sleep_var=tk.StringVar(value=str(self._cfg.get("sleep_seconds",0.8)))
        self._savev_var=tk.StringVar(value=str(self._cfg.get("save_every",100)))
        row_entry(c3,t("wait_sec"),self._sleep_var,0)
        row_entry(c3,t("save_every"),self._savev_var,2)

        sort_keys=["abc","abc_desc","num","num_desc","none"]
        sort_labels=[t(f"sort_abc_asc"),t(f"sort_abc_desc"),
                     t(f"sort_num_asc"),t(f"sort_num_desc"),t(f"sort_none")]
        sort_map=dict(zip(sort_keys,sort_labels))
        rev_sort={v:k for k,v in sort_map.items()}
        cur_sort=sort_map.get(self._cfg.get("sort_order","abc"),t("sort_abc_asc"))
        self._sort_var=ctk.StringVar(value=cur_sort)
        self._rev_sort=rev_sort
        row_menu(c3,t("sort_order"),self._sort_var,sort_labels,4,w=180)
        self._price=ctk.BooleanVar(value=self._cfg.get("show_price",True))
        row_sw(c3,t("show_price"),self._price,6,desc=t("show_price_desc"))

        # ── Duplicates ────────────────────────────────────────────────────
        sec(t("section_duplicates"), 8)
        c4=card(9)
        self._dup_rm  =ctk.BooleanVar(value=self._cfg.get("auto_rm_dup",False))
        self._dup_ask =ctk.BooleanVar(value=self._cfg.get("dont_ask_dup",False))
        self._dup_save=ctk.BooleanVar(value=self._cfg.get("save_dup_txt",True))
        row_sw(c4,t("auto_rm_label"),self._dup_rm,0)
        row_sw(c4,t("dont_ask_label"),self._dup_ask,2)
        row_sw(c4,t("save_dup_label"),self._dup_save,4)

        self._btns(2,[
            (t("btn_save"),  self._save,    C["accent2"], C["accent"]),
            (t("btn_cancel"),self.destroy,  C["card"],    C["border2"]),
        ])

    def _save(self):
        try: sl=float(self._sleep_var.get())
        except: sl=0.8
        try: sv=int(self._savev_var.get())
        except: sv=100
        lang_name=self._lang_var.get()
        lang_code=LANG_CODES.get(lang_name,"hu")
        theme_dark_label=TRANSLATIONS.get(lang_code,TRANSLATIONS["hu"]).get("theme_dark","Sötét")
        self._cfg.update({
            "language":       lang_code,
            "theme":          "dark" if self._theme_var.get()==theme_dark_label else "light",
            "compact_log":    self._compact.get(),
            "start_minimized":self._minimized.get(),
            "auto_update":    self._auto_upd.get(),
            "update_interval":self._rev_interval.get(self._interval_var.get(),"on_start"),
            "sleep_seconds":  sl,
            "save_every":     sv,
            "sort_order":     self._rev_sort.get(self._sort_var.get(),"abc"),
            "show_price":     self._price.get(),
            "auto_rm_dup":    self._dup_rm.get(),
            "dont_ask_dup":   self._dup_ask.get(),
            "save_dup_txt":   self._dup_save.get(),
        })
        self._on_save(self._cfg)
        self.destroy()

# ══════════════════════════════════════════════════════════════════════════════
# OTHER DIALOGS
# ══════════════════════════════════════════════════════════════════════════════
class UpdateBanner(ctk.CTkFrame):
    def __init__(self, parent, latest, on_update, on_dismiss):
        super().__init__(parent,fg_color=C["card"],corner_radius=10,
                         border_width=1,border_color=C["accent2"])
        self.grid_columnconfigure(1,weight=1)
        ctk.CTkLabel(self,text="🚀",font=ctk.CTkFont(size=16)
                     ).grid(row=0,column=0,padx=(14,6),pady=10)
        ctk.CTkLabel(self,text=f"{t('upd_banner')}: {latest}",
                     font=ctk.CTkFont(size=13,weight="bold"),
                     text_color=C["accent"]).grid(row=0,column=1,sticky="w",pady=10)
        bf=ctk.CTkFrame(self,fg_color="transparent")
        bf.grid(row=0,column=2,padx=(0,10),pady=8)
        ctk.CTkButton(bf,text=t("upd_btn_now"),width=90,height=30,corner_radius=6,
                      fg_color=C["accent2"],hover_color=C["accent"],
                      font=ctk.CTkFont(size=12),command=on_update).pack(side="left",padx=(0,6))
        ctk.CTkButton(bf,text=t("upd_btn_later"),width=70,height=30,corner_radius=6,
                      fg_color=C["card2"],hover_color=C["border2"],
                      font=ctk.CTkFont(size=12),text_color=C["text3"],
                      command=on_dismiss).pack(side="left")


class UpdateDialog(D):
    def __init__(self, parent, current, latest, notes):
        super().__init__(parent, t("upd_dialog_title"), 580, 420)
        self.result="later"; self.grid_rowconfigure(1,weight=1)
        self._hdr("🚀",f"{t('upd_dialog_hdr')}: {latest}  ({t('upd_dialog_current')}: {current})",C["green"])
        c=ctk.CTkFrame(self,fg_color=C["card"],corner_radius=10)
        c.grid(row=1,column=0,padx=16,pady=(12,8),sticky="nsew")
        c.grid_columnconfigure(0,weight=1); c.grid_rowconfigure(0,weight=1)
        tb=ctk.CTkTextbox(c,fg_color=C["card2"],text_color=C["text2"],font=ctk.CTkFont(size=12))
        tb.grid(row=0,column=0,padx=12,pady=12,sticky="nsew")
        tb.insert("1.0",notes.strip() or t("upd_no_notes")); tb.configure(state="disabled")
        self._btns(2,[(t("upd_btn_install"),self._upd,C["green2"],C["green"]),
                      (t("upd_btn_later"),  self._later,C["card"],C["border2"])])
        self.protocol("WM_DELETE_WINDOW",self._later)
    def _upd(self):   self.result="update"; self.destroy()
    def _later(self): self.result="later";  self.destroy()


class DuplicateDialog(D):
    def __init__(self,parent,dup_items,total,unique_count,dup_rows):
        super().__init__(parent,t("dup_title"),620,500)
        self.ok=False; self.remove=False; self.save=False
        self.grid_rowconfigure(1,weight=1)
        self._hdr("⚠",f"{t('dup_hdr')} — {dup_rows} {t('dup_rows')}",C["amber"])
        c=ctk.CTkFrame(self,fg_color=C["card"],corner_radius=10)
        c.grid(row=1,column=0,padx=16,pady=(12,8),sticky="nsew")
        c.grid_columnconfigure(0,weight=1); c.grid_rowconfigure(1,weight=1)
        ctk.CTkLabel(c,text=f"{t('dup_total')}: {total}   •   {t('dup_unique')}: {unique_count}   •   {t('dup_rows')}: {dup_rows}",
                     text_color=C["text2"],font=ctk.CTkFont(size=12)
                     ).grid(row=0,column=0,padx=14,pady=(12,6),sticky="w")
        tb=ctk.CTkTextbox(c,font=ctk.CTkFont(family="Consolas",size=11),
                          fg_color=C["card2"],text_color=C["text2"],corner_radius=6)
        tb.grid(row=1,column=0,padx=12,pady=(0,12),sticky="nsew")
        for code,count in dup_items.items(): tb.insert("end",f"  {code}  ×{count}\n")
        tb.configure(state="disabled")
        opts=ctk.CTkFrame(self,fg_color=C["card"],corner_radius=10)
        opts.grid(row=2,column=0,padx=16,pady=4,sticky="ew")
        self._rm=ctk.BooleanVar(value=True); self._sv=ctk.BooleanVar(value=True)
        for i,(txt,var) in enumerate([(t("dup_rm_opt"),self._rm),(t("dup_save_opt"),self._sv)]):
            ctk.CTkCheckBox(opts,text=txt,variable=var,fg_color=C["accent"],font=ctk.CTkFont(size=12)
                            ).grid(row=i,column=0,padx=14,pady=(12 if i==0 else 4, 4 if i==0 else 12),sticky="w")
        self._btns(3,[(t("btn_continue_dup"),self._ok,C["accent2"],C["accent"]),
                      (t("btn_cancel"),self.destroy,C["card"],C["border2"])])
        self.protocol("WM_DELETE_WINDOW",self.destroy)
    def _ok(self): self.ok=True; self.remove=self._rm.get(); self.save=self._sv.get(); self.destroy()


class UnknownDialog(D):
    def __init__(self,parent,cikkszam):
        super().__init__(parent,t("unk_title"),440,220)
        self.action="stop"; self.apply_all=False
        self._hdr("?",f"{t('unk_hdr')}:  {cikkszam}",C["amber"])
        ctk.CTkLabel(self,text=t("unk_body"),text_color=C["text2"],font=ctk.CTkFont(size=12)
                     ).grid(row=1,column=0,padx=20,pady=(16,8),sticky="w")
        self._all=ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(self,text=t("unk_apply_all"),variable=self._all,fg_color=C["accent"],
                        font=ctk.CTkFont(size=12)).grid(row=2,column=0,padx=20,pady=(0,12),sticky="w")
        self._btns(3,[(t("btn_skip"), self._skip,C["card"],C["border2"]),
                      (t("btn_stop_unk"),self._stop,C["red2"],C["red"])])
        self.protocol("WM_DELETE_WINDOW",self._stop)
    def _skip(self): self.action="skip"; self.apply_all=self._all.get(); self.destroy()
    def _stop(self): self.action="stop"; self.apply_all=self._all.get(); self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# WIDGETS
# ══════════════════════════════════════════════════════════════════════════════
class StatCard(ctk.CTkFrame):
    def __init__(self,parent,label,value="—",accent=None,icon=""):
        super().__init__(parent,fg_color=C["card"],corner_radius=12,
                         border_width=1,border_color=C["border"])
        self._accent=accent or C["text2"]; self.grid_columnconfigure(0,weight=1)
        top=ctk.CTkFrame(self,fg_color="transparent")
        top.grid(row=0,column=0,padx=16,pady=(14,4),sticky="ew")
        ctk.CTkLabel(top,text=f"{icon}  {label}" if icon else label,
                     font=ctk.CTkFont(size=11),text_color=C["text3"]).pack(side="left")
        self._val=ctk.CTkLabel(self,text=str(value),
                               font=ctk.CTkFont(size=26,weight="bold"),text_color=self._accent)
        self._val.grid(row=1,column=0,padx=16,pady=(0,14),sticky="w")
    def set(self,v,color=None): self._val.configure(text=str(v),text_color=color or self._accent)


class ResultTable(ctk.CTkScrollableFrame):
    def __init__(self,parent,**kw):
        kw.setdefault("fg_color","transparent")
        super().__init__(parent,**kw)
        self.grid_columnconfigure(0,weight=3); self.grid_columnconfigure(1,weight=2)
        self.grid_columnconfigure(2,weight=2); self._rows=0; self._draw_header()
    def _draw_header(self):
        for col,(text,anch) in enumerate([("Cikkszám","w"),("Elérhetőség","center"),("Ár","center")]):
            ctk.CTkLabel(self,text=text,font=ctk.CTkFont(size=11,weight="bold"),
                         text_color=C["text3"]).grid(row=0,column=col,
                         padx=(14 if col==0 else 4),pady=(4,6),sticky=anch)
    def add_row(self,code,avail,price=""):
        r=self._rows+1; self._rows+=1
        styles={"Van":(C["green"],C["green_bg"],"●"),"Nincs":(C["red"],C["red_bg"],"●"),
                "Külső raktár":(C["amber"],C["amber_bg"],"●")}
        color,bg,dot=styles.get(avail,(C["text3"],C["card2"],"○"))
        rb=C["card"] if r%2==0 else C["card2"]
        ctk.CTkLabel(self,text=code,font=ctk.CTkFont(family="Consolas",size=12),
                     text_color=C["text"],fg_color=rb,anchor="w",corner_radius=0
                     ).grid(row=r,column=0,padx=(12,4),pady=1,sticky="ew")
        af=ctk.CTkFrame(self,fg_color=bg,corner_radius=6)
        af.grid(row=r,column=1,padx=4,pady=1,sticky="ew")
        ctk.CTkLabel(af,text=f"{dot}  {avail}",font=ctk.CTkFont(size=11,weight="bold"),
                     text_color=color).pack(pady=3,padx=8)
        ctk.CTkLabel(self,text=price or "—",font=ctk.CTkFont(size=12),
                     text_color=C["text2"],fg_color=rb,anchor="center",corner_radius=0
                     ).grid(row=r,column=2,padx=(4,12),pady=1,sticky="ew")
    def clear(self):
        for w in self.winfo_children():
            if int(w.grid_info().get("row",0))>0: w.destroy()
        self._rows=0

# ── UI helpers ────────────────────────────────────────────────────────────────
def _sb_sec(parent,key,row):
    ctk.CTkLabel(parent,text=t(key).upper(),font=ctk.CTkFont(size=9,weight="bold"),
                 text_color=C["text3"]).grid(row=row,column=0,padx=20,pady=(18,4),sticky="w")

def _card(parent,row):
    f=ctk.CTkFrame(parent,fg_color=C["card"],corner_radius=10,
                   border_width=1,border_color=C["border"])
    f.grid(row=row,column=0,padx=12,pady=(0,2),sticky="ew")
    f.grid_columnconfigure(0,weight=1); return f

def _flbl(parent,key,row):
    ctk.CTkLabel(parent,text=t(key),font=ctk.CTkFont(size=11),
                 text_color=C["text3"]).grid(row=row,column=0,padx=14,pady=(10,2),sticky="w")

def _fent(parent,row,placeholder="",show=""):
    e=ctk.CTkEntry(parent,placeholder_text=placeholder,show=show,
                   fg_color=C["card2"],border_color=C["border2"],
                   text_color=C["text"],placeholder_text_color=C["text3"],
                   corner_radius=7,height=34)
    e.grid(row=row,column=0,padx=14,pady=(0,10),sticky="ew"); return e

def _btn(parent,key,row,col,fg,hov,cmd,span=1,bold=False,h=38,text=None):
    b=ctk.CTkButton(parent,text=text or t(key),command=cmd,height=h,corner_radius=8,
                    fg_color=fg,hover_color=hov,
                    font=ctk.CTkFont(size=12,weight="bold" if bold else "normal"))
    b.grid(row=row,column=col,columnspan=span,padx=6,pady=4,sticky="ew"); return b


# ══════════════════════════════════════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self._cfg=load_config()
        set_language(self._cfg.get("language","hu"))
        self._apply_theme(self._cfg["theme"],init=True)

        self.title("CikkChecker"); self.geometry("1440x900"); self.minsize(1100,700)
        self.configure(fg_color=C["bg"])

        # Window icon
        self._logo_img = None
        self._load_window_icon()

        self._proc=None; self._stop_req=False; self._can_continue=False
        self._unknown_mode=None; self._log_q=queue.Queue()
        self._stats={"van":0,"nincs":0,"kulso":0,"ism":0}
        self._update_banner_widget=None
        self._latest_version=None; self._latest_notes=""; self._latest_asset=""

        self._build_ui()
        self._apply_cfg_to_ui()
        self._poll_log()
        self._update_start_btn()
        self.protocol("WM_DELETE_WINDOW",self._quit)

        if self._cfg.get("auto_update",True) and self._cfg.get("update_interval","on_start")!="never":
            self.after(3000,lambda: threading.Thread(target=self._check_update,daemon=True).start())

    # ── Logo / icon ──────────────────────────────────────────────────────────
    def _load_window_icon(self):
        try:
            from PIL import Image, ImageTk
            if os.path.exists(ICON_PATH):
                self.iconbitmap(ICON_PATH)
            elif os.path.exists(LOGO_PATH):
                img = Image.open(LOGO_PATH).resize((32,32), Image.LANCZOS)
                self._icon_tk = ImageTk.PhotoImage(img)
                self.iconphoto(True, self._icon_tk)
        except Exception:
            pass

    def _get_logo_image(self, size=(44,44)):
        try:
            from PIL import Image
            import customtkinter as _ctk
            if os.path.exists(LOGO_PATH):
                img = Image.open(LOGO_PATH).resize(size, Image.LANCZOS)
                return _ctk.CTkImage(light_image=img, dark_image=img, size=size)
        except Exception:
            pass
        return None

    # ── Theme & language ──────────────────────────────────────────────────────
    def _apply_theme(self,theme_name,init=False):
        global C; C=THEMES.get(theme_name,THEMES["dark"])
        ctk.set_appearance_mode("dark" if theme_name=="dark" else "light")
        if not init:
            self.configure(fg_color=C["bg"])
            for w in self.winfo_children(): w.destroy()
            self._update_banner_widget=None
            self._build_ui(); self._apply_cfg_to_ui()

    def _apply_cfg_to_ui(self):
        cfg=self._cfg
        if not hasattr(self,"_url"): return
        for attr,key in [("_url","last_url"),("_xlsx","last_xlsx"),("_txt","last_txt")]:
            val=cfg.get(key,"")
            if val:
                w=getattr(self,attr); w.delete(0,"end"); w.insert(0,val)
        self._sleep.delete(0,"end"); self._sleep.insert(0,str(cfg.get("sleep_seconds",0.8)))
        self._save_ev.delete(0,"end"); self._save_ev.insert(0,str(cfg.get("save_every",100)))
        sort_map={"abc":t("sort_abc_asc"),"abc_desc":t("sort_abc_desc"),
                  "num":t("sort_num_asc"),"num_desc":t("sort_num_desc"),"none":t("sort_none")}
        self._sort_var.set(sort_map.get(cfg.get("sort_order","abc"),t("sort_abc_asc")))
        self._auto_rm.set(cfg.get("auto_rm_dup",False))
        self._dont_ask.set(cfg.get("dont_ask_dup",False))
        self._save_dup.set(cfg.get("save_dup_txt",True))

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(1,weight=1); self.grid_rowconfigure(0,weight=1)

        # Sidebar
        sb=ctk.CTkScrollableFrame(self,width=300,fg_color=C["sidebar"],corner_radius=0,
                                  border_width=0,scrollbar_button_color=C["border"])
        sb.grid(row=0,column=0,sticky="nsew"); sb.grid_columnconfigure(0,weight=1)

        # Logo row
        lr=ctk.CTkFrame(sb,fg_color="transparent")
        lr.grid(row=0,column=0,padx=(16,12),pady=(24,4),sticky="ew")
        lr.grid_columnconfigure(1,weight=1)
        lt=ctk.CTkFrame(lr,fg_color="transparent"); lt.grid(row=0,column=0,sticky="w")
        logo_img = self._get_logo_image(size=(38,38))
        if logo_img:
            ctk.CTkLabel(lt, text="", image=logo_img, fg_color="transparent").pack(side="left", padx=(0,8))
        ctk.CTkLabel(lt,text="Cikk",font=ctk.CTkFont(size=22,weight="bold"),
                     text_color=C["text"]).pack(side="left")
        ctk.CTkLabel(lt,text="Checker",font=ctk.CTkFont(size=22,weight="bold"),
                     text_color=C["accent"]).pack(side="left")
        ctk.CTkLabel(lt,text=f"  v{APP_VERSION}",font=ctk.CTkFont(size=11),
                     text_color=C["text3"]).pack(side="left",pady=(5,0))
        ctk.CTkButton(lr,text="⚙",width=34,height=34,corner_radius=8,
                      fg_color=C["card"],hover_color=C["border2"],
                      font=ctk.CTkFont(size=16),text_color=C["text3"],
                      command=self._open_settings).grid(row=0,column=2,sticky="e")

        ctk.CTkFrame(sb,height=1,fg_color=C["border"]).grid(row=1,column=0,padx=20,pady=(8,0),sticky="ew")

        # Connection
        _sb_sec(sb,"s_connection",2); c1=_card(sb,3)
        _flbl(c1,"url_label",0); self._url=_fent(c1,1,"https://")
        self._url.bind("<FocusOut>",self._on_url_blur)
        dr=ctk.CTkFrame(c1,fg_color="transparent")
        dr.grid(row=2,column=0,padx=14,pady=(0,4),sticky="ew"); dr.grid_columnconfigure(0,weight=1)
        self._login_var=ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(dr,text=t("login_required"),variable=self._login_var,
                        fg_color=C["accent"],font=ctk.CTkFont(size=12),
                        command=self._toggle_login).grid(row=0,column=0,sticky="w")
        self._detect_lbl=ctk.CTkLabel(dr,text="",font=ctk.CTkFont(size=10),text_color=C["text3"])
        self._detect_lbl.grid(row=1,column=0,pady=(2,6),sticky="w")

        # Login
        _sb_sec(sb,"s_login",4); self._login_card=_card(sb,5)
        _flbl(self._login_card,"username",0); self._user=_fent(self._login_card,1)
        _flbl(self._login_card,"password",2); self._pw=_fent(self._login_card,3,show="*")

        # Files
        _sb_sec(sb,"s_files",6); c3=_card(sb,7)
        _flbl(c3,"codes_txt",0); self._txt=_fent(c3,1)
        _btn(c3,"browse_txt",2,0,C["border2"],C["border"],self._browse_txt)
        _flbl(c3,"excel_output",3); self._xlsx=_fent(c3,4)
        _btn(c3,"browse_xlsx",5,0,C["border2"],C["border"],self._browse_xlsx)
        ctk.CTkFrame(c3,height=6,fg_color="transparent").grid(row=6,column=0)

        # Settings
        _sb_sec(sb,"s_settings",8); c4=_card(sb,9)
        c4.grid_columnconfigure((0,1),weight=1)
        _flbl(c4,"wait_sec",0)
        self._sleep=ctk.CTkEntry(c4,fg_color=C["card2"],border_color=C["border2"],
                                 text_color=C["text"],corner_radius=7,height=32)
        self._sleep.insert(0,"0.8"); self._sleep.grid(row=1,column=0,padx=(14,6),pady=(0,10),sticky="ew")
        _flbl(c4,"save_every",0)
        self._save_ev=ctk.CTkEntry(c4,fg_color=C["card2"],border_color=C["border2"],
                                   text_color=C["text"],corner_radius=7,height=32)
        self._save_ev.insert(0,"100"); self._save_ev.grid(row=1,column=1,padx=(6,14),pady=(0,10),sticky="ew")
        _flbl(c4,"sort_order",2)
        self._sort_var=ctk.StringVar(value=t("sort_abc_asc"))
        ctk.CTkOptionMenu(c4,values=[t("sort_abc_asc"),t("sort_abc_desc"),
                                     t("sort_num_asc"),t("sort_num_desc"),t("sort_none")],
                          variable=self._sort_var,fg_color=C["card2"],button_color=C["border2"],
                          button_hover_color=C["accent2"],text_color=C["text"],
                          dropdown_fg_color=C["card"],dropdown_hover_color=C["border2"],
                          font=ctk.CTkFont(size=12),corner_radius=7
                          ).grid(row=3,column=0,columnspan=2,padx=14,pady=(0,10),sticky="ew")
        self._auto_rm=ctk.BooleanVar(value=False)
        self._dont_ask=ctk.BooleanVar(value=False)
        self._save_dup=ctk.BooleanVar(value=True)
        for i,(key,var) in enumerate([("auto_rm_dup",self._auto_rm),
                                       ("dont_ask_dup",self._dont_ask),
                                       ("save_dup_txt",self._save_dup)]):
            ctk.CTkCheckBox(c4,text=t(key),variable=var,fg_color=C["accent"],
                            font=ctk.CTkFont(size=11)
                            ).grid(row=4+i,column=0,columnspan=2,padx=14,
                                   pady=(6 if i==0 else 2, 10 if i==2 else 2),sticky="w")

        # Actions
        _sb_sec(sb,"s_actions",10); c5=_card(sb,11)
        c5.grid_columnconfigure((0,1),weight=1)
        self._start_btn=ctk.CTkButton(c5,text=t("btn_start"),height=44,corner_radius=8,
                                      fg_color=C["green2"],hover_color=C["green"],
                                      font=ctk.CTkFont(size=14,weight="bold"),
                                      command=self._start_action)
        self._start_btn.grid(row=0,column=0,columnspan=2,padx=12,pady=(14,6),sticky="ew")
        _btn(c5,"btn_stop",   1,0,C["red2"],   C["red"],   self._stop)
        _btn(c5,"btn_restart",1,1,C["amber2"], C["amber"], self._restart)
        _btn(c5,"btn_check_upd",2,0,C["border2"],C["border"],
             lambda: threading.Thread(target=self._check_update,daemon=True).start(),span=2)
        _btn(c5,"btn_clear_log",3,0,C["border2"],C["border"],self._clear_log,span=2)
        ctk.CTkFrame(c5,height=8,fg_color="transparent").grid(row=4,column=0)

        self._toggle_login()

        # Main area
        main=ctk.CTkFrame(self,fg_color="transparent")
        main.grid(row=0,column=1,sticky="nsew",padx=20,pady=20)
        main.grid_columnconfigure(0,weight=1); main.grid_rowconfigure(3,weight=1)
        self._banner_slot=main

        # Stat cards
        stats=ctk.CTkFrame(main,fg_color="transparent")
        stats.grid(row=1,column=0,sticky="ew",pady=(0,14))
        for i in range(5): stats.grid_columnconfigure(i,weight=1)
        self._s_total=StatCard(stats,t("stat_total"),"—", C["text2"],"📋")
        self._s_done =StatCard(stats,t("stat_done"), "0", C["accent"],"✓")
        self._s_van  =StatCard(stats,t("stat_van"),  "0", C["green"],"●")
        self._s_nincs=StatCard(stats,t("stat_nincs"),"0", C["red"],  "●")
        self._s_kulso=StatCard(stats,t("stat_kulso"),"0", C["amber"],"●")
        for i,w in enumerate([self._s_total,self._s_done,self._s_van,self._s_nincs,self._s_kulso]):
            w.grid(row=0,column=i,padx=5,sticky="ew")

        # Progress
        prog=ctk.CTkFrame(main,fg_color=C["card"],corner_radius=12,
                          border_width=1,border_color=C["border"])
        prog.grid(row=2,column=0,sticky="ew",pady=(0,14)); prog.grid_columnconfigure(0,weight=1)
        tr=ctk.CTkFrame(prog,fg_color="transparent")
        tr.grid(row=0,column=0,padx=18,pady=(14,6),sticky="ew"); tr.grid_columnconfigure(0,weight=1)
        self._status_lbl=ctk.CTkLabel(tr,text=t("status_waiting"),
                                      font=ctk.CTkFont(size=15,weight="bold"),text_color=C["text"])
        self._status_lbl.grid(row=0,column=0,sticky="w")
        self._pct_lbl=ctk.CTkLabel(tr,text=t("progress_label"),
                                   font=ctk.CTkFont(size=12),text_color=C["text3"])
        self._pct_lbl.grid(row=0,column=1,sticky="e")
        self._progress=ctk.CTkProgressBar(prog,height=6,corner_radius=3,
                                          progress_color=C["accent"],fg_color=C["border"])
        self._progress.grid(row=1,column=0,padx=18,pady=(0,14),sticky="ew"); self._progress.set(0)

        # Bottom
        bot=ctk.CTkFrame(main,fg_color="transparent")
        bot.grid(row=3,column=0,sticky="nsew")
        bot.grid_columnconfigure(0,weight=5); bot.grid_columnconfigure(1,weight=2)
        bot.grid_rowconfigure(0,weight=1)

        rc=ctk.CTkFrame(bot,fg_color=C["card"],corner_radius=12,
                        border_width=1,border_color=C["border"])
        rc.grid(row=0,column=0,padx=(0,10),sticky="nsew")
        rc.grid_columnconfigure(0,weight=1); rc.grid_rowconfigure(1,weight=1)
        rh=ctk.CTkFrame(rc,fg_color="transparent")
        rh.grid(row=0,column=0,padx=16,pady=(12,6),sticky="ew"); rh.grid_columnconfigure(0,weight=1)
        ctk.CTkLabel(rh,text=t("results_title"),font=ctk.CTkFont(size=13,weight="bold"),
                     text_color=C["text2"]).grid(row=0,column=0,sticky="w")
        ctk.CTkButton(rh,text=t("btn_clear_res"),width=60,height=26,corner_radius=6,
                      fg_color=C["border"],hover_color=C["border2"],
                      font=ctk.CTkFont(size=11),text_color=C["text3"],
                      command=self._clear_results).grid(row=0,column=1,sticky="e")
        self._table=ResultTable(rc,fg_color="transparent")
        self._table.grid(row=1,column=0,padx=8,pady=(0,8),sticky="nsew")

        rs=ctk.CTkFrame(bot,fg_color="transparent")
        rs.grid(row=0,column=1,sticky="nsew")
        rs.grid_columnconfigure(0,weight=1); rs.grid_rowconfigure(0,weight=3); rs.grid_rowconfigure(1,weight=2)

        lc=ctk.CTkFrame(rs,fg_color=C["card"],corner_radius=12,border_width=1,border_color=C["border"])
        lc.grid(row=0,column=0,pady=(0,10),sticky="nsew")
        lc.grid_columnconfigure(0,weight=1); lc.grid_rowconfigure(1,weight=1)
        ctk.CTkLabel(lc,text=t("log_title"),font=ctk.CTkFont(size=12,weight="bold"),
                     text_color=C["text3"]).grid(row=0,column=0,padx=14,pady=(10,4),sticky="w")
        self._log_box=ctk.CTkTextbox(lc,font=ctk.CTkFont(family="Consolas",size=10),
                                     fg_color=C["card2"],text_color=C["text2"],corner_radius=8)
        self._log_box.grid(row=1,column=0,padx=10,pady=(0,10),sticky="nsew")

        mc=ctk.CTkFrame(rs,fg_color=C["card"],corner_radius=12,border_width=1,border_color=C["border"])
        mc.grid(row=1,column=0,sticky="nsew")
        mc.grid_columnconfigure(0,weight=1); mc.grid_rowconfigure(1,weight=1)
        ctk.CTkLabel(mc,text=t("manual_title"),font=ctk.CTkFont(size=12,weight="bold"),
                     text_color=C["text3"]).grid(row=0,column=0,padx=14,pady=(10,2),sticky="w")
        self._manual=ctk.CTkTextbox(mc,font=ctk.CTkFont(family="Consolas",size=11),
                                    fg_color=C["card2"],text_color=C["text2"],corner_radius=8)
        self._manual.grid(row=1,column=0,padx=10,pady=(0,10),sticky="nsew")

    # ── Settings ──────────────────────────────────────────────────────────────
    def _open_settings(self):
        def on_save(new_cfg):
            old_theme=self._cfg.get("theme","dark")
            old_lang=self._cfg.get("language","hu")
            self._cfg.update(new_cfg); save_config(self._cfg)
            lang_changed=new_cfg.get("language","hu")!=old_lang
            theme_changed=new_cfg["theme"]!=old_theme
            if lang_changed:
                set_language(new_cfg["language"])
            if theme_changed or lang_changed:
                self._apply_theme(new_cfg["theme"])
                self._apply_cfg_to_ui()
            else:
                self._apply_cfg_to_ui()
        SettingsDialog(self,self._cfg,on_save)

    # ── Update ────────────────────────────────────────────────────────────────
    def _check_update(self):
        try:
            r=requests.get(f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest",timeout=10)
            r.raise_for_status(); data=r.json()
        except: return
        latest=data.get("tag_name","").strip()
        notes=data.get("body","").strip()
        def _v(s):
            s=s.strip().lstrip("v")
            return tuple(int(x) if x.isdigit() else 0 for x in s.split("."))
        if not latest or _v(latest)<=_v(APP_VERSION): self._log(t("upd_not_found")); return
        asset_url=next((a["browser_download_url"] for a in data.get("assets",[])
                        if a.get("name")==UPDATE_ASSET),None)
        if not asset_url: return
        self._latest_version=latest; self._latest_notes=notes; self._latest_asset=asset_url
        self._log(f"{t('upd_banner')}: {latest}")
        self.after(0,self._show_update_banner)

    def _show_update_banner(self):
        if self._update_banner_widget:
            try: self._update_banner_widget.destroy()
            except: pass
        banner=UpdateBanner(self._banner_slot,self._latest_version,
                            on_update=self._open_update_dialog,on_dismiss=self._dismiss_banner)
        banner.grid(row=0,column=0,sticky="ew",pady=(0,14))
        self._update_banner_widget=banner

    def _dismiss_banner(self):
        if self._update_banner_widget:
            self._update_banner_widget.destroy(); self._update_banner_widget=None

    def _open_update_dialog(self):
        self._dismiss_banner()
        dlg=UpdateDialog(self,APP_VERSION,self._latest_version,self._latest_notes)
        self.wait_window(dlg)
        if dlg.result=="update":
            threading.Thread(target=self._dl_update,
                             args=(self._latest_asset,self._latest_version),daemon=True).start()

    def _dl_update(self,url,latest):
        self._log(f"{t('upd_downloading')}: {latest}..."); self._set_status(f"{t('upd_downloading')}: {latest}...")
        try:
            tmp=os.path.join(tempfile.gettempdir(),UPDATE_ASSET)
            with requests.get(url,stream=True,timeout=60) as r:
                r.raise_for_status()
                with open(tmp,"wb") as f:
                    for chunk in r.iter_content(8192):
                        if chunk: f.write(chunk)
            import subprocess as sp; sp.Popen([tmp],shell=True)
            self.after(1000,self._quit)
        except Exception as e:
            self._log(f"{t('upd_error')}: {e}"); self._set_status(t("upd_failed"))

    # ── UI helpers ────────────────────────────────────────────────────────────
    def _toggle_login(self):
        if self._login_var.get(): self._login_card.grid(row=5,column=0,padx=12,pady=(0,2),sticky="ew")
        else: self._login_card.grid_remove()

    def _on_url_blur(self,event=None):
        url=self._url.get().strip()
        if not url or url=="https://": return
        threading.Thread(target=self._detect_login,args=(url,),daemon=True).start()

    def _detect_login(self,url):
        self.after(0,lambda: self._detect_lbl.configure(text=t("detecting"),text_color=C["text3"]))
        try:
            if not url.startswith("http"): url="https://"+url
            resp=requests.get(f"{url}/product-search/test/0?1",timeout=8,allow_redirects=True)
            html=resp.text.lower()
            required=("/login" in resp.url or
                      any(k in html for k in ["bejelentkezés","belépés","felhasználónév","user-name","password"]))
            def _u():
                self._login_var.set(required); self._toggle_login()
                if required:
                    self._detect_lbl.configure(text=t("login_auto_yes"),text_color=C["amber"])
                else:
                    self._detect_lbl.configure(text=t("login_auto_no"),text_color=C["green"])
            self.after(0,_u)
        except:
            self.after(0,lambda: self._detect_lbl.configure(text=t("login_check_fail"),text_color=C["text3"]))

    def _update_start_btn(self):
        self._start_btn.configure(text=t("btn_continue") if self._can_continue else t("btn_start"))

    def _log(self,msg):
        ts=datetime.now().strftime("%H:%M:%S"); self._log_q.put(f"[{ts}]  {msg}")

    def _poll_log(self):
        lines=[]
        try:
            while True: lines.append(self._log_q.get_nowait())
        except queue.Empty: pass
        if lines:
            self._log_box.insert("end","\n".join(lines)+"\n"); self._log_box.see("end")
        self.after(150,self._poll_log)

    def _set_status(self,msg): self.after(0,lambda: self._status_lbl.configure(text=msg))

    def _set_progress(self,cur,tot):
        pct=int(cur/tot*100) if tot else 0
        def _u():
            self._progress.set(0 if not tot else cur/tot)
            self._pct_lbl.configure(text=f"{cur} / {tot}  ({pct}%)")
            self._s_done.set(cur)
        self.after(0,_u)

    def _update_stats(self):
        self._s_van.set(self._stats["van"])
        self._s_nincs.set(self._stats["nincs"])
        self._s_kulso.set(self._stats["kulso"])

    def _clear_log(self): self._log_box.delete("1.0","end")
    def _clear_results(self): self._table.clear()

    def _browse_txt(self):
        p=filedialog.askopenfilename(title="TXT",filetypes=[("Text","*.txt"),("*","*.*")])
        if p: self._txt.delete(0,"end"); self._txt.insert(0,p); self._cfg["last_txt"]=p

    def _browse_xlsx(self):
        p=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if p: self._xlsx.delete(0,"end"); self._xlsx.insert(0,p); self._cfg["last_xlsx"]=p

    def _quit(self):
        if hasattr(self,"_url"):
            self._cfg["last_url"]=self._url.get().strip()
            self._cfg["last_xlsx"]=self._xlsx.get().strip()
            self._cfg["last_txt"]=self._txt.get().strip()
        save_config(self._cfg); self._stop(); self.destroy()

    def _get_sort_arg(self):
        rev={"abc":t("sort_abc_asc"),"abc_desc":t("sort_abc_desc"),
             "num":t("sort_num_asc"),"num_desc":t("sort_num_desc"),"none":t("sort_none")}
        rev2={v:k for k,v in rev.items()}
        return rev2.get(self._sort_var.get(),"abc")

    # ── Process ───────────────────────────────────────────────────────────────
    def _start_action(self):
        if self._can_continue: self._run(skip_dup=True)
        else: self._run()

    def _stop(self):
        self._stop_req=True
        if self._proc:
            try: self._proc.stdin.write("stop\n"); self._proc.stdin.flush()
            except: pass
        self._can_continue=True; self._update_start_btn(); self._log(t("log_stop_req"))

    def _restart(self):
        self._stop_req=True
        if self._proc:
            try: self._proc.stdin.write("stop\n"); self._proc.stdin.flush()
            except: pass
        self._can_continue=False
        self._stats={"van":0,"nincs":0,"kulso":0,"ism":0}
        self._update_start_btn(); self._log(t("log_restart_req"))
        self.after(800,lambda: self._run(skip_dup=False,fresh=True))

    def _collect_codes(self):
        codes=[]
        txt=self._txt.get().strip(); manual=self._manual.get("1.0","end").strip()
        if txt and os.path.exists(txt):
            with open(txt,encoding="utf-8") as f:
                codes+=[l.strip() for l in f if l.strip()]
        if manual:
            codes+=[l.strip() for l in manual.splitlines() if l.strip()]
        return codes

    def _run(self,skip_dup=False,fresh=False):
        if self._proc and self._proc.poll() is None:
            messagebox.showwarning("",t("warn_already_running")); return
        base_url=self._url.get().strip(); user=self._user.get().strip()
        pw=self._pw.get().strip(); xlsx=self._xlsx.get().strip()
        req_login=self._login_var.get()
        if not base_url or base_url=="https://":
            messagebox.showerror("",t("err_no_url")); return
        if not xlsx:
            messagebox.showerror("",t("err_no_xlsx")); return
        if req_login and (not user or not pw):
            messagebox.showerror("",t("err_no_creds")); return
        if not xlsx.lower().endswith(".xlsx"):
            messagebox.showerror("",t("err_xlsx_ext")); return
        try: sleep_s=float(self._sleep.get().strip())
        except: sleep_s=0.8
        try: save_every=int(self._save_ev.get().strip())
        except: save_every=100
        codes=self._collect_codes()
        if not codes:
            messagebox.showerror("",t("err_no_codes")); return
        xlsx_folder=os.path.dirname(xlsx) or "."
        os.makedirs(xlsx_folder,exist_ok=True)
        csv_path=os.path.join(xlsx_folder,"cikkchecker_results.csv")
        if fresh and os.path.exists(csv_path):
            try: os.remove(csv_path)
            except: pass
        counter=Counter(codes); dup_items={c:n for c,n in counter.items() if n>1}
        unique=list(dict.fromkeys(codes))
        if dup_items and not skip_dup:
            dup_rows=sum(n-1 for n in counter.values() if n>1)
            if self._dont_ask.get():
                if self._save_dup.get(): self._save_dup_txt(xlsx_folder,dup_items)
                if self._auto_rm.get(): codes=unique
            else:
                dlg=DuplicateDialog(self,dup_items,len(codes),len(unique),dup_rows)
                self.wait_window(dlg)
                if not dlg.ok: return
                if dlg.save:   self._save_dup_txt(xlsx_folder,dup_items)
                if dlg.remove: codes=unique
        tmp=tempfile.NamedTemporaryFile(mode="w",suffix=".txt",encoding="utf-8",delete=False)
        tmp.write("\n".join(codes)); tmp.close()
        self._stats={"van":0,"nincs":0,"kulso":0,"ism":0}
        self._s_total.set(len(set(codes))); self._s_done.set(0)
        self._s_van.set(0); self._s_nincs.set(0); self._s_kulso.set(0)
        self._progress.set(0); self._stop_req=False; self._unknown_mode=None
        if not os.path.exists(BINARY_PATH):
            messagebox.showerror("",f"{t('err_no_binary')}:\n{BINARY_PATH}"); return
        cmd=[BINARY_PATH,"--base-url",base_url,"--user",user,"--password",pw,
             "--codes-file",tmp.name,"--csv-output",csv_path,
             "--sleep-seconds",str(sleep_s),"--save-every",str(save_every),
             "--sort-order",self._get_sort_arg()]
        if req_login: cmd.append("--requires-login")
        self._log("═"*42); self._log(f"{t('log_starting')} — {len(codes)} {t('log_codes')}")
        self._set_status(t("status_starting")); self._can_continue=True; self._update_start_btn()
        self._xlsx_path=xlsx; self._csv_path=csv_path
        self._proc=subprocess.Popen(cmd,stdin=subprocess.PIPE,stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE,text=True,bufsize=1)
        threading.Thread(target=self._read_proc,args=(tmp.name,),daemon=True).start()

    def _save_dup_txt(self,folder,dup_items):
        p=os.path.join(folder,"duplicates.txt")
        with open(p,"w",encoding="utf-8") as f:
            for c,n in dup_items.items(): f.write(f"{c};{n}\n")
        self._log(f"{t('log_dup_saved')}: {p}")

    def _read_proc(self,tmp_path):
        try:
            for line in self._proc.stdout:
                line=line.strip()
                if not line: continue
                try: msg=json.loads(line)
                except: self._log(line); continue
                kind=msg.get("kind","")
                if kind=="log":      self._log(msg["msg"])
                elif kind=="status": self._set_status(msg["msg"])
                elif kind=="progress": self._set_progress(msg["current"],msg["total"])
                elif kind=="login_detected":
                    req=msg.get("required",False)
                    def _u(r=req):
                        self._login_var.set(r); self._toggle_login()
                        self._detect_lbl.configure(
                            text=t("login_auto_yes") if r else t("login_auto_no"),
                            text_color=C["amber"] if r else C["green"])
                    self.after(0,_u)
                elif kind=="result":
                    avail=msg.get("elerhetoseg",""); price=msg.get("ar",""); code=msg.get("cikkszam","")
                    if avail=="Van": self._stats["van"]+=1
                    elif avail=="Nincs": self._stats["nincs"]+=1
                    elif avail=="Külső raktár": self._stats["kulso"]+=1
                    else: self._stats["ism"]+=1
                    self.after(0,self._update_stats)
                    self.after(0,lambda c=code,a=avail,p=price: self._table.add_row(c,a,p))
                elif kind=="unknown":
                    cmd=self._handle_unknown(msg["cikkszam"])
                    try: self._proc.stdin.write(cmd+"\n"); self._proc.stdin.flush()
                    except: pass
                elif kind=="error": self._log(f"[ERR] {msg['msg']}")
                elif kind=="done":
                    self._log(t("upd_done_export"))
                    export_excel(self._csv_path,self._xlsx_path)
                    self._log(f"{t('upd_excel_saved')}: {self._xlsx_path}")
                    self._set_status(t("upd_status_ok"))
                    self._can_continue=False; self.after(0,self._update_start_btn)
        except Exception as e: self._log(f"{t('log_read_err')}: {e}")
        finally:
            try: os.unlink(tmp_path)
            except: pass

    def _handle_unknown(self,cikkszam):
        if self._unknown_mode=="skip_all": return "skip"
        if self._unknown_mode=="stop_all": return "stop"
        box={"action":"stop","apply":False}; ev=threading.Event()
        def _show():
            dlg=UnknownDialog(self,cikkszam); self.wait_window(dlg)
            box["action"]=dlg.action; box["apply"]=dlg.apply_all; ev.set()
        self.after(0,_show); ev.wait()
        if box["apply"]:
            if box["action"]=="skip": self._unknown_mode="skip_all"
            else:                     self._unknown_mode="stop_all"
        return box["action"]


if __name__=="__main__":
    App().mainloop()
