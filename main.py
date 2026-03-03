# -*- coding: utf-8 -*-
"""
main.py — Aplikasi Desktop "Fill-in-the-Blanks" Formulir Ritual
Framework: CustomTkinter (modern Tkinter)

Fitur:
  1. Form Input  : Entry fields sesuai schema Excel Sheet 2
                    (Nama, Panggilan, Nama Mandarin, Penyebutan, Dari, Keluarga,
                     Keterangan).
  2. Tabel Preview: Treeview menampilkan data tersimpan di SQLite.
  3. Tombol Cetak : Generate PDF layer transparan & buka otomatis.
  4. Import Excel : Import data dari file Excel (Sheet 2).
  5. Export Backup: Export seluruh database ke Excel sebagai cadangan.
  6. Calibration  : Offset X/Y untuk koreksi posisi cetak printer.
"""

import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import customtkinter as ctk
from pypinyin import pinyin, Style as PinyinStyle

# ============================================================
# Import modul internal
# ============================================================
# Pastikan root project ada di sys.path
_ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if _ROOT_DIR not in sys.path:
    sys.path.insert(0, _ROOT_DIR)

from database.database import (
    init_db, insert_record, update_record, get_all_records, get_all_orders,
    get_items_by_order, delete_record, delete_order, update_order_name,
    bulk_delete_orders, import_from_excel, export_to_excel,
)
from modules.pdf_engine import generate_pdf_bytes
from modules.pdf_preview import PDFPreviewWindow
from modules.updater import UpdateChecker
from modules.dictionary_window import DictionaryWindow


# ============================================================
# Versi Aplikasi
# ============================================================
APP_VERSION = "1.2.0"

# ============================================================
# Konstanta UI
# ============================================================
APP_TITLE = "Formulir Ritual — Fill-in-the-Blanks (A4)"
APP_WIDTH = 1100
APP_HEIGHT = 720

# ============================================================
# Daftar Panggilan (Indonesia → Mandarin)
# ============================================================
# (nama_indonesia, aksara_mandarin, keluarga_otomatis)
PANGGILAN_OPTIONS: list[tuple[str, str, str]] = [
    ("Ayah", "父親", "Ayah Kandung"),
    ("Kakek (Pihak Ayah)", "祖父", "Kakek Kandung"),
    ("Nenek (Pihak Ayah)", "祖母", "Nenek Kandung"),
    ("Kakak Laki-laki Ayah", "伯父", "Paman (Kakak Ayah)"),
    ("Istri Kakak Laki-laki Ayah", "伯母", "Istri Paman (Kakak Ayah)"),
    ("Adik Laki-laki Ayah", "叔父", "Paman (Adik Ayah)"),
    ("Istri Adik Laki-laki Ayah", "嬸母", "Istri Paman (Adik Ayah)"),
    ("Saudara Perempuan Ayah (Bibi)", "姑母", "Bibi (Saudara Ayah)"),
    ("Suami Saudara Perempuan Ayah", "姑丈", "Suami Bibi (Saudara Ayah)"),
    ("Kakak Dari Kakek", "太伯祖考", "Kakak Kakek"),
    ("Ibu", "母親", "Ibu Kandung"),
    ("Kakek (Pihak Ibu)", "外祖父", "Kakek Luar"),
    ("Nenek (Pihak Ibu)", "外祖母", "Nenek Luar"),
    ("Saudara Laki-laki Ibu", "舅父", "Paman (Saudara Ibu)"),
    ("Istri Saudara Laki-laki Ibu", "舅母", "Istri Paman (Saudara Ibu)"),
    ("Saudara Perempuan Ibu", "姨母", "Bibi (Saudara Ibu)"),
    ("Suami Saudara Perempuan Ibu", "姨丈", "Suami Bibi (Saudara Ibu)"),
    # -- Kategori bernomor (dinamis via spinner) --
    ("Kakak Laki-laki", "亡胞__兄", "Kakak Laki-laki"),
    ("Kakak Perempuan", "亡胞__姊", "Kakak Perempuan"),
    ("Adik Laki-laki", "亡胞__弟", "Adik Laki-laki"),
    ("Adik Perempuan", "亡胞__妹", "Adik Perempuan"),
    # -----------------------------------------------
    ("Kakak/Adik Ipar Laki-laki", "姊夫", "Ipar Laki-laki"),
    ("Saudara Laki-laki Istri", "内兄", "Saudara Laki-laki Istri"),
    ("Suami Saudara Perempuan Istri", "襟兄", "Suami Saudara Perempuan Istri"),
    ("Kakak/Adik Ipar Perempuan", "嫂嫂", "Ipar Perempuan"),
    ("Saudara Perempuan Istri", "内姊", "Saudara Perempuan Istri"),
    ("Istri Saudara Laki-laki Istri", "内嫂", "Istri Saudara Laki-laki Istri"),
    ("Ayah Angkat", "養父", "Ayah Angkat"),
    ("Ibu Angkat", "養母", "Ibu Angkat"),
    ("Kakek Angkat (Ayah)", "養祖父", "Kakek Angkat"),
    ("Nenek Angkat (Ayah)", "養祖母", "Nenek Angkat"),
    ("Kakek Angkat (Ibu)", "養外祖父", "Kakek Angkat Luar"),
    ("Nenek Angkat (Ibu)", "養外祖母", "Nenek Angkat Luar"),
    ("Menantu Laki-laki", "女婿", "Menantu Laki-laki"),
    ("Menantu Perempuan", "媳妇", "Menantu Perempuan"),
]

# Kategori yang membutuhkan nomor urut (spinner)
_NUMBERED_CATEGORIES: set[str] = {
    "Kakak Laki-laki", "Kakak Perempuan",
    "Adik Laki-laki", "Adik Perempuan",
}

# Angka Mandarin untuk nomor urut
_MANDARIN_NUMBERS: dict[int, str] = {
    1: "長", 2: "二", 3: "三", 4: "四", 5: "五",
    6: "六", 7: "七", 8: "八", 9: "九", 10: "十",
}
# Untuk Adik, nomor 1 pakai 大 bukan 長
_MANDARIN_NUMBERS_ADIK: dict[int, str] = {
    1: "大", 2: "二", 3: "三", 4: "四", 5: "五",
    6: "六", 7: "七", 8: "八", 9: "九", 10: "十",
}


def _generate_numbered_mandarin(base_template: str, num: int,
                                 is_adik: bool = False) -> str:
    """Generate aksara Mandarin dinamis dari template + nomor.

    Template format: '亡胞__兄' → __ diganti angka Mandarin.
    """
    lookup = _MANDARIN_NUMBERS_ADIK if is_adik else _MANDARIN_NUMBERS
    cn_num = lookup.get(num, str(num))
    return base_template.replace("__", cn_num)


# Bangun lookup dict & display list
_PANGGILAN_MAP: dict[str, str] = {}       # display → mandarin (template)
_PANGGILAN_KELUARGA: dict[str, str] = {}  # display → keluarga
_PANGGILAN_DISPLAY: list[str] = []
_NUMBERED_DISPLAY: set[str] = set()       # display strings yang butuh spinner
for _indo, _mandarin, _keluarga in PANGGILAN_OPTIONS:
    # Untuk kategori bernomor, tampilkan tanpa preview mandarin
    if _indo in _NUMBERED_CATEGORIES:
        _display = f"{_indo}  (🔢 pilih nomor)"
        _NUMBERED_DISPLAY.add(_display)
    else:
        _display = f"{_indo}  ({_mandarin})"
    _PANGGILAN_DISPLAY.append(_display)
    _PANGGILAN_MAP[_display] = _mandarin
    _PANGGILAN_KELUARGA[_display] = _keluarga

# Opsi "Isi Manual" — user bisa ketik sendiri aksara Mandarin Panggilan
_PANGGILAN_MANUAL = "(Isi Manual)"
_PANGGILAN_DISPLAY.insert(0, _PANGGILAN_MANUAL)

# ============================================================
# Daftar "Dari" (Hubungan Keluarga)
# ============================================================
# (nama_display, mandarin_laki, mandarin_perempuan)
# Jika mandarin_perempuan kosong "" → item gender-neutral (tanpa toggle L/P)
# Template __ diganti angka Mandarin secara dinamis (butuh spinner)
DARI_OPTIONS: list[tuple[str, str, str]] = [
    # --- Kandung ---
    ("Anak Lk dan Pr",             "众孝眷",             ""),
    ("Anak Ke-",                    "孝__子",           "孝__女"),
    ("Anak Bungsu (Terakhir)",      "孝幼子",           "孝幼女"),
    ("Adik",                        "愚__弟",           "愚__妹"),
    ("Kakak",                       "胞__兄",           "胞__姊"),
    # --- Angkat ---
    ("Anak Angkat",                 "孝養子",           "孝養女"),
    # --- Tiri ---
    ("Anak Tiri",                   "孝繼子",           "孝繼女"),
    # --- Menantu ---
    ("Menantu",                     "孝女婿",           "孝兒媳"),
    # --- Cucu ---
    ("Cucu Kandung",                "孝孫",             "孝孫女"),
    ("Cucu Luar",                   "孝外孫",           "孝外孫女"),
    # --- Keponakan ---
    ("Keponakan (Sdr Laki-laki)",   "孝姪",             "孝姪女"),
    ("Keponakan (Sdr Perempuan)",   "孝外甥",           "孝外甥女"),
    # --- Ipar ---
    ("Saudara Ipar (Adik)",         "孝內弟",           "孝內姊"),
    ("Saudara Ipar (Kakak)",        "内兄",              "内妹"),
    # --- Frasa Umum ---
    ("众孝眷 偕 合家敬奉",          "众孝眷 偕 合家敬奉",     ""),
    ("孝子贤孙 偕 合家敬奉",        "孝子贤孙 偕 合家敬奉",   ""),
    ("合家敬奉 叩首",              "合家敬奉 叩首",         ""),
]

# Kategori Dari yang membutuhkan spinner (nama sebelum generate display)
_DARI_NUMBERED_CATEGORIES: set[str] = {"Anak Ke-", "Adik", "Kakak"}

# Angka Mandarin untuk urutan Anak (beda: 2 = 次)
_DARI_ANAK_NUMBERS: dict[int, str] = {
    1: "長", 2: "次", 3: "三", 4: "四", 5: "五",
    6: "六", 7: "七", 8: "八", 9: "九", 10: "十",
}
# Kakak pakai _MANDARIN_NUMBERS (長, 二, 三 …)
# Adik pakai _MANDARIN_NUMBERS_ADIK (大, 二, 三 …)


def _generate_dari_numbered_mandarin(template: str, num: int,
                                     choice: str) -> str:
    """Generate aksara Mandarin Dari dari template + nomor.

    choice = display string dropdown, dipakai tentukan lookup angka.
    """
    if choice.startswith("Anak"):
        lookup = _DARI_ANAK_NUMBERS
    elif choice.startswith("Adik"):
        lookup = _MANDARIN_NUMBERS_ADIK
    else:  # Kakak
        lookup = _MANDARIN_NUMBERS
    cn_num = lookup.get(num, str(num))
    return template.replace("__", cn_num)


# Bangun lookup dicts untuk DARI
_DARI_MAP_L: dict[str, str] = {}     # display → mandarin laki-laki
_DARI_MAP_P: dict[str, str] = {}     # display → mandarin perempuan
_DARI_DISPLAY: list[str] = []
_DARI_GENDERED: set[str] = set()     # display strings yang punya opsi L/P
_DARI_NUMBERED: set[str] = set()     # display strings yang butuh spinner
for _nama, _laki, _perempuan in DARI_OPTIONS:
    if _nama in _DARI_NUMBERED_CATEGORIES:
        _disp = f"{_nama}  (🔢 pilih nomor)"
        _DARI_NUMBERED.add(_disp)
    elif _perempuan:
        _disp = f"{_nama}  (♂ {_laki} / ♀ {_perempuan})"
    else:
        _disp = f"{_nama}  ({_laki})"
    if _perempuan:
        _DARI_GENDERED.add(_disp)
    _DARI_DISPLAY.append(_disp)
    _DARI_MAP_L[_disp] = _laki
    _DARI_MAP_P[_disp] = _perempuan

# Opsi "Isi Manual" — user bisa ketik sendiri aksara Mandarin Dari
_DARI_MANUAL = "(Isi Manual)"
_DARI_DISPLAY.insert(0, _DARI_MANUAL)

# Display strings dari 3 frasa umum (tanpa suffix 敬奉 叩首)
_DARI_FRASA_UMUM: set[str] = set()
for _nama, _laki, _perempuan in DARI_OPTIONS:
    if _nama in ("众孝眷 偕 合家敬奉", "孝子贤孙 偕 合家敬奉",
                 "合家敬奉 叩首"):
        if _perempuan:
            _d = f"{_nama}  (♂ {_laki} / ♀ {_perempuan})"
        else:
            _d = f"{_nama}  ({_laki})"
        _DARI_FRASA_UMUM.add(_d)


def _reverse_lookup_panggilan(
    mandarin_value: str,
) -> tuple[str | None, int | None]:
    """Reverse lookup aksara Mandarin Panggilan → (display_string, spinner_number).

    Returns:
        (display, None)  untuk item non-bernomor,
        (display, num)   untuk item bernomor,
        (None, None)     jika tidak ditemukan.
    """
    # Coba cocokkan item non-bernomor langsung
    for display, mandarin in _PANGGILAN_MAP.items():
        if display not in _NUMBERED_DISPLAY and mandarin == mandarin_value:
            return (display, None)

    # Coba cocokkan template bernomor
    for display in _NUMBERED_DISPLAY:
        template = _PANGGILAN_MAP[display]
        parts = template.split("__")
        if len(parts) != 2:
            continue
        prefix, suffix = parts
        if not mandarin_value.startswith(prefix):
            continue
        if suffix and not mandarin_value.endswith(suffix):
            continue
        end_idx = len(mandarin_value) - len(suffix) if suffix else len(mandarin_value)
        cn_num = mandarin_value[len(prefix):end_idx]
        if not cn_num:
            continue
        is_adik = "Adik" in display
        lookup = _MANDARIN_NUMBERS_ADIK if is_adik else _MANDARIN_NUMBERS
        for num, char in lookup.items():
            if char == cn_num:
                return (display, num)

    return (None, None)


def _reverse_lookup_dari(
    mandarin_value: str,
) -> tuple[str | None, int | None, str]:
    """Reverse lookup aksara Mandarin Dari → (display_string, spinner_number, gender).

    Suffix ' 敬奉 叩首' otomatis dihapus sebelum pencocokan.

    Returns:
        (display, None, gender) untuk item non-bernomor,
        (display, num, gender)  untuk item bernomor,
        (None, None, "L")       jika tidak ditemukan.
    """
    # Strip suffix "敬奉 叩首" jika ada (disimpan ke DB, perlu dihapus utk lookup)
    clean = mandarin_value.removesuffix(" 敬奉 叩首")

    # Coba cocokkan item non-bernomor langsung
    for display in _DARI_DISPLAY:
        if display in _DARI_NUMBERED:
            continue
        if _DARI_MAP_L.get(display) == clean:
            return (display, None, "L")
        if display in _DARI_GENDERED and _DARI_MAP_P.get(display) == clean:
            return (display, None, "P")

    # Coba cocokkan template bernomor
    for display in _DARI_NUMBERED:
        for gender, dari_map in [("L", _DARI_MAP_L), ("P", _DARI_MAP_P)]:
            template = dari_map.get(display, "")
            if not template or "__" not in template:
                continue
            parts = template.split("__")
            prefix, suffix = parts[0], parts[1]
            if not clean.startswith(prefix):
                continue
            if suffix and not clean.endswith(suffix):
                continue
            end_idx = (
                len(clean) - len(suffix) if suffix
                else len(clean)
            )
            cn_num = clean[len(prefix):end_idx]
            if not cn_num:
                continue
            if display.startswith("Anak"):
                num_lookup = _DARI_ANAK_NUMBERS
            elif display.startswith("Adik"):
                num_lookup = _MANDARIN_NUMBERS_ADIK
            else:
                num_lookup = _MANDARIN_NUMBERS
            for num, char in num_lookup.items():
                if char == cn_num:
                    return (display, num, gender)

    return (None, None, "L")


# ============================================================
# Helper: Right-Click Context Menu (Copy, Cut, Paste)
# ============================================================
def _setup_context_menu(root_widget) -> None:
    """Pasang context menu Copy/Cut/Paste ke semua Entry & Text widget."""

    def _show_context_menu(event):
        widget = event.widget
        menu = tk.Menu(widget, tearoff=0)
        # Cek apakah widget bisa diedit
        is_readonly = False
        try:
            state = str(widget.cget("state"))
            is_readonly = state in ("readonly", "disabled")
        except (tk.TclError, AttributeError):
            pass

        try:
            has_selection = widget.selection_get()
        except (tk.TclError, AttributeError):
            has_selection = ""

        if has_selection:
            menu.add_command(
                label="📋 Copy", accelerator="Ctrl+C",
                command=lambda: widget.event_generate("<<Copy>>"),
            )
        if has_selection and not is_readonly:
            menu.add_command(
                label="✂️ Cut", accelerator="Ctrl+X",
                command=lambda: widget.event_generate("<<Cut>>"),
            )
        if not is_readonly:
            menu.add_command(
                label="📌 Paste", accelerator="Ctrl+V",
                command=lambda: widget.event_generate("<<Paste>>"),
            )
            menu.add_separator()
            menu.add_command(
                label="🔘 Select All", accelerator="Ctrl+A",
                command=lambda: (
                    widget.select_range(0, "end")
                    if isinstance(widget, (tk.Entry, ttk.Entry))
                    else widget.tag_add("sel", "1.0", "end")
                ),
            )

        if menu.index("end") is not None:
            menu.tk_popup(event.x_root, event.y_root)
        menu.grab_release()

    # Bind ke semua entry/text widget via class bindings
    for cls in ("Entry", "TEntry", "Text"):
        root_widget.bind_class(cls, "<Button-3>", _show_context_menu)


# ============================================================
# Widget Kustom: Dropdown dengan scroll
# ============================================================
class ScrollableDropdown(ctk.CTkFrame):
    """Dropdown kustom dengan daftar yang bisa di-scroll via mouse wheel."""

    _PLACEHOLDER = "-- Pilih Panggilan --"

    def __init__(self, master, values: list[str] | None = None,
                 command=None, width: int = 320,
                 font=None, placeholder: str | None = None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._values = values or []
        self._command = command
        self._selected = placeholder or self._PLACEHOLDER
        self._popup: tk.Toplevel | None = None
        self._btn_refs: list[ctk.CTkButton] = []

        # Entry readonly untuk menampilkan pilihan
        self._entry = ctk.CTkEntry(self, width=width - 32, font=font,
                                    state="readonly")
        self._entry.grid(row=0, column=0, sticky="w")
        self._entry.configure(state="normal")
        self._entry.insert(0, self._selected)
        self._entry.configure(state="readonly")

        # Tombol panah
        self._arrow = ctk.CTkButton(
            self, text="▼", width=32, font=ctk.CTkFont(size=11),
            command=self._toggle,
        )
        self._arrow.grid(row=0, column=1, padx=(0, 0), sticky="e")

        # Klik pada entry juga membuka dropdown
        self._entry.bind("<Button-1>", lambda e: self._toggle())

    # -- Public API (compatible dengan CTkComboBox) --
    def get(self) -> str:
        return self._selected

    def set(self, value: str) -> None:
        self._selected = value
        self._entry.configure(state="normal")
        self._entry.delete(0, "end")
        self._entry.insert(0, value)
        self._entry.configure(state="readonly")

    def configure(self, **kwargs):  # noqa: D102
        if "values" in kwargs:
            self._values = kwargs.pop("values")
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        super().configure(**kwargs)

    # -- Internal --
    def _toggle(self) -> None:
        if self._popup and self._popup.winfo_exists():
            self._close()
        else:
            self._open()

    def _open(self) -> None:
        if self._popup and self._popup.winfo_exists():
            return

        # Posisi tepat di bawah entry
        self.update_idletasks()
        x = self._entry.winfo_rootx()
        y = self._entry.winfo_rooty() + self._entry.winfo_height() + 2
        w = self._entry.winfo_width() + 32

        self._popup = tk.Toplevel(self)
        self._popup.wm_overrideredirect(True)
        self._popup.geometry(f"{w}x310+{x}+{y}")
        self._popup.configure(bg="#2b2b2b")

        # Frame scrollable
        sf = ctk.CTkScrollableFrame(self._popup, width=w - 20, height=300,
                                     fg_color=("#f5f5f5", "#2b2b2b"))
        sf.pack(fill="both", expand=True, padx=1, pady=1)

        self._btn_refs.clear()
        for val in self._values:
            btn = ctk.CTkButton(
                sf, text=val, anchor="w",
                height=28, corner_radius=4,
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("#d0d0d0", "#3d3d3d"),
                font=ctk.CTkFont(size=12),
                command=lambda v=val: self._select(v),
            )
            btn.pack(fill="x", padx=2, pady=1)
            self._btn_refs.append(btn)

        # Highlight pilihan saat ini
        self._highlight_current()

        # Tutup saat klik di luar popup
        self._popup.bind("<FocusOut>", self._on_focus_out)
        self._popup.focus_set()

    def _highlight_current(self) -> None:
        for btn in self._btn_refs:
            if btn.cget("text") == self._selected:
                btn.configure(fg_color=("#c8ddf0", "#3a5a7c"))
                break

    def _select(self, value: str) -> None:
        self.set(value)
        self._close()
        if self._command:
            self._command(value)

    def _close(self) -> None:
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None

    def _on_focus_out(self, event) -> None:
        """Tutup popup jika fokus pindah ke luar."""
        if self._popup and self._popup.winfo_exists():
            # Sedikit delay — agar klik tombol di dalam popup masih terbaca
            self._popup.after(120, self._check_focus)

    def _check_focus(self) -> None:
        if self._popup and self._popup.winfo_exists():
            try:
                focused = self._popup.focus_get()
                # Jika fokus sudah bukan di popup → tutup
                if focused is None or not str(focused).startswith(str(self._popup)):
                    self._close()
            except KeyError:
                self._close()


# ============================================================
# Kelas Utama Aplikasi
# ============================================================
class RitualFormApp(ctk.CTk):
    """Aplikasi utama CustomTkinter untuk input formulir ritual."""

    def __init__(self) -> None:
        super().__init__()

        # --- Window setup ---
        self.title(APP_TITLE)
        self.geometry(f"{APP_WIDTH}x{APP_HEIGHT}")
        self.minsize(800, 600)
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # --- Inisialisasi database ---
        try:
            init_db()
        except RuntimeError as e:
            messagebox.showerror("Database Error", str(e))

        # --- State edit ---
        self._editing_uuid: str | None = None

        # --- State bulk-delete (checkbox per order) ---
        self._bulk_check_vars: dict[str, tk.BooleanVar] = {}

        # --- Aktifkan context menu klik kanan (Copy/Cut/Paste) ---
        _setup_context_menu(self)

        # --- Build UI ---
        self._build_input_frame()
        self._build_button_frame()
        self._build_table_frame()

        # --- Load data awal ke tabel ---
        self._refresh_table()

        # --- Cek update di background ---
        self.after(1500, self._check_for_updates)

    # ========================================================
    # Kamus Mini
    # ========================================================
    def _open_dictionary(self) -> None:
        """Buka jendela Kamus Mini Aksara Mandarin."""
        DictionaryWindow(master=self)

    # ========================================================
    # Auto-Update
    # ========================================================
    def _check_for_updates(self) -> None:
        """Cek GitHub Releases untuk versi baru (background thread)."""
        self._updater = UpdateChecker(APP_VERSION)
        self._updater.check_in_background(
            on_update_available=lambda ver, name, notes: self.after(
                0, self._show_update_dialog, ver, name, notes
            ),
        )

    def _show_update_dialog(self, version: str, name: str, notes: str) -> None:
        """Tampilkan dialog notifikasi update."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Update Tersedia")
        dialog.geometry("480x360")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.lift()
        dialog.focus_force()

        # Center dialog
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 480) // 2
        y = self.winfo_y() + (self.winfo_height() - 360) // 2
        dialog.geometry(f"+{x}+{y}")

        # Icon & heading
        ctk.CTkLabel(
            dialog, text="🔄  Update Tersedia!",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(padx=20, pady=(20, 5))

        ctk.CTkLabel(
            dialog,
            text=f"Versi baru: v{version}  (saat ini: v{APP_VERSION})",
            font=ctk.CTkFont(size=13),
            text_color="gray",
        ).pack(padx=20, pady=(0, 5))

        if name:
            ctk.CTkLabel(
                dialog, text=name,
                font=ctk.CTkFont(size=14, weight="bold"),
            ).pack(padx=20, pady=(5, 2))

        # Release notes (scrollable)
        if notes:
            notes_box = ctk.CTkTextbox(dialog, height=120, wrap="word")
            notes_box.pack(padx=20, pady=(5, 10), fill="x")
            notes_box.insert("1.0", notes)
            notes_box.configure(state="disabled")

        # Progress bar (hidden initially)
        progress_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        progress_frame.pack(padx=20, fill="x")
        progress_bar = ctk.CTkProgressBar(progress_frame)
        progress_label = ctk.CTkLabel(
            progress_frame, text="", font=ctk.CTkFont(size=11),
        )

        # Buttons
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(padx=20, pady=(10, 15), fill="x")

        btn_later = ctk.CTkButton(
            btn_frame, text="Nanti", width=120,
            fg_color="gray", hover_color="#666",
            command=dialog.destroy,
        )
        btn_later.pack(side="left", padx=(0, 10))

        def _start_download():
            btn_download.configure(state="disabled", text="Mengunduh...")
            btn_later.configure(state="disabled")
            progress_bar.pack(fill="x", pady=(0, 3))
            progress_bar.set(0)
            progress_label.pack()

            def _on_progress(downloaded, total):
                if total > 0:
                    pct = downloaded / total
                    mb_dl = downloaded / (1024 * 1024)
                    mb_tot = total / (1024 * 1024)
                    self.after(0, lambda: progress_bar.set(pct))
                    self.after(
                        0,
                        lambda d=mb_dl, t=mb_tot: progress_label.configure(
                            text=f"{d:.1f} MB / {t:.1f} MB"
                        ),
                    )

            def _on_done(installer_path):
                self.after(0, lambda: _launch_installer(installer_path))

            def _on_error(msg):
                self.after(
                    0,
                    lambda: (
                        messagebox.showerror("Update Gagal", msg, parent=dialog),
                        btn_download.configure(state="normal", text="Download & Install"),
                        btn_later.configure(state="normal"),
                    ),
                )

            self._updater.download_and_install(
                progress_callback=_on_progress,
                on_done=_on_done,
                on_error=_on_error,
            )

        def _launch_installer(path):
            dialog.destroy()
            messagebox.showinfo(
                "Update",
                "Installer sudah didownload.\n"
                "Aplikasi akan ditutup dan installer akan berjalan.",
            )
            UpdateChecker.launch_installer_and_exit(path)

        btn_download = ctk.CTkButton(
            btn_frame, text="Download & Install", width=180,
            fg_color="#1976D2", hover_color="#1565C0",
            command=_start_download,
        )
        btn_download.pack(side="right")

    # ========================================================
    # Frame Input
    # ========================================================
    def _build_input_frame(self) -> None:
        """Membangun frame form input: info pemesan + detail formulir ritual."""
        master_frame = ctk.CTkFrame(self, corner_radius=10)
        master_frame.pack(padx=15, pady=(15, 5), fill="x", expand=False)

        # ── Header ──
        ctk.CTkLabel(
            master_frame, text="📝 Data Formulir Ritual",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(padx=10, pady=(10, 5), anchor="w")

        # ═══════════════════════════════════════════════════════
        # Bagian 1: Info Pemesan (→ tabel orders)
        # ═══════════════════════════════════════════════════════
        order_frame = ctk.CTkFrame(master_frame, corner_radius=8, fg_color="#E3F2FD")
        order_frame.pack(padx=10, pady=(2, 5), fill="x")

        ctk.CTkLabel(
            order_frame, text="👤 Nama Pemesan",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#1565C0",
        ).grid(row=0, column=0, columnspan=4, padx=10, pady=(8, 3), sticky="w")

        ctk.CTkLabel(
            order_frame, text="Nama:",
            font=ctk.CTkFont(size=12),
        ).grid(row=1, column=0, padx=(10, 5), pady=(2, 8), sticky="e")

        self.entry_nama = ctk.CTkEntry(
            order_frame, width=280,
            placeholder_text="cth: Ajon Bengkel, Rudi Jajan, Lili",
            font=ctk.CTkFont(size=12),
        )
        self.entry_nama.grid(row=1, column=1, padx=5, pady=(2, 8), sticky="w")

        ctk.CTkLabel(
            order_frame,
            text="ℹ Nama yang sama akan otomatis dikelompokkan",
            font=ctk.CTkFont(size=10), text_color="gray50",
        ).grid(row=1, column=2, padx=(15, 10), pady=(2, 8), sticky="w")

        # ═══════════════════════════════════════════════════════
        # Bagian 2: Detail Formulir Ritual (→ tabel ritual_items)
        # ═══════════════════════════════════════════════════════
        item_frame = ctk.CTkFrame(master_frame, corner_radius=8)
        item_frame.pack(padx=10, pady=(2, 10), fill="x")

        ctk.CTkLabel(
            item_frame, text="📋 Detail Formulir",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, columnspan=6, padx=10, pady=(8, 3), sticky="w")

        # --- Baris 1: Panggilan (Dropdown) ---
        ctk.CTkLabel(item_frame, text="Panggilan:").grid(
            row=1, column=0, padx=(10, 5), pady=4, sticky="e"
        )
        self.combo_panggilan = ScrollableDropdown(
            item_frame,
            width=320,
            values=_PANGGILAN_DISPLAY,
            command=self._on_panggilan_selected,
            font=ctk.CTkFont(size=12),
        )
        self.combo_panggilan.grid(row=1, column=1, padx=5, pady=4, sticky="w")

        # Frame untuk spinner nomor urut (tampil/sembunyi dinamis)
        self._spinner_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
        self._spinner_frame.grid(row=1, column=2, padx=(5, 0), pady=4, sticky="w")

        ctk.CTkLabel(
            self._spinner_frame, text="Ke-",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(side="left", padx=(0, 3))

        self._spin_var = tk.IntVar(value=1)
        self._spin_number = tk.Spinbox(
            self._spinner_frame, from_=1, to=10, width=3,
            textvariable=self._spin_var, font=("Segoe UI", 12),
            command=self._on_spin_changed, state="readonly",
        )
        self._spin_number.pack(side="left")
        self._spin_var.trace_add("write", lambda *_: self._on_spin_changed())

        # Awalnya spinner disembunyikan
        self._spinner_frame.grid_remove()

        # --- Entry manual Panggilan (tersembunyi, muncul ketika "Isi Manual" dipilih) ---
        self._panggilan_manual_frame = ctk.CTkFrame(
            item_frame, fg_color="transparent",
        )
        self._entry_panggilan_manual = ctk.CTkEntry(
            self._panggilan_manual_frame, width=200,
            placeholder_text="Ketik aksara Mandarin…",
            font=ctk.CTkFont(family="SimSun", size=13),
        )
        self._entry_panggilan_manual.pack(side="left")
        self._entry_panggilan_manual.bind(
            "<KeyRelease>", self._on_panggilan_manual_typed,
        )
        # Awalnya tersembunyi (belum di-grid)

        # Label preview aksara Mandarin terpilih (merah bold)
        self._lbl_panggilan_mandarin = ctk.CTkLabel(
            item_frame, text="",
            font=ctk.CTkFont(family="SimSun", size=18, weight="bold"),
            text_color="#C62828",
        )
        self._lbl_panggilan_mandarin.grid(row=1, column=3, padx=(10, 5), pady=4, sticky="w")

        # --- Baris 2: Atas Nama Mandarin & Penyebutan ---
        ctk.CTkLabel(item_frame, text="Atas Nama Mandarin:").grid(
            row=2, column=0, padx=(10, 5), pady=4, sticky="e"
        )
        self.entry_mandarin = ctk.CTkEntry(
            item_frame, width=180,
            placeholder_text="cth: 陳瓊雲, 張添蓮",
        )
        self.entry_mandarin.grid(row=2, column=1, padx=5, pady=4, sticky="w")
        # Auto-fill Pinyin saat mengetik aksara Mandarin
        self.entry_mandarin.bind("<KeyRelease>", self._on_mandarin_changed)

        ctk.CTkLabel(item_frame, text="Penyebutan:").grid(
            row=2, column=2, padx=(15, 5), pady=4, sticky="e"
        )
        self.entry_penyebutan = ctk.CTkEntry(
            item_frame, width=180,
            placeholder_text="cth: Chén Qióng Yún",
        )
        self.entry_penyebutan.grid(row=2, column=3, padx=5, pady=4, sticky="w")

        # --- Baris 3: Dari (Dropdown + Gender Toggle) ---
        ctk.CTkLabel(item_frame, text="Dari:").grid(
            row=3, column=0, padx=(10, 5), pady=4, sticky="e"
        )
        self.combo_dari = ScrollableDropdown(
            item_frame,
            width=350,
            values=_DARI_DISPLAY,
            command=self._on_dari_selected,
            font=ctk.CTkFont(size=12),
            placeholder="-- Pilih Dari --",
        )
        self.combo_dari.grid(row=3, column=1, padx=5, pady=4, sticky="w")

        # Container untuk spinner + gender toggle (col 2)
        self._dari_controls_frame = ctk.CTkFrame(
            item_frame, fg_color="transparent", height=1,
        )
        self._dari_controls_frame.grid(row=3, column=2, padx=(5, 0), pady=4, sticky="w")

        # --- Spinner nomor urut Dari (awalnya tersembunyi) ---
        self._dari_spinner_frame = ctk.CTkFrame(
            self._dari_controls_frame, fg_color="transparent",
        )
        ctk.CTkLabel(
            self._dari_spinner_frame, text="Ke-",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(side="left", padx=(0, 3))

        self._dari_spin_var = tk.IntVar(value=1)
        self._dari_spin_number = tk.Spinbox(
            self._dari_spinner_frame, from_=1, to=10, width=3,
            textvariable=self._dari_spin_var, font=("Segoe UI", 12),
            command=self._on_dari_spin_changed, state="readonly",
        )
        self._dari_spin_number.pack(side="left")
        self._dari_spin_var.trace_add(
            "write", lambda *_: self._on_dari_spin_changed(),
        )
        # Spinner awalnya tersembunyi (pack_forget default karena belum di-pack)

        # --- Toggle gender (L/P) — awalnya tersembunyi ---
        self._dari_gender_frame = ctk.CTkFrame(
            self._dari_controls_frame, fg_color="transparent",
        )

        self._dari_gender_var = tk.StringVar(value="L")
        ctk.CTkRadioButton(
            self._dari_gender_frame, text="♂ Laki",
            variable=self._dari_gender_var, value="L",
            font=ctk.CTkFont(size=11),
            command=self._on_dari_gender_changed,
        ).pack(side="left", padx=(0, 5))
        ctk.CTkRadioButton(
            self._dari_gender_frame, text="♀ Perempuan",
            variable=self._dari_gender_var, value="P",
            font=ctk.CTkFont(size=11),
            command=self._on_dari_gender_changed,
        ).pack(side="left")
        # Gender toggle awalnya tersembunyi (pack_forget default)

        # --- Entry manual Dari (tersembunyi, muncul ketika "Isi Manual" dipilih) ---
        self._dari_manual_frame = ctk.CTkFrame(
            self._dari_controls_frame, fg_color="transparent",
        )
        self._entry_dari_manual = ctk.CTkEntry(
            self._dari_manual_frame, width=200,
            placeholder_text="Ketik aksara Mandarin…",
            font=ctk.CTkFont(family="SimSun", size=13),
        )
        self._entry_dari_manual.pack(side="left")
        self._entry_dari_manual.bind("<KeyRelease>", self._on_dari_manual_typed)
        # Awalnya tersembunyi

        # Frame wrapper untuk preview Mandarin + suffix
        self._dari_preview_frame = ctk.CTkFrame(
            item_frame, fg_color="transparent", height=1,
        )
        self._dari_preview_frame.grid(row=3, column=3, padx=(10, 5), pady=4, sticky="w")

        # Preview Mandarin dari pilihan "Dari" (merah bold)
        self._lbl_dari_mandarin = ctk.CTkLabel(
            self._dari_preview_frame, text="",
            font=ctk.CTkFont(family="SimSun", size=16, weight="bold"),
            text_color="#C62828",
        )
        self._lbl_dari_mandarin.pack(side="left")

        # Suffix "敬奉 叩首" (tersembunyi sampai ada pilihan)
        self._lbl_dari_suffix = ctk.CTkLabel(
            self._dari_preview_frame, text="  敬奉 叩首",
            font=ctk.CTkFont(family="SimSun", size=16, weight="bold"),
            text_color="#C62828",
        )
        # Awalnya tersembunyi

        # --- Baris 4: Keterangan ---
        ctk.CTkLabel(item_frame, text="Keterangan:").grid(
            row=4, column=0, padx=(10, 5), pady=(4, 10), sticky="e"
        )
        self.entry_keterangan = ctk.CTkEntry(
            item_frame, width=180,
            placeholder_text="cth: 合家敬奉, 第一第三女",
        )
        self.entry_keterangan.grid(row=4, column=1, padx=5, pady=(4, 10), sticky="w")

    # ========================================================
    # Frame Tombol Aksi
    # ========================================================
    def _build_button_frame(self) -> None:
        """Membangun baris tombol aksi: Simpan, Template, Import, Export."""
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(padx=15, pady=5, fill="x", expand=False)

        # Sub-frame agar tombol Batal muncul di samping Simpan
        self._save_frame = ctk.CTkFrame(frame, fg_color="transparent")
        self._save_frame.pack(side="left", padx=0, pady=0)

        self._btn_save = ctk.CTkButton(
            self._save_frame, text="💾 Simpan Data", width=140,
            command=self._on_save,
        )
        self._btn_save.pack(side="left", padx=8, pady=10)

        self._btn_cancel_edit = ctk.CTkButton(
            self._save_frame, text="❌ Batal", width=100,
            fg_color="#757575", command=self._cancel_edit,
        )
        # Awalnya tersembunyi — tampil saat mode edit


        ctk.CTkButton(
            frame, text="� Template Excel", width=140,
            fg_color="#7B1FA2", command=self._on_generate_template,
        ).pack(side="left", padx=8, pady=10)

        ctk.CTkButton(
            frame, text="�📂 Import Excel", width=140, fg_color="#E67E22", command=self._on_import_excel
        ).pack(side="left", padx=8, pady=10)

        ctk.CTkButton(
            frame, text="� Export Backup", width=140,
            fg_color="#00695C", command=self._on_export_backup,
        ).pack(side="left", padx=8, pady=10)

        ctk.CTkButton(
            frame, text="�🔄 Refresh", width=120, command=self._refresh_table
        ).pack(side="right", padx=8, pady=10)

        ctk.CTkButton(
            frame, text="📖 Kamus", width=120,
            fg_color="#1565C0", hover_color="#0D47A1",
            command=self._open_dictionary,
        ).pack(side="right", padx=8, pady=10)

    # ========================================================
    # Tabel Preview (Treeview)
    # ========================================================
    def _build_table_frame(self) -> None:
        """Membangun tabel preview data menggunakan ttk.Treeview (grouped per pemesan)."""
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(padx=15, pady=(5, 15), fill="both", expand=True)

        ctk.CTkLabel(
            frame, text="📋 Data Tersimpan (Grouped per Pemesan)",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(padx=10, pady=(10, 5), anchor="w")

        # --- Toolbar: Pencarian + Expand/Collapse ---
        toolbar = ctk.CTkFrame(frame, fg_color="transparent")
        toolbar.pack(fill="x", padx=10, pady=(0, 5))

        # Search icon + entry
        ctk.CTkLabel(
            toolbar, text="🔍", font=ctk.CTkFont(size=14),
        ).pack(side="left", padx=(0, 4))
        self._search_var = tk.StringVar()
        self._search_entry = ctk.CTkEntry(
            toolbar, width=250,
            placeholder_text="Cari nama pemesan…",
            textvariable=self._search_var,
        )
        self._search_entry.pack(side="left", padx=(0, 8))
        self._search_var.trace_add("write", lambda *_: self._on_search_changed())

        # Tombol Expand All / Collapse All
        self._btn_expand_all = ctk.CTkButton(
            toolbar, text="▶ Buka Semua", width=120,
            fg_color="#1976D2", hover_color="#1565C0",
            command=self._expand_all,
        )
        self._btn_expand_all.pack(side="right", padx=(4, 0))

        self._btn_collapse_all = ctk.CTkButton(
            toolbar, text="▼ Tutup Semua", width=120,
            fg_color="#757575", hover_color="#616161",
            command=self._collapse_all,
        )
        self._btn_collapse_all.pack(side="right", padx=(4, 0))

        # Tombol Bulk Delete
        self._btn_bulk_delete = ctk.CTkButton(
            toolbar, text="🗑️ Hapus Terpilih", width=140,
            fg_color="#E53935", hover_color="#B71C1C",
            command=self._on_bulk_delete,
        )
        self._btn_bulk_delete.pack(side="right", padx=(4, 0))

        # Kolom tabel
        columns = (
            "uuid", "panggilan", "mandarin",
            "penyebutan", "dari", "keluarga", "keterangan",
            "tanggal", "aksi",
        )
        self.tree = ttk.Treeview(frame, columns=columns, show="tree headings", height=12)

        # Kolom tree (nama pemesan sebagai parent)
        self.tree.heading("#0", text="Nama Pemesan")
        self.tree.column("#0", width=160, anchor="w")

        # Header kolom
        self.tree.heading("uuid", text="ID")
        self.tree.heading("panggilan", text="Panggilan")
        self.tree.heading("mandarin", text="Mandarin")
        self.tree.heading("penyebutan", text="Penyebutan")
        self.tree.heading("dari", text="Dari")
        self.tree.heading("keluarga", text="Keluarga")
        self.tree.heading("keterangan", text="Keterangan")
        self.tree.heading("tanggal", text="Dibuat")
        self.tree.heading("aksi", text="Aksi")

        # Lebar kolom
        self.tree.column("uuid", width=65, anchor="center")
        self.tree.column("panggilan", width=100, anchor="center")
        self.tree.column("mandarin", width=100, anchor="center")
        self.tree.column("penyebutan", width=100, anchor="center")
        self.tree.column("dari", width=90, anchor="center")
        self.tree.column("keluarga", width=90, anchor="center")
        self.tree.column("keterangan", width=80, anchor="center")
        self.tree.column("tanggal", width=110, anchor="center")
        self.tree.column("aksi", width=100, anchor="center")

        # Style untuk parent (nama pemesan)
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)

        # Tag warna untuk parent rows
        self.tree.tag_configure("order", background="#E3F2FD", font=("Segoe UI", 10, "bold"))
        self.tree.tag_configure("item_even", background="#FFFFFF")
        self.tree.tag_configure("item_odd", background="#F5F5F5")

        # Scrollbar vertikal
        scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_y.set)

        # Scrollbar horizontal
        scrollbar_x = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=scrollbar_x.set)

        self.tree.pack(side="top", fill="both", expand=True, padx=(10, 0), pady=(0, 0))
        scrollbar_y.pack(side="right", fill="y", padx=(0, 10), pady=(0, 0))
        scrollbar_x.pack(side="bottom", fill="x", padx=10, pady=(0, 10))

        # Klik pada kolom Aksi untuk edit/hapus/cetak per-baris
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)

    # ========================================================
    # Event Handlers
    # ========================================================
    def _on_tree_click(self, event) -> None:
        """Handler klik pada Treeview: deteksi klik kolom Aksi untuk edit/hapus/cetak."""
        # Identifikasi baris & kolom yang diklik
        row_iid = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)

        if not row_iid or not col_id:
            return

        # Klik pada parent (order group → nama pemesan)
        if str(row_iid).startswith("order_"):
            # Klik di kolom #0 (tree column / nama pemesan) → rename
            if col_id == "#0":
                self._on_rename_order(row_iid)
            return

        # Kolom aksi = kolom terakhir (#9 karena #0 = tree column)
        if col_id != "#9":
            return

        # Pilih baris yang diklik
        self.tree.selection_set(row_iid)

        # Tentukan aksi berdasarkan posisi x dalam kolom
        col_bbox = self.tree.bbox(row_iid, column="aksi")
        if not col_bbox:
            return

        col_x, _col_y, col_w, _col_h = col_bbox
        rel_x = event.x - col_x  # posisi x relatif dalam kolom

        # Bagi kolom jadi 3 zona: Edit | Hapus | Cetak
        zone_width = col_w / 3
        if rel_x < zone_width:
            self._on_edit()
        elif rel_x < zone_width * 2:
            self._on_delete()
        else:
            self._on_print()

    def _on_rename_order(self, row_iid: str) -> None:
        """Handler rename nama pemesan: Tampilkan dialog untuk mengubah nama."""
        order_uuid = str(row_iid).replace("order_", "")

        # Ambil nama saat ini dari treeview
        item = self.tree.item(row_iid)
        current_text = item.get("text", "")
        current_name = (
            current_text
            .replace("\U0001f4c1 ", "")
            .replace("\u270f\ufe0f ", "")
            .split("  (")[0]
            .strip()
        )

        # Dialog input nama baru
        dialog = ctk.CTkInputDialog(
            title="Ubah Nama Pemesan",
            text=f"Nama saat ini: {current_name}\n\nMasukkan nama baru:",
        )
        new_name = dialog.get_input()

        if not new_name or not new_name.strip():
            return
        new_name = new_name.strip()

        if new_name == current_name:
            return

        try:
            result = update_order_name(order_uuid, new_name)
            if result.get("merged"):
                messagebox.showinfo(
                    "Berhasil (Digabung)",
                    f"Nama pemesan diubah ke '{new_name}'.\n"
                    f"Item digabungkan dengan grup yang sudah ada.",
                )
            else:
                messagebox.showinfo(
                    "Berhasil",
                    f"Nama pemesan berhasil diubah menjadi '{new_name}'.",
                )
            self._refresh_table()
        except RuntimeError as e:
            messagebox.showerror("Gagal Mengubah Nama", str(e))

    def _on_bulk_delete(self) -> None:
        """Handler tombol Hapus Terpilih: Hapus semua grup pemesan yang terpilih."""
        # Ambil semua order yang terpilih di Treeview
        selected = self.tree.selection()
        order_uuids = []
        order_names = []

        for iid in selected:
            if str(iid).startswith("order_"):
                uuid = str(iid).replace("order_", "")
                order_uuids.append(uuid)
                # Ambil nama untuk konfirmasi
                item_text = self.tree.item(iid).get("text", "")
                name = (
                    item_text
                    .replace("\U0001f4c1 ", "")
                    .replace("\u270f\ufe0f ", "")
                    .split("  (")[0]
                    .strip()
                )
                order_names.append(name)

        if not order_uuids:
            messagebox.showwarning(
                "Tidak Ada Pilihan",
                "Pilih satu atau lebih grup pemesan (baris 📁) di tabel.\n\n"
                "Tips: Tahan Ctrl dan klik beberapa grup pemesan\n"
                "untuk memilih lebih dari satu.",
            )
            return

        # Konfirmasi
        names_list = "\n".join(f"  • {n}" for n in order_names)
        confirm = messagebox.askyesno(
            "Konfirmasi Hapus Massal",
            f"Apakah Anda yakin ingin menghapus {len(order_uuids)} grup pemesan?\n\n"
            f"{names_list}\n\n"
            f"Semua item di dalamnya juga akan dihapus!\n"
            f"Tindakan ini tidak dapat dibatalkan.",
        )
        if not confirm:
            return

        try:
            deleted = bulk_delete_orders(order_uuids)
            messagebox.showinfo(
                "Berhasil",
                f"{deleted} grup pemesan beserta semua item-nya berhasil dihapus.",
            )
            self._refresh_table()
        except RuntimeError as e:
            messagebox.showerror("Gagal Menghapus", str(e))

    def _on_save(self) -> None:
        """Handler tombol Simpan/Update: Validasi input lalu simpan ke database."""
        panggilan = self._get_panggilan_mandarin()
        keluarga = self._get_panggilan_keluarga()
        mandarin = self.entry_mandarin.get().strip()
        dari = self._get_dari_mandarin()
        nama = self.entry_nama.get().strip()
        penyebutan = self.entry_penyebutan.get().strip()
        keterangan = self.entry_keterangan.get().strip()

        # Validasi: Field wajib tidak boleh kosong
        if not all([panggilan, mandarin, dari]):
            messagebox.showwarning(
                "Input Tidak Lengkap",
                "Harap isi field wajib:\n"
                "• Panggilan (稱呼)\n• Nama Mandarin\n• Dari (陽上)",
            )
            return

        try:
            if self._editing_uuid:
                # Mode Edit: Update record yang ada
                success = update_record(
                    record_uuid=self._editing_uuid,
                    panggilan=panggilan,
                    nama_mandarin=mandarin,
                    dari=dari,
                    nama=nama,
                    penyebutan=penyebutan,
                    keluarga=keluarga,
                    keterangan=keterangan,
                )
                if success:
                    messagebox.showinfo("Berhasil", "Data berhasil diupdate.")
                else:
                    messagebox.showwarning("Gagal", "Record tidak ditemukan.")
                self._editing_uuid = None
                self._btn_save.configure(text="💾 Simpan Data")
                self._btn_cancel_edit.pack_forget()
            else:
                # Mode Baru: Insert record baru
                record_uuid = insert_record(
                    panggilan=panggilan,
                    nama_mandarin=mandarin,
                    dari=dari,
                    nama=nama,
                    penyebutan=penyebutan,
                    keluarga=keluarga,
                    keterangan=keterangan,
                )
                messagebox.showinfo("Berhasil", f"Data tersimpan.\nUUID: {record_uuid[:8]}...")
            self._clear_inputs()
            self._refresh_table()
        except RuntimeError as e:
            messagebox.showerror("Gagal Menyimpan", str(e))

    def _on_print(self) -> None:
        """Handler tombol Cetak: Preview PDF dan cetak langsung ke printer."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Tidak Ada Pilihan", "Pilih satu baris item di tabel untuk dicetak.")
            return

        sel_iid = selected[0]

        # Jika user memilih parent (order), tampilkan pesan
        if str(sel_iid).startswith("order_"):
            messagebox.showwarning(
                "Pilih Item",
                "Anda memilih grup pemesan.\n"
                "Silakan klik salah satu item (child) di bawahnya untuk dicetak.",
            )
            return

        # Ambil data dari baris terpilih (child item)
        item = self.tree.item(sel_iid)
        values = item["values"]
        # Kolom: 0=uuid(short), 1=panggilan, 2=mandarin, 3=penyebutan,
        #         4=dari, 5=keluarga, 6=keterangan, 7=tanggal

        # Ambil nama pemesan dari parent
        parent_iid = self.tree.parent(sel_iid)
        parent_item = self.tree.item(parent_iid)
        parent_text = parent_item.get("text", "")
        nama_pemesan = parent_text.replace("\U0001f4c1 ", "").replace("✏️ ", "").split("  (")[0].strip()

        data = {
            "nama": nama_pemesan,
            "panggilan": values[1],
            "nama_mandarin": values[2],
            "penyebutan": values[3],
            "dari": values[4],
            "keluarga": values[5],
            "keterangan": values[6],
        }

        # Offset default (kalibrasi belum diaktifkan)
        offset_x, offset_y = 0.0, 0.0

        try:
            pdf_bytes = generate_pdf_bytes(
                data, offset_x=offset_x, offset_y=offset_y,
            )
            title = f"Preview — {values[1]} {values[2]}"
            PDFPreviewWindow(self, pdf_bytes=pdf_bytes, title_text=title)
        except FileNotFoundError as e:
            messagebox.showerror("Font Tidak Ditemukan", str(e))
        except RuntimeError as e:
            messagebox.showerror("Gagal Membuat PDF", str(e))

    def _on_delete(self) -> None:
        """Handler tombol Hapus: Hapus item atau seluruh grup pemesan terpilih."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Tidak Ada Pilihan", "Pilih satu baris data di tabel untuk dihapus.")
            return

        sel_iid = selected[0]

        # Jika memilih parent (order group) → hapus seluruh grup
        if str(sel_iid).startswith("order_"):
            order_uuid = str(sel_iid).replace("order_", "")
            parent_item = self.tree.item(sel_iid)
            parent_text = parent_item.get("text", "")

            confirm = messagebox.askyesno(
                "Konfirmasi Hapus Grup",
                f"Apakah Anda yakin ingin menghapus seluruh grup:\n"
                f"{parent_text}\n\n"
                f"Semua item di dalamnya juga akan dihapus!",
            )
            if not confirm:
                return

            try:
                success = delete_order(order_uuid)
                if success:
                    messagebox.showinfo("Berhasil", "Grup pemesan dan semua item-nya berhasil dihapus.")
                    self._refresh_table()
                else:
                    messagebox.showwarning("Tidak Ditemukan", "Grup tidak ditemukan di database.")
            except RuntimeError as e:
                messagebox.showerror("Gagal Menghapus", str(e))
            return

        # Memilih child (item) → hapus satu item
        item = self.tree.item(sel_iid)
        short_uuid = item["values"][0]

        # Cari full UUID dari database
        try:
            records = get_all_records()
            full_uuid = None
            for r in records:
                if r["uuid"].startswith(str(short_uuid)):
                    full_uuid = r["uuid"]
                    break
            if not full_uuid:
                messagebox.showwarning("Tidak Ditemukan", "Record tidak ditemukan di database.")
                return
        except RuntimeError as e:
            messagebox.showerror("Gagal", str(e))
            return

        confirm = messagebox.askyesno(
            "Konfirmasi Hapus",
            f"Apakah Anda yakin ingin menghapus item:\n"
            f"UUID: {short_uuid}... ?\n"
            f"Mandarin: {item['values'][2]}",
        )
        if not confirm:
            return

        try:
            success = delete_record(full_uuid)
            if success:
                messagebox.showinfo("Berhasil", "Record berhasil dihapus.")
                self._refresh_table()
            else:
                messagebox.showwarning("Tidak Ditemukan", "Record tidak ditemukan di database.")
        except RuntimeError as e:
            messagebox.showerror("Gagal Menghapus", str(e))

    def _on_edit(self) -> None:
        """Handler tombol Edit: Muat data terpilih ke form untuk diedit."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning(
                "Tidak Ada Pilihan",
                "Pilih satu baris item di tabel untuk diedit.",
            )
            return

        sel_iid = selected[0]

        # Jika user memilih parent (order), tampilkan pesan
        if str(sel_iid).startswith("order_"):
            messagebox.showwarning(
                "Pilih Item",
                "Anda memilih grup pemesan.\n"
                "Silakan klik salah satu item (child) di bawahnya untuk diedit.",
            )
            return

        # Ambil data dari baris terpilih
        item = self.tree.item(sel_iid)
        values = item["values"]
        short_uuid = str(values[0])

        # Cari full UUID dari database
        full_uuid = self._find_full_uuid(short_uuid)
        if not full_uuid:
            messagebox.showwarning(
                "Tidak Ditemukan", "Record tidak ditemukan di database.",
            )
            return

        # Ambil nama pemesan dari parent
        parent_iid = self.tree.parent(sel_iid)
        parent_item = self.tree.item(parent_iid)
        parent_text = parent_item.get("text", "")
        nama_pemesan = parent_text.replace("📁 ", "").replace("✏️ ", "").split("  (")[0].strip()

        # Muat data ke form
        self._load_record_to_form(
            full_uuid=full_uuid,
            nama=nama_pemesan,
            panggilan_mandarin=str(values[1]),
            mandarin=str(values[2]),
            penyebutan=str(values[3]),
            dari_mandarin=str(values[4]),
            keterangan=str(values[6]),
        )

    def _load_record_to_form(
        self,
        full_uuid: str,
        nama: str,
        panggilan_mandarin: str,
        mandarin: str,
        penyebutan: str,
        dari_mandarin: str,
        keterangan: str,
    ) -> None:
        """Muat data record ke form untuk diedit (mode edit)."""
        # Aktifkan mode edit
        self._editing_uuid = full_uuid
        self._btn_save.configure(text="💾 Update Data")
        self._btn_cancel_edit.pack(side="left", padx=(0, 8), pady=10)

        # Bersihkan form dulu
        self._clear_inputs()

        # Isi Nama Pemesan
        self.entry_nama.insert(0, nama)

        # Isi Panggilan dropdown (reverse lookup)
        panggilan_disp, panggilan_num = _reverse_lookup_panggilan(
            panggilan_mandarin,
        )
        if panggilan_disp:
            self.combo_panggilan.set(panggilan_disp)
            if panggilan_num is not None:
                self._spin_var.set(panggilan_num)
            self._on_panggilan_selected(panggilan_disp)
        else:
            # Fallback: gunakan mode manual dan isi entry dengan value asli
            self.combo_panggilan.set(_PANGGILAN_MANUAL)
            self._on_panggilan_selected(_PANGGILAN_MANUAL)
            self._entry_panggilan_manual.insert(0, panggilan_mandarin)
            self._on_panggilan_manual_typed()

        # Isi Nama Mandarin (tanpa trigger auto-pinyin)
        self.entry_mandarin.insert(0, mandarin)

        # Isi Penyebutan
        self.entry_penyebutan.insert(0, penyebutan)

        # Isi Dari dropdown (reverse lookup)
        dari_disp, dari_num, dari_gender = _reverse_lookup_dari(dari_mandarin)
        if dari_disp:
            self._dari_gender_var.set(dari_gender)
            self.combo_dari.set(dari_disp)
            if dari_num is not None:
                self._dari_spin_var.set(dari_num)
            self._on_dari_selected(dari_disp)
        else:
            # Fallback: gunakan mode manual dan isi entry dengan value asli
            clean = dari_mandarin.replace("敬奉", "").replace("叩首", "").strip()
            self.combo_dari.set(_DARI_MANUAL)
            self._on_dari_selected(_DARI_MANUAL)
            self._entry_dari_manual.insert(0, clean)
            self._on_dari_manual_typed()

        # Isi Keterangan
        self.entry_keterangan.insert(0, keterangan)

    def _cancel_edit(self) -> None:
        """Batal mode edit, kembalikan form ke mode input baru."""
        self._editing_uuid = None
        self._btn_save.configure(text="💾 Simpan Data")
        self._btn_cancel_edit.pack_forget()
        self._clear_inputs()

    def _find_full_uuid(self, short_uuid: str) -> str | None:
        """Cari full UUID dari prefix pendek (8 karakter pertama)."""
        try:
            records = get_all_records()
            for r in records:
                if r["uuid"].startswith(short_uuid):
                    return r["uuid"]
        except RuntimeError:
            pass
        return None


    # ========================================================
    # Utility Methods
    # ========================================================
    def _refresh_table(self) -> None:
        """Memuat ulang seluruh data dari database ke tabel Treeview (grouped)."""
        search_text = ""
        if hasattr(self, "_search_var"):
            search_text = self._search_var.get().strip().lower()
        self._populate_table(search_text)

    def _populate_table(self, search_text: str = "") -> None:
        """Isi tabel Treeview, opsional difilter berdasarkan nama pemesan / nama mandarin."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Reset bulk-delete checkboxes
        self._bulk_check_vars.clear()

        try:
            orders = get_all_orders()
            for order in orders:
                display_name = order["nama"] or "(tanpa nama)"

                # Cek apakah nama pemesan cocok dengan pencarian
                name_match = (not search_text
                              or search_text in display_name.lower())

                # Ambil items untuk cek nama mandarin jika nama pemesan tidak cocok
                items = get_items_by_order(order["uuid"])

                if not name_match:
                    # Cek apakah ada item yang nama mandarin-nya cocok
                    has_mandarin_match = any(
                        search_text in (r.get("nama_mandarin", "") or "").lower()
                        for r in items
                    )
                    if not has_mandarin_match:
                        continue

                # Parent row = nama pemesan (with edit icon)
                order_iid = f"order_{order['uuid']}"
                item_count = order["item_count"]
                self.tree.insert(
                    "",
                    "end",
                    iid=order_iid,
                    text=f"📁 ✏️ {display_name}  ({item_count} item)",
                    values=("", "", "", "", "", "", "", "", ""),
                    open=True,
                    tags=("order",),
                )

                # Child rows = ritual items
                for idx, r in enumerate(items):
                    # Jika pencarian aktif dan nama pemesan tidak cocok,
                    # hanya tampilkan item yang nama mandarin-nya cocok
                    if search_text and not name_match:
                        mandarin = (r.get("nama_mandarin", "") or "").lower()
                        if search_text not in mandarin:
                            continue

                    tag = "item_even" if idx % 2 == 0 else "item_odd"
                    self.tree.insert(
                        order_iid,
                        "end",
                        values=(
                            r["uuid"][:8],
                            r["panggilan"],
                            r["nama_mandarin"],
                            r["penyebutan"],
                            r["dari"],
                            r["keluarga"],
                            r["keterangan"],
                            r["created_at"],
                            "✏️  🗑️  🖨️",
                        ),
                        tags=(tag,),
                    )
        except RuntimeError as e:
            messagebox.showerror("Gagal Memuat Data", str(e))

    def _on_search_changed(self) -> None:
        """Callback saat teks pencarian berubah — filter tabel secara real-time."""
        search_text = self._search_var.get().strip().lower()
        self._populate_table(search_text)

    def _expand_all(self) -> None:
        """Buka (expand) semua grup pemesan di tabel."""
        for iid in self.tree.get_children():
            self.tree.item(iid, open=True)

    def _collapse_all(self) -> None:
        """Tutup (collapse) semua grup pemesan di tabel."""
        for iid in self.tree.get_children():
            self.tree.item(iid, open=False)

    def _is_numbered_choice(self, choice: str) -> bool:
        """Cek apakah pilihan dropdown butuh spinner nomor."""
        return choice in _NUMBERED_DISPLAY

    def _is_adik_choice(self, choice: str) -> bool:
        """Cek apakah pilihan dropdown termasuk kategori Adik."""
        return choice.startswith("Adik ")

    def _on_panggilan_selected(self, choice: str) -> None:
        """Callback saat user memilih panggilan dari dropdown."""
        is_manual = (choice == _PANGGILAN_MANUAL)

        # Tampilkan/sembunyikan entry manual
        if is_manual:
            self._panggilan_manual_frame.grid(
                row=1, column=2, padx=(5, 0), pady=4, sticky="w",
            )
            self._spinner_frame.grid_remove()
            self._entry_panggilan_manual.focus_set()
            self._on_panggilan_manual_typed()
        elif self._is_numbered_choice(choice):
            # Tampilkan spinner & generate mandarin dari nomor
            self._panggilan_manual_frame.grid_remove()
            self._spinner_frame.grid()
            self._update_numbered_preview(choice)
        else:
            # Sembunyikan spinner & entry manual, tampilkan mandarin biasa
            self._panggilan_manual_frame.grid_remove()
            self._spinner_frame.grid_remove()
            mandarin = _PANGGILAN_MAP.get(choice, "")
            self._lbl_panggilan_mandarin.configure(
                text=mandarin if mandarin else ""
            )

    def _on_panggilan_manual_typed(self, event=None) -> None:
        """Callback saat user mengetik di entry manual Panggilan."""
        text = self._entry_panggilan_manual.get().strip()
        self._lbl_panggilan_mandarin.configure(text=text)

    def _on_spin_changed(self) -> None:
        """Callback saat spinner nomor berubah."""
        choice = self.combo_panggilan.get()
        if self._is_numbered_choice(choice):
            self._update_numbered_preview(choice)

    def _update_numbered_preview(self, choice: str) -> None:
        """Update preview mandarin sesuai template + nomor spinner."""
        template = _PANGGILAN_MAP.get(choice, "")
        try:
            num = self._spin_var.get()
        except (tk.TclError, ValueError):
            num = 1
        is_adik = self._is_adik_choice(choice)
        mandarin = _generate_numbered_mandarin(template, num, is_adik=is_adik)
        self._lbl_panggilan_mandarin.configure(text=mandarin)

    def _get_panggilan_mandarin(self) -> str:
        """Ambil aksara Mandarin dari pilihan panggilan dropdown."""
        choice = self.combo_panggilan.get()

        # Jika mode manual, ambil dari entry manual
        if choice == _PANGGILAN_MANUAL:
            return self._entry_panggilan_manual.get().strip()

        template = _PANGGILAN_MAP.get(choice, choice.strip())
        if self._is_numbered_choice(choice):
            try:
                num = self._spin_var.get()
            except (tk.TclError, ValueError):
                num = 1
            is_adik = self._is_adik_choice(choice)
            return _generate_numbered_mandarin(template, num, is_adik=is_adik)
        return template

    def _get_panggilan_keluarga(self) -> str:
        """Ambil keluarga otomatis dari pilihan panggilan dropdown."""
        choice = self.combo_panggilan.get()
        if choice == _PANGGILAN_MANUAL:
            return ""  # Manual: user isi keluarga sendiri via Keterangan
        base = _PANGGILAN_KELUARGA.get(choice, "")
        if self._is_numbered_choice(choice):
            try:
                num = self._spin_var.get()
            except (tk.TclError, ValueError):
                num = 1
            return f"{base} {num}"
        return base

    def _on_mandarin_changed(self, event=None) -> None:
        """Auto-fill Pinyin ke field Penyebutan saat aksara Mandarin diketik."""
        text = self.entry_mandarin.get().strip()
        if not text:
            self.entry_penyebutan.delete(0, "end")
            return

        # Konversi aksara Mandarin ke Pinyin (huruf pertama kapital per suku kata)
        try:
            syllables = pinyin(text, style=PinyinStyle.TONE)
            result = " ".join(s[0].capitalize() for s in syllables if s[0])
        except Exception:  # noqa: BLE001
            result = ""

        # Isi otomatis ke Penyebutan (user masih bisa edit manual)
        self.entry_penyebutan.delete(0, "end")
        self.entry_penyebutan.insert(0, result)

    # -- Dari dropdown handlers --
    def _on_dari_selected(self, choice: str) -> None:
        """Callback saat user memilih dari dropdown Dari."""
        is_manual = (choice == _DARI_MANUAL)
        is_numbered = choice in _DARI_NUMBERED
        is_gendered = choice in _DARI_GENDERED

        # Tampilkan/sembunyikan entry manual
        if is_manual:
            self._dari_manual_frame.pack(side="left", padx=(0, 10))
            self._entry_dari_manual.focus_set()
        else:
            self._dari_manual_frame.pack_forget()

        # Tampilkan/sembunyikan spinner
        if is_numbered:
            self._dari_spinner_frame.pack(side="left", padx=(0, 10))
        else:
            self._dari_spinner_frame.pack_forget()

        # Tampilkan/sembunyikan gender toggle
        if is_gendered:
            self._dari_gender_frame.pack(side="left")
        else:
            self._dari_gender_frame.pack_forget()

        self._update_dari_preview(choice)

    def _on_dari_spin_changed(self) -> None:
        """Callback saat spinner nomor Dari berubah."""
        choice = self.combo_dari.get()
        if choice in _DARI_NUMBERED:
            self._update_dari_preview(choice)

    def _on_dari_gender_changed(self) -> None:
        """Callback saat toggle gender Dari berubah."""
        choice = self.combo_dari.get()
        self._update_dari_preview(choice)

    def _on_dari_manual_typed(self, event=None) -> None:
        """Callback saat user mengetik di entry manual Dari."""
        text = self._entry_dari_manual.get().strip()
        self._lbl_dari_mandarin.configure(text=text)
        # Mode manual: tidak tampilkan suffix 敬奉 叩首
        self._lbl_dari_suffix.pack_forget()

    def _update_dari_preview(self, choice: str) -> None:
        """Update preview mandarin Dari sesuai pilihan + gender + nomor."""
        # Jika manual, preview diatur oleh _on_dari_manual_typed
        if choice == _DARI_MANUAL:
            self._on_dari_manual_typed()
            return

        if choice in _DARI_GENDERED and self._dari_gender_var.get() == "P":
            mandarin = _DARI_MAP_P.get(choice, "")
        else:
            mandarin = _DARI_MAP_L.get(choice, "")

        # Jika item bernomor, generate dari template + spinner
        if choice in _DARI_NUMBERED:
            try:
                num = self._dari_spin_var.get()
            except (tk.TclError, ValueError):
                num = 1
            mandarin = _generate_dari_numbered_mandarin(
                mandarin, num, choice,
            )

        self._lbl_dari_mandarin.configure(text=mandarin)

        # Tampilkan/sembunyikan suffix "敬奉 叩首"
        if mandarin and choice not in _DARI_FRASA_UMUM:
            self._lbl_dari_suffix.pack(side="left")
        else:
            self._lbl_dari_suffix.pack_forget()

    def _get_dari_mandarin(self) -> str:
        """Ambil aksara Mandarin Dari dari dropdown + gender + nomor + suffix."""
        choice = self.combo_dari.get()

        # Jika mode manual, ambil dari entry manual (tanpa suffix)
        if choice == _DARI_MANUAL:
            return self._entry_dari_manual.get().strip()

        if choice in _DARI_GENDERED and self._dari_gender_var.get() == "P":
            mandarin = _DARI_MAP_P.get(choice, "")
        else:
            mandarin = _DARI_MAP_L.get(choice, choice.strip())

        if choice in _DARI_NUMBERED:
            try:
                num = self._dari_spin_var.get()
            except (tk.TclError, ValueError):
                num = 1
            mandarin = _generate_dari_numbered_mandarin(
                mandarin, num, choice,
            )

        # Tambahkan suffix "敬奉 叩首" kecuali untuk frasa umum
        if mandarin and choice not in _DARI_FRASA_UMUM:
            mandarin = f"{mandarin} 敬奉 叩首"

        return mandarin

    def _clear_inputs(self) -> None:
        """Mengosongkan semua field input setelah simpan berhasil."""
        self.entry_nama.delete(0, "end")
        self.combo_panggilan.set("-- Pilih Panggilan --")
        self._lbl_panggilan_mandarin.configure(text="")
        self._spinner_frame.grid_remove()
        self._panggilan_manual_frame.grid_remove()
        self._entry_panggilan_manual.delete(0, "end")
        self._spin_var.set(1)
        self.entry_mandarin.delete(0, "end")
        self.entry_penyebutan.delete(0, "end")
        self.combo_dari.set("-- Pilih Dari --")
        self._lbl_dari_mandarin.configure(text="")
        self._lbl_dari_suffix.pack_forget()
        self._dari_spinner_frame.pack_forget()
        self._dari_spin_var.set(1)
        self._dari_gender_frame.pack_forget()
        self._dari_gender_var.set("L")
        self._dari_manual_frame.pack_forget()
        self._entry_dari_manual.delete(0, "end")
        self.entry_keterangan.delete(0, "end")

    def _on_import_excel(self) -> None:
        """Handler tombol Import Excel: Pilih file .xlsx, import ke DB."""
        filepath = filedialog.askopenfilename(
            title="Pilih File Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )
        if not filepath:
            return

        # Bersihkan memori sebelum import (mencegah crash setelah lama idle)
        import gc
        gc.collect()

        try:
            count = import_from_excel(filepath)
            messagebox.showinfo(
                "Import Berhasil",
                f"Berhasil mengimport {count} record dari:\n{os.path.basename(filepath)}",
            )
            self._refresh_table()
        except ImportError:
            messagebox.showerror(
                "Module Tidak Ditemukan",
                "Module openpyxl belum terinstal.\n"
                "Jalankan: pip install openpyxl",
            )
        except RuntimeError as e:
            messagebox.showerror("Gagal Import", str(e))
        except Exception as e:
            messagebox.showerror(
                "Gagal Import",
                f"Terjadi kesalahan tak terduga:\n{type(e).__name__}: {e}\n\n"
                "Coba restart aplikasi dan ulangi import.",
            )
        finally:
            gc.collect()

    def _on_generate_template(self) -> None:
        """Handler tombol Template: Generate file template Excel import."""
        from modules.excel_template import generate_template

        filepath = filedialog.asksaveasfilename(
            title="Simpan Template Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="template_import_ritual.xlsx",
        )
        if not filepath:
            return

        try:
            result = generate_template(filepath)
            messagebox.showinfo(
                "Template Berhasil",
                f"Template Excel berhasil dibuat:\n{result}\n\n"
                "Isi kolom Panggilan & Dari menggunakan dropdown.\n"
                "Saat import, otomatis dikonversi ke aksara Mandarin.",
            )
            self._open_file(result)
        except RuntimeError as e:
            messagebox.showerror("Gagal", str(e))

    def _on_export_backup(self) -> None:
        """Handler tombol Export Backup: Simpan seluruh data DB ke file Excel."""
        from datetime import datetime as _dt

        default_name = f"backup_ritual_{_dt.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        filepath = filedialog.asksaveasfilename(
            title="Simpan Backup Database",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=default_name,
        )
        if not filepath:
            return

        try:
            count, result_path = export_to_excel(filepath)
            if count == 0:
                messagebox.showwarning(
                    "Database Kosong",
                    "Tidak ada data untuk di-export.\n"
                    "Tambahkan data terlebih dahulu.",
                )
                return
            messagebox.showinfo(
                "Export Berhasil",
                f"Berhasil export {count} record ke:\n"
                f"{os.path.basename(result_path)}\n\n"
                "File ini dapat di-import kembali kapan saja\n"
                'menggunakan tombol "Import Excel".',
            )
            self._open_file(result_path)
        except RuntimeError as e:
            messagebox.showerror("Gagal Export", str(e))

    @staticmethod
    def _open_file(filepath: str) -> None:
        """
        Membuka file menggunakan aplikasi default di OS.
        Windows: os.startfile | macOS: open | Linux: xdg-open
        """
        try:
            if sys.platform == "win32":
                os.startfile(filepath)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", filepath])
            else:
                subprocess.Popen(["xdg-open", filepath])
        except OSError as e:
            messagebox.showwarning("Gagal Membuka File", f"Tidak bisa membuka file:\n{e}")


# ============================================================
# Entry Point
# ============================================================
if __name__ == "__main__":
    app = RitualFormApp()
    app.mainloop()
