# -*- coding: utf-8 -*-
"""
pdf_engine.py -- Engine PDF (ReportLab) untuk mencetak layer teks
di atas template label.pdf pada kertas US Letter (612x792pt).

Layout Output:
  - Dua panel identik (kiri & kanan) pada satu halaman.
  - Setiap panel berisi data mendiang dari database:
    * Kolom Panggilan + Nama Mandarin (vertikal)
    * Kolom Dari / pengirim (vertikal)
    * Keluarga (horizontal, bahasa Indonesia)
    * Prefix otomatis berdasarkan panggilan
    * Teks ritual tetap
    * Label tanggal + zodiak

Koordinat:
  Semua posisi dalam POINTS (pt), origin = kiri-bawah (ReportLab).
  Dipetakan dari analisis new.pdf reference output.

Fitur Calibration Offset:
  Parameter offset_x dan offset_y menggeser SELURUH teks secara global
  agar user bisa mengkompensasi ketidakpresisian printer.
"""

import os
import io
from datetime import datetime

from reportlab.lib.units import mm as _mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter


# ============================================================
# Konstanta Ukuran Kertas (US Letter, sesuai template label.pdf)
# ============================================================
PAGE_WIDTH_PT = 612.0    # 215.9mm
PAGE_HEIGHT_PT = 792.0   # 279.4mm

# Path template label.pdf
_TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "assets", "label.pdf"
)


# ============================================================
# Helper: Konversi mm ke points
# ============================================================
def mm(value: float) -> float:
    """Konversi milimeter ke points untuk ReportLab."""
    return value * _mm


# ============================================================
# Registrasi Font Mandarin
# ============================================================
_FONT_REGISTERED = False
FONT_NAME = "HanyiSentyPagoda"

# Font utama: HanyiSentyPagoda (sama dengan label new.pdf)
# Fallback: SimSun, Microsoft YaHei, KaiTi
_FONT_CANDIDATES = [
    # Font utama di assets/fonts project
    os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "..", "assets", "fonts", "HanyiSentyPagoda.ttf",
    ),
    # Fallback: font sistem Windows
    r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\msyh.ttc",
    os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "..", "assets", "fonts", "simsun.ttc",
    ),
    os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "..", "assets", "fonts", "kaiti.ttf",
    ),
]


def _register_font() -> None:
    """Mendaftarkan font Mandarin ke ReportLab (singleton)."""
    global _FONT_REGISTERED
    if _FONT_REGISTERED:
        return

    for font_path in _FONT_CANDIDATES:
        if os.path.isfile(font_path):
            try:
                pdfmetrics.registerFont(TTFont(FONT_NAME, font_path))
                _FONT_REGISTERED = True
                return
            except Exception:
                continue

    raise FileNotFoundError(
        "Font tidak ditemukan!\n"
        "Letakkan file HanyiSentyPagoda.ttf di folder assets/fonts/\n"
        f"Atau pastikan font tersedia di: {_FONT_CANDIDATES}"
    )


# ============================================================
# Posisi Teks pada Layout Dual-Panel (dalam POINTS)
# ============================================================

# Offset X antara panel kiri dan kanan
PANEL_OFFSET = 262.56    # ~92.6mm

# --- Kolom Dari (vertikal, atas-ke-bawah) ---
DARI_X = 68.28            # ~24.1mm dari kiri
DARI_Y_START = 366.48     # ~129.3mm dari bawah (karakter pertama)
DARI_SPACING = 28.68      # ~10.1mm per karakter

# --- Kolom Panggilan + Nama Mandarin (vertikal) ---
MENDIANG_X = 165.48       # ~58.4mm dari kiri
MENDIANG_Y_START = 338.04 # ~119.2mm dari bawah (karakter pertama)
MENDIANG_SPACING = 28.68  # ~10.1mm per karakter

# --- Prefix (di atas kolom mendiang) ---
PREFIX_X = 166.56          # ~58.8mm
PREFIX_Y_START = 394.32    # ~139.1mm
PREFIX_SPACING = 25.20     # ~8.9mm

# --- Teks ritual tetap (di bawah kolom mendiang) ---
RITUAL_X = 165.48          # sama dengan mendiang
RITUAL_Y_START = 176.64    # ~62.3mm
RITUAL_SPACING = 28.68     # sama dengan mendiang

# --- Keluarga / Indonesian label (horizontal, dekat bawah) ---
KELUARGA_X = 162.84        # ~57.4mm
KELUARGA_Y = 19.32         # ~6.8mm dari bawah

# --- Zodiak tahun ---
ZODIAC_X = 259.68          # ~91.6mm
ZODIAC_STEM_Y = 367.32     # posisi stem
ZODIAC_BRANCH_Y = 329.64   # posisi branch

# --- Label tanggal tetap ---
DATE_X = 259.68            # ~91.6mm
YEAR_LABEL_Y = 250.20      # nian
MONTH_LABEL_Y = 176.28     # yue
DAY_LABEL_Y = 108.48       # ri

# --- Tanda segel ---
FENG_X = 257.76            # ~90.9mm
FENG_Y = 44.28             # ~15.6mm
FU_X = 63.36               # ~22.3mm
FU_Y = 44.28               # ~15.6mm

# --- Ukuran font (setelah penyesuaian skala dari PDF asli) ---
FONT_MAIN = 16.8           # Panggilan, Mandarin, Dari
FONT_PREFIX = 14.9         # Prefix
FONT_KELUARGA = 9.3        # Label keluarga (Indonesia)
FONT_SEAL = 26.0           # Segel
FONT_DATE = 22.3           # Zodiak & label tanggal

# Teks ritual tetap (selalu sama)
_RITUAL_TEXT = "\u4e00\u4f4d\u6b63\u9b42\u6536\u9886"

# Karakter penanda gender mendiang untuk deteksi prefix
_FEMALE_INDICATORS = frozenset(
    "\u6bcd\u59d1\u5b38\u59bb\u5ab3\u59d0\u59b9\u5ac2\u59e8"
)

# Tabel zodiak China
_HEAVENLY_STEMS = "\u7532\u4e59\u4e19\u4e01\u620a\u5df1\u5e9a\u8f9b\u58ec\u7678"
_EARTHLY_BRANCHES = (
    "\u5b50\u4e11\u5bc5\u536f\u8fb0\u5df3\u5348\u672a\u7533\u9149\u620c\u4ea5"
)


# ============================================================
# Helper Functions
# ============================================================
def _get_ancestor_prefix(panggilan: str) -> str:
    """Tentukan prefix mendiang berdasarkan karakter panggilan.

    Returns:
        Prefix perempuan atau laki-laki.
    """
    for char in panggilan:
        if char in _FEMALE_INDICATORS:
            return "\u5148\u59a3"
    return "\u5148\u8003"


def _get_zodiac_year(year=None):
    """Hitung Heavenly Stem + Earthly Branch untuk tahun tertentu.

    Args:
        year: Tahun Masehi. Default = tahun sekarang.

    Returns:
        Tuple (stem, branch).
    """
    if year is None:
        year = datetime.now().year
    stem = _HEAVENLY_STEMS[(year - 4) % 10]
    branch = _EARTHLY_BRANCHES[(year - 4) % 12]
    return stem, branch


def _draw_vertical_chars(
    c,
    text,
    x_pt,
    y_start_pt,
    spacing_pt,
    offset_x_pt=0.0,
    offset_y_pt=0.0,
):
    """Gambar karakter vertikal (atas ke bawah), skip spasi.

    Args:
        c:            Canvas ReportLab.
        text:         String yang akan ditulis.
        x_pt:         Posisi X (points).
        y_start_pt:   Posisi Y karakter pertama (points).
        spacing_pt:   Jarak antar posisi karakter (points).
        offset_x_pt:  Calibration offset X (points).
        offset_y_pt:  Calibration offset Y (points).
    """
    idx = 0
    for char in text:
        if char != " ":
            final_x = x_pt + offset_x_pt
            final_y = y_start_pt - idx * spacing_pt + offset_y_pt
            c.drawString(final_x, final_y, char)
        idx += 1


def _draw_panel(c, data, x_offset, offset_x, offset_y):
    """Gambar semua teks untuk satu panel.

    Args:
        c:         Canvas ReportLab.
        data:      Dictionary berisi field formulir.
        x_offset:  Offset X panel (0 = kiri, PANEL_OFFSET = kanan).
        offset_x:  Calibration offset X (points).
        offset_y:  Calibration offset Y (points).
    """
    panggilan = data.get("panggilan", "")
    nama_mandarin = data.get("nama_mandarin", "")
    dari = data.get("dari", "")
    keterangan = data.get("keterangan", "")
    keluarga = data.get("keluarga", "")

    # --- 1. Kolom Dari + Keterangan (vertikal) ---
    dari_text = dari
    if keterangan:
        dari_text += keterangan
    c.setFont(FONT_NAME, FONT_MAIN)
    _draw_vertical_chars(
        c, dari_text,
        x_pt=DARI_X + x_offset,
        y_start_pt=DARI_Y_START,
        spacing_pt=DARI_SPACING,
        offset_x_pt=offset_x,
        offset_y_pt=offset_y,
    )

    # --- 2. Prefix ---
    prefix = _get_ancestor_prefix(panggilan)
    c.setFont(FONT_NAME, FONT_PREFIX)
    _draw_vertical_chars(
        c, prefix,
        x_pt=PREFIX_X + x_offset,
        y_start_pt=PREFIX_Y_START,
        spacing_pt=PREFIX_SPACING,
        offset_x_pt=offset_x,
        offset_y_pt=offset_y,
    )

    # --- 3. Kolom Panggilan + Nama Mandarin (vertikal) ---
    mendiang_text = panggilan + nama_mandarin
    c.setFont(FONT_NAME, FONT_MAIN)
    _draw_vertical_chars(
        c, mendiang_text,
        x_pt=MENDIANG_X + x_offset,
        y_start_pt=MENDIANG_Y_START,
        spacing_pt=MENDIANG_SPACING,
        offset_x_pt=offset_x,
        offset_y_pt=offset_y,
    )

    # --- 4. Teks ritual 一位正魂收领 ---
    # TIDAK digambar karena sudah ada di template label.pdf (warna merah)

    # --- 5. Keluarga (horizontal, bahasa Indonesia) ---
    if keluarga:
        short_keluarga = keluarga.split(" (")[0]
        short_keluarga = short_keluarga.replace(" Kandung", "")
        c.setFont(FONT_NAME, FONT_KELUARGA)
        c.drawString(
            KELUARGA_X + x_offset + offset_x,
            KELUARGA_Y + offset_y,
            short_keluarga,
        )

    # --- 6. Zodiak tahun ---
    stem, branch = _get_zodiac_year()
    c.setFont(FONT_NAME, FONT_DATE)
    c.drawString(
        ZODIAC_X + x_offset + offset_x,
        ZODIAC_STEM_Y + offset_y,
        stem,
    )
    c.drawString(
        ZODIAC_X + x_offset + offset_x,
        ZODIAC_BRANCH_Y + offset_y,
        branch,
    )

    # --- 7, 8: 年月日, 封附, 一位正魂收领 ---
    # TIDAK digambar karena sudah ada di template label.pdf (warna merah)


# ============================================================
# Fungsi Utama: Generate PDF
# ============================================================
def generate_pdf(
    data,
    output_path,
    offset_x=0.0,
    offset_y=0.0,
):
    """Menghasilkan PDF dual-panel dengan template label.pdf.

    Args:
        data: Dictionary berisi field formulir:
            - panggilan     : str
            - nama_mandarin : str
            - dari          : str
            - keterangan    : str (opsional)
            - keluarga      : str
        output_path: Path file PDF output.
        offset_x: Calibration offset horizontal (mm). Positif = geser kanan.
        offset_y: Calibration offset vertikal (mm). Positif = geser atas.

    Returns:
        str: Path absolut file PDF yang dihasilkan.
    """
    try:
        _register_font()

        # Konversi offset dari mm ke points
        ox_pt = mm(offset_x)
        oy_pt = mm(offset_y)

        overlay_path = output_path + ".overlay.tmp"
        c = canvas.Canvas(
            overlay_path,
            pagesize=(PAGE_WIDTH_PT, PAGE_HEIGHT_PT),
        )

        # Gambar panel KIRI
        _draw_panel(c, data, x_offset=0.0, offset_x=ox_pt, offset_y=oy_pt)

        # Gambar panel KANAN (identik, geser X)
        _draw_panel(
            c, data, x_offset=PANEL_OFFSET,
            offset_x=ox_pt, offset_y=oy_pt,
        )

        c.showPage()
        c.save()

        # Merge: template label.pdf (background) + overlay (teks)
        _merge_template_and_overlay(
            template_path=_TEMPLATE_PATH,
            overlay_path=overlay_path,
            output_path=output_path,
        )

        # Hapus file overlay sementara
        if os.path.isfile(overlay_path):
            os.remove(overlay_path)

        return os.path.abspath(output_path)

    except FileNotFoundError:
        raise
    except Exception as e:
        raise RuntimeError(f"Gagal membuat PDF: {e}") from e


def generate_pdf_bytes(
    data,
    offset_x=0.0,
    offset_y=0.0,
):
    """Menghasilkan PDF dual-panel sebagai bytes (in-memory).

    Sama seperti generate_pdf() tapi mengembalikan bytes tanpa
    menyimpan ke file. Cocok untuk preview dan cetak langsung.

    Args:
        data: Dictionary berisi field formulir.
        offset_x: Calibration offset horizontal (mm).
        offset_y: Calibration offset vertikal (mm).

    Returns:
        bytes: Data PDF dalam bytes.
    """
    try:
        _register_font()

        ox_pt = mm(offset_x)
        oy_pt = mm(offset_y)

        # Buat overlay teks ke buffer memori
        overlay_buf = io.BytesIO()
        c = canvas.Canvas(
            overlay_buf,
            pagesize=(PAGE_WIDTH_PT, PAGE_HEIGHT_PT),
        )

        _draw_panel(c, data, x_offset=0.0, offset_x=ox_pt, offset_y=oy_pt)
        _draw_panel(
            c, data, x_offset=PANEL_OFFSET,
            offset_x=ox_pt, offset_y=oy_pt,
        )

        c.showPage()
        c.save()

        # Merge template + overlay di memori
        overlay_buf.seek(0)
        overlay_reader = PdfReader(overlay_buf)
        overlay_page = overlay_reader.pages[0]

        if os.path.isfile(_TEMPLATE_PATH):
            template_reader = PdfReader(_TEMPLATE_PATH)
            base_page = template_reader.pages[0]
            base_page.merge_page(overlay_page)
            writer = PdfWriter()
            writer.add_page(base_page)
        else:
            writer = PdfWriter()
            writer.add_page(overlay_page)

        output_buf = io.BytesIO()
        writer.write(output_buf)
        return output_buf.getvalue()

    except FileNotFoundError:
        raise
    except Exception as e:
        raise RuntimeError(f"Gagal membuat PDF: {e}") from e


# ============================================================
# Helper: Merge template PDF + overlay teks
# ============================================================
def _merge_template_and_overlay(
    template_path,
    overlay_path,
    output_path,
):
    """Menggabungkan template label.pdf (background) dengan overlay teks."""
    overlay_reader = PdfReader(overlay_path)
    overlay_page = overlay_reader.pages[0]

    if os.path.isfile(template_path):
        template_reader = PdfReader(template_path)
        base_page = template_reader.pages[0]
        base_page.merge_page(overlay_page)
        writer = PdfWriter()
        writer.add_page(base_page)
    else:
        writer = PdfWriter()
        writer.add_page(overlay_page)

    with open(output_path, "wb") as f:
        writer.write(f)


# ============================================================
# Fungsi Kalibrasi: Cetak halaman test grid
# ============================================================
def generate_calibration_pdf(output_path):
    """Menghasilkan PDF kalibrasi berisi grid titik-titik setiap 10mm.

    Returns:
        str: Path absolut file PDF kalibrasi.
    """
    try:
        _register_font()

        c = canvas.Canvas(
            output_path,
            pagesize=(PAGE_WIDTH_PT, PAGE_HEIGHT_PT),
        )

        c.setFont(FONT_NAME, 6)
        c.setStrokeColorRGB(0, 0, 0)
        c.setFillColorRGB(0, 0, 0)

        max_x_mm = int(PAGE_WIDTH_PT / _mm) + 1
        max_y_mm = int(PAGE_HEIGHT_PT / _mm) + 1

        for x in range(0, max_x_mm, 10):
            for y in range(0, max_y_mm, 10):
                px = mm(x)
                py = mm(y)
                c.circle(px, py, 0.5, fill=1)

                if x % 50 == 0 and y % 50 == 0:
                    c.drawString(px + 2, py + 2, f"{x},{y}")

        c.setFont(FONT_NAME, 10)
        c.drawString(
            mm(10),
            PAGE_HEIGHT_PT - mm(5),
            "CALIBRATION GRID -- US Letter (216x279mm) -- Titik setiap 10mm",
        )

        c.showPage()
        c.save()

        return os.path.abspath(output_path)

    except Exception as e:
        raise RuntimeError(f"Gagal membuat PDF kalibrasi: {e}") from e


# ============================================================
# Test jika dijalankan langsung
# ============================================================
if __name__ == "__main__":
    test_data = {
        "panggilan": "\u7236\u89aa",
        "nama_mandarin": "\u9b4f\u4e9e\u660c",
        "dari": "\u5b5d\u4e94\u5b50 \u656c\u5949 \u53e9\u9996",
        "keluarga": "Ayah Kandung",
        "keterangan": "",
    }

    out = generate_pdf(test_data, "test_output.pdf", offset_x=0, offset_y=0)
    print(f"PDF berhasil dibuat: {out}")

    cal = generate_calibration_pdf("test_calibration.pdf")
    print(f"PDF kalibrasi berhasil dibuat: {cal}")
