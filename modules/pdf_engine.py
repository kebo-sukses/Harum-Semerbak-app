# -*- coding: utf-8 -*-
"""
pdf_engine.py — Engine PDF (ReportLab) untuk mencetak layer teks
di atas template label.pdf pada kertas A4 (210mm × 297mm).

Metode cetak:
  1. Merge template label.pdf sebagai background layer.
  2. Overlay teks hitam di posisi (X, Y) yang tepat.
  Saat di-print, teks jatuh di tempat yang benar di atas formulir.

Satuan:
  Semua koordinat dan ukuran menggunakan milimeter (mm).
  Konversi ke points (pt) dilakukan secara internal oleh helper `mm()`.

Fitur Calibration Offset:
  Parameter offset_x dan offset_y menggeser SELURUH teks secara global
  agar user bisa mengkompensasi ketidakpresisian printer.
"""

import os
from reportlab.lib.units import mm as _mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter


# ============================================================
# Konstanta Ukuran Kertas A4 (dalam mm)
# ============================================================
A4_WIDTH_MM = 210       # Lebar kertas A4 dalam milimeter
A4_HEIGHT_MM = 297      # Tinggi kertas A4 dalam milimeter

# Path template label.pdf
_TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "assets", "label.pdf"
)


# ============================================================
# Helper: Konversi mm ke points (ReportLab menggunakan points)
# 1 mm = 2.834645669 points
# ============================================================
def mm(value: float) -> float:
    """Konversi milimeter ke points untuk ReportLab."""
    return value * _mm


# ============================================================
# Registrasi Font Mandarin
# ============================================================
_FONT_REGISTERED = False
FONT_NAME = "SimSun"

# Daftar path font yang mungkin tersimpan di sistem
_FONT_CANDIDATES = [
    # Path Windows standar
    r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\msyh.ttc",
    # Path relatif di folder assets/fonts project
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "fonts", "simsun.ttc"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "fonts", "kaiti.ttf"),
]


def _register_font() -> None:
    """
    Mendaftarkan font Mandarin ke ReportLab.
    Mencari dari daftar kandidat path font.
    Hanya dieksekusi sekali (singleton pattern).

    Raises:
        FileNotFoundError: Jika tidak ada font Mandarin yang ditemukan.
    """
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
        "Font Mandarin tidak ditemukan!\n"
        "Letakkan file 'simsun.ttc' atau 'kaiti.ttf' di folder assets/fonts/\n"
        f"Atau pastikan font tersedia di: {_FONT_CANDIDATES}"
    )


# ============================================================
# Peta Koordinat Teks pada Formulir Ritual (dalam mm)
#
# Koordinat diukur dari KIRI BAWAH kertas (origin ReportLab).
#   X = jarak dari tepi kiri kertas
#   Y = jarak dari tepi bawah kertas
#
# Layout formulir A4 (210×297mm, dari label.pdf):
# ┌─────────────────────────────────────────┐  ← 297mm (atas)
# │  [Header formulir dari label.pdf]        │
# │  [陽上]       [太歲]                     │
# │  [Panggilan + Mandarin]  [乙巳]          │
# │              [年/月/日]                  │
# └─────────────────────────────────────────┘  ← 0mm (bawah)
#
# Teks pada formulir ditulis VERTIKAL (atas ke bawah).
# Kita cetak per-karakter secara vertikal dengan jarak tetap.
# ============================================================


def _draw_vertical_text(
    c: canvas.Canvas,
    text: str,
    x_mm: float,
    y_start_mm: float,
    font_size: float,
    line_spacing_mm: float,
    offset_x: float,
    offset_y: float,
) -> None:
    """
    Menggambar teks secara vertikal (atas ke bawah), satu karakter per baris.

    Args:
        c:               Canvas ReportLab.
        text:            String yang akan ditulis vertikal.
        x_mm:            Posisi X (mm) dari kiri kertas.
        y_start_mm:      Posisi Y (mm) karakter pertama dari bawah kertas.
        font_size:       Ukuran font dalam points.
        line_spacing_mm: Jarak antar karakter secara vertikal (mm).
        offset_x:        Calibration offset X global (mm).
        offset_y:        Calibration offset Y global (mm).
    """
    for i, char in enumerate(text):
        # X final = posisi dasar + offset kalibrasi
        final_x = mm(x_mm + offset_x)
        # Y final = posisi awal - (indeks × jarak) + offset kalibrasi
        # Dikurangi karena teks turun ke bawah (Y berkurang)
        final_y = mm(y_start_mm + offset_y) - mm(i * line_spacing_mm)
        c.drawString(final_x, final_y, char)


def _draw_horizontal_text(
    c: canvas.Canvas,
    text: str,
    x_mm: float,
    y_mm: float,
    font_size: float,
    offset_x: float,
    offset_y: float,
) -> None:
    """
    Menggambar teks secara horizontal (kiri ke kanan) — untuk tanggal, dll.

    Args:
        c:        Canvas ReportLab.
        text:     String yang akan ditulis horizontal.
        x_mm:     Posisi X (mm) dari kiri kertas.
        y_mm:     Posisi Y (mm) dari bawah kertas.
        font_size: Ukuran font dalam points.
        offset_x: Calibration offset X global (mm).
        offset_y: Calibration offset Y global (mm).
    """
    final_x = mm(x_mm + offset_x)      # X + offset kalibrasi
    final_y = mm(y_mm + offset_y)       # Y + offset kalibrasi
    c.drawString(final_x, final_y, text)


# ============================================================
# Fungsi Utama: Generate PDF
# ============================================================
def generate_pdf(
    data: dict,
    output_path: str,
    offset_x: float = 0.0,
    offset_y: float = 0.0,
) -> str:
    """
    Menghasilkan PDF dengan template label.pdf sebagai background
    dan overlay teks, pada kertas A4 (210×297mm).

    Args:
        data: Dictionary berisi field formulir (schema Excel Sheet 2):
            - panggilan     : str — Sebutan Mandarin (母親許門, 父親, 祖父)
            - nama_mandarin : str — Nama Mandarin mendiang (梁氏橋玉)
            - dari          : str — Pengirim / hubungan (孝男, 外孫女敬奉)
            - keterangan    : str — Info tambahan (合家敬奉) [opsional]
            - tahun_lunar   : str — Tahun lunar / 太歲 (乙巳) [opsional]
            - bulan_lunar   : str — Bulan lunar (正月) [opsional]
            - hari_lunar    : str — Hari lunar (十五) [opsional]
        output_path: Path file PDF output.
        offset_x: Calibration offset horizontal (mm). Positif = geser ke kanan.
        offset_y: Calibration offset vertikal (mm). Positif = geser ke atas.

    Returns:
        str: Path absolut file PDF yang dihasilkan.

    Raises:
        FileNotFoundError: Jika font Mandarin tidak ditemukan.
        Exception: Jika terjadi error saat membuat PDF.
    """
    try:
        # Pastikan font terdaftar
        _register_font()

        # Buat canvas overlay teks dengan ukuran kertas A4
        page_w = mm(A4_WIDTH_MM)    # Lebar halaman (points)
        page_h = mm(A4_HEIGHT_MM)   # Tinggi halaman (points)

        # File sementara untuk overlay teks
        overlay_path = output_path + ".overlay.tmp"
        c = canvas.Canvas(overlay_path, pagesize=(page_w, page_h))

        # --------------------------------------------------------
        # Ambil data dari dictionary (schema Excel Sheet 2)
        # --------------------------------------------------------
        panggilan = data.get("panggilan", "")         # 母親許門 / 父親 / 祖父
        nama_mandarin = data.get("nama_mandarin", "") # 梁氏橋玉
        dari = data.get("dari", "")                   # 孝男 / 外孫女敬奉
        keterangan = data.get("keterangan", "")       # 合家敬奉
        tahun_lunar = data.get("tahun_lunar", "")
        bulan_lunar = data.get("bulan_lunar", "")
        hari_lunar = data.get("hari_lunar", "")

        # --------------------------------------------------------
        # KOLOM KIRI: Panggilan + Nama Mandarin (ditulis vertikal)
        # Gabungan: "母親許門 梁氏橋玉" — sebutan + nama mendiang
        # Posisi: X=68mm dari kiri, mulai Y=200mm dari bawah
        # Teks ditulis atas-ke-bawah satu karakter per baris
        # Jarak antar karakter: 12mm
        # Font size: 14pt
        # --------------------------------------------------------
        mendiang_text = panggilan + nama_mandarin
        c.setFont(FONT_NAME, 14)
        _draw_vertical_text(
            c,
            text=mendiang_text,
            x_mm=68,                # X = 68mm dari kiri kertas
            y_start_mm=200,         # Y = 200mm dari bawah kertas (titik awal)
            font_size=14,
            line_spacing_mm=12,     # Jarak vertikal antar karakter = 12mm
            offset_x=offset_x,
            offset_y=offset_y,
        )

        # --------------------------------------------------------
        # KOLOM TENGAH-KIRI: Dari / Pengirim + Keterangan (vertikal)
        # Gabungan: "孝男" atau "外孫女敬奉" + opsional "合家敬奉"
        # Posisi: X=42mm dari kiri, mulai Y=170mm dari bawah
        # Jarak antar karakter: 12mm
        # Font size: 14pt
        # --------------------------------------------------------
        pengirim_text = dari
        if keterangan:
            pengirim_text += keterangan
        c.setFont(FONT_NAME, 14)
        _draw_vertical_text(
            c,
            text=pengirim_text,
            x_mm=42,                # X = 42mm dari kiri kertas
            y_start_mm=170,         # Y = 170mm dari bawah kertas
            font_size=14,
            line_spacing_mm=12,     # Jarak vertikal antar karakter = 12mm
            offset_x=offset_x,
            offset_y=offset_y,
        )

        # --------------------------------------------------------
        # KOLOM KANAN: Tahun Lunar / 太歲 (ditulis vertikal)
        # Posisi: X=155mm dari kiri, mulai Y=200mm dari bawah
        # Misal: "乙巳" → dua karakter vertikal
        # Jarak antar karakter: 14mm
        # Font size: 16pt (sedikit lebih besar untuk tahun)
        # --------------------------------------------------------
        c.setFont(FONT_NAME, 16)
        _draw_vertical_text(
            c,
            text=tahun_lunar,
            x_mm=155,               # X = 155mm dari kiri kertas
            y_start_mm=200,         # Y = 200mm dari bawah kertas
            font_size=16,
            line_spacing_mm=14,     # Jarak vertikal antar karakter = 14mm
            offset_x=offset_x,
            offset_y=offset_y,
        )

        # --------------------------------------------------------
        # BARIS BAWAH: Bulan Lunar (horizontal)
        # Posisi: X=140mm, Y=80mm dari bawah kertas
        # Font size: 14pt
        # --------------------------------------------------------
        c.setFont(FONT_NAME, 14)
        _draw_horizontal_text(
            c,
            text=bulan_lunar,
            x_mm=140,               # X = 140mm dari kiri kertas
            y_mm=80,                # Y = 80mm dari bawah kertas
            font_size=14,
            offset_x=offset_x,
            offset_y=offset_y,
        )

        # --------------------------------------------------------
        # BARIS BAWAH: Hari Lunar (horizontal)
        # Posisi: X=170mm, Y=55mm dari bawah kertas
        # Font size: 14pt
        # --------------------------------------------------------
        c.setFont(FONT_NAME, 14)
        _draw_horizontal_text(
            c,
            text=hari_lunar,
            x_mm=170,               # X = 170mm dari kiri kertas
            y_mm=55,                # Y = 55mm dari bawah kertas
            font_size=14,
            offset_x=offset_x,
            offset_y=offset_y,
        )

        # --------------------------------------------------------
        # Simpan overlay teks
        # --------------------------------------------------------
        c.showPage()
        c.save()

        # --------------------------------------------------------
        # Merge: template label.pdf (background) + overlay (teks)
        # --------------------------------------------------------
        _merge_template_and_overlay(
            template_path=_TEMPLATE_PATH,
            overlay_path=overlay_path,
            output_path=output_path,
            page_w=page_w,
            page_h=page_h,
        )

        # Hapus file overlay sementara
        if os.path.isfile(overlay_path):
            os.remove(overlay_path)

        return os.path.abspath(output_path)

    except FileNotFoundError:
        raise
    except Exception as e:
        raise RuntimeError(f"Gagal membuat PDF: {e}") from e


# ============================================================
# Helper: Merge template PDF + overlay teks
# ============================================================
def _merge_template_and_overlay(
    template_path: str,
    overlay_path: str,
    output_path: str,
    page_w: float,
    page_h: float,
) -> None:
    """
    Menggabungkan template label.pdf (background) dengan overlay teks.
    Jika template tidak ditemukan, output hanya berisi overlay teks.

    Args:
        template_path: Path ke file template PDF.
        overlay_path:  Path ke overlay teks PDF.
        output_path:   Path file output gabungan.
        page_w:        Lebar halaman (points).
        page_h:        Tinggi halaman (points).
    """
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
# Membantu user menemukan offset_x / offset_y yang tepat.
# ============================================================
def generate_calibration_pdf(output_path: str) -> str:
    """
    Menghasilkan PDF kalibrasi berisi grid titik-titik
    setiap 10mm untuk membantu menentukan offset printer.

    Args:
        output_path: Path file PDF output kalibrasi.

    Returns:
        str: Path absolut file PDF kalibrasi.
    """
    try:
        _register_font()

        page_w = mm(A4_WIDTH_MM)
        page_h = mm(A4_HEIGHT_MM)
        c = canvas.Canvas(output_path, pagesize=(page_w, page_h))

        c.setFont(FONT_NAME, 6)
        c.setStrokeColorRGB(0, 0, 0)       # Warna hitam
        c.setFillColorRGB(0, 0, 0)

        # Gambar grid setiap 10mm
        for x in range(0, A4_WIDTH_MM + 1, 10):         # 0, 10, 20, ... 210mm
            for y in range(0, A4_HEIGHT_MM + 1, 10):     # 0, 10, 20, ... 297mm
                px = mm(x)      # Konversi X mm ke points
                py = mm(y)      # Konversi Y mm ke points

                # Gambar titik kecil (lingkaran radius 0.5pt)
                c.circle(px, py, 0.5, fill=1)

                # Label koordinat setiap 50mm agar mudah dibaca
                if x % 50 == 0 and y % 50 == 0:
                    c.drawString(px + 2, py + 2, f"{x},{y}")

        # Judul halaman kalibrasi
        c.setFont(FONT_NAME, 10)
        c.drawString(mm(10), mm(A4_HEIGHT_MM - 5), "CALIBRATION GRID — A4 (210×297mm) — Titik setiap 10mm")

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
        "panggilan": "母親許門",
        "nama_mandarin": "梁氏橋玉",
        "dari": "孝男",
        "keterangan": "合家敬奉",
        "tahun_lunar": "乙巳",
        "bulan_lunar": "正月",
        "hari_lunar": "十五",
    }

    out = generate_pdf(test_data, "test_output.pdf", offset_x=0, offset_y=0)
    print(f"PDF berhasil dibuat: {out}")

    cal = generate_calibration_pdf("test_calibration.pdf")
    print(f"PDF kalibrasi berhasil dibuat: {cal}")
