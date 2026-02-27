# -*- coding: utf-8 -*-
"""
database.py — Modul SQLite untuk menyimpan data formulir ritual.

Schema disesuaikan dengan file Excel 'segel data4.xlsx' Sheet 2:
  Kolom A: NO           → id (auto)
  Kolom B: NAMA         → nama (nama panggilan Indonesia)
  Kolom C: PANGGILAN    → panggilan (sebutan Mandarin: 母親許門, 父親, 祖父)
  Kolom D: ATAS NAMA    → nama_mandarin (nama Mandarin mendiang: 梁氏橋玉)
  Kolom E: PENYEBUTAN   → penyebutan (romanisasi/Hokkien: Nio Kiaw Gek)
  Kolom F: DARI         → dari (pengirim/hubungan: 孝男, 外孫女敬奉)
  Kolom G: KELUARGA     → keluarga (relasi Indonesia: Ibu Kandung, dll)
  Kolom H: KETERANGAN   → keterangan (info tambahan: 合家敬奉, dll)

Tambahan field cetak:
  - tahun_lunar : TEXT (tahun lunar, misal: 乙巳)
  - bulan_lunar : TEXT (bulan lunar, misal: 正月)
  - hari_lunar  : TEXT (hari lunar, misal: 十五)
"""

import sqlite3
import uuid
import os
from datetime import datetime


# ============================================================
# Path default database — disimpan di folder /database
# ============================================================
_DB_DIR = os.path.dirname(os.path.abspath(__file__))
_DB_PATH = os.path.join(_DB_DIR, "ritual_forms.db")


def _get_connection(db_path: str = _DB_PATH) -> sqlite3.Connection:
    """
    Membuka koneksi ke database SQLite.
    Mengaktifkan WAL mode untuk performa baca/tulis yang lebih baik.

    Args:
        db_path: Path absolut ke file .db

    Returns:
        sqlite3.Connection
    """
    try:
        conn = sqlite3.connect(db_path)
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.row_factory = sqlite3.Row  # Agar hasil query bisa diakses via nama kolom
        return conn
    except sqlite3.Error as e:
        raise RuntimeError(f"Gagal membuka database: {e}") from e


def init_db(db_path: str = _DB_PATH) -> None:
    """
    Membuat tabel `ritual_forms` jika belum ada.
    Dipanggil sekali saat aplikasi pertama kali dijalankan.

    Args:
        db_path: Path absolut ke file .db
    """
    try:
        conn = _get_connection(db_path)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS ritual_forms (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                uuid           TEXT    NOT NULL UNIQUE,
                nama           TEXT    DEFAULT '',
                panggilan      TEXT    NOT NULL,
                nama_mandarin  TEXT    NOT NULL,
                penyebutan     TEXT    DEFAULT '',
                dari           TEXT    NOT NULL,
                keluarga       TEXT    DEFAULT '',
                keterangan     TEXT    DEFAULT '',
                tahun_lunar    TEXT    DEFAULT '',
                bulan_lunar    TEXT    DEFAULT '',
                hari_lunar     TEXT    DEFAULT '',
                created_at     TEXT    NOT NULL
            );
        """)
        conn.commit()
    except sqlite3.Error as e:
        raise RuntimeError(f"Gagal membuat tabel: {e}") from e
    finally:
        conn.close()


def insert_record(
    panggilan: str,
    nama_mandarin: str,
    dari: str,
    nama: str = "",
    penyebutan: str = "",
    keluarga: str = "",
    keterangan: str = "",
    tahun_lunar: str = "",
    bulan_lunar: str = "",
    hari_lunar: str = "",
    db_path: str = _DB_PATH,
) -> str:
    """
    Menyimpan satu record baru ke tabel `ritual_forms`.

    Args:
        panggilan:     Sebutan Mandarin (母親許門, 父親, 祖父).
        nama_mandarin: Nama Mandarin mendiang (梁氏橋玉).
        dari:          Pengirim / hubungan (孝男, 外孫女敬奉).
        nama:          Nama panggilan Indonesia (opsional).
        penyebutan:    Romanisasi / Hokkien (opsional).
        keluarga:      Relasi keluarga Indonesia (opsional).
        keterangan:    Info tambahan (opsional).
        tahun_lunar:   Tahun lunar (misal 乙巳).
        bulan_lunar:   Bulan lunar (misal 正月).
        hari_lunar:    Hari lunar (misal 十五).
        db_path:       Path ke database.

    Returns:
        str: UUID dari record yang baru dibuat.
    """
    record_uuid = str(uuid.uuid4())
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        conn = _get_connection(db_path)
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO ritual_forms
                (uuid, nama, panggilan, nama_mandarin,
                 penyebutan, dari, keluarga, keterangan,
                 tahun_lunar, bulan_lunar, hari_lunar, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """,
            (
                record_uuid, nama, panggilan, nama_mandarin,
                penyebutan, dari, keluarga, keterangan,
                tahun_lunar, bulan_lunar, hari_lunar, now,
            ),
        )
        conn.commit()
        return record_uuid
    except sqlite3.Error as e:
        raise RuntimeError(f"Gagal menyimpan data: {e}") from e
    finally:
        conn.close()


def get_all_records(db_path: str = _DB_PATH) -> list[dict]:
    """
    Mengambil seluruh record dari tabel `ritual_forms`,
    diurutkan berdasarkan waktu pembuatan (terbaru di atas).

    Args:
        db_path: Path ke database.

    Returns:
        list[dict]: Daftar record dalam bentuk dictionary.
    """
    try:
        conn = _get_connection(db_path)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT * FROM ritual_forms ORDER BY created_at DESC;"
        )
        rows = cursor.fetchall()
        return [dict(row) for row in rows]
    except sqlite3.Error as e:
        raise RuntimeError(f"Gagal mengambil data: {e}") from e
    finally:
        conn.close()


def get_record_by_uuid(record_uuid: str, db_path: str = _DB_PATH) -> dict | None:
    """
    Mengambil satu record berdasarkan UUID.

    Args:
        record_uuid: UUID record yang dicari.
        db_path:     Path ke database.

    Returns:
        dict | None: Dictionary record, atau None jika tidak ditemukan.
    """
    try:
        conn = _get_connection(db_path)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT * FROM ritual_forms WHERE uuid = ?;",
            (record_uuid,),
        )
        row = cursor.fetchone()
        return dict(row) if row else None
    except sqlite3.Error as e:
        raise RuntimeError(f"Gagal mengambil data: {e}") from e
    finally:
        conn.close()


def delete_record(record_uuid: str, db_path: str = _DB_PATH) -> bool:
    """
    Menghapus satu record berdasarkan UUID.

    Args:
        record_uuid: UUID record yang akan dihapus.
        db_path:     Path ke database.

    Returns:
        bool: True jika berhasil dihapus, False jika tidak ditemukan.
    """
    try:
        conn = _get_connection(db_path)
        cursor = conn.cursor()
        cursor.execute(
            "DELETE FROM ritual_forms WHERE uuid = ?;",
            (record_uuid,),
        )
        conn.commit()
        return cursor.rowcount > 0
    except sqlite3.Error as e:
        raise RuntimeError(f"Gagal menghapus data: {e}") from e
    finally:
        conn.close()


# ============================================================
# Import Excel → Database
# ============================================================
def import_from_excel(
    excel_path: str,
    tahun_lunar: str = "",
    bulan_lunar: str = "",
    hari_lunar: str = "",
    db_path: str = _DB_PATH,
) -> int:
    """
    Import data dari file Excel Sheet 2 ke database.
    Kolom: A=NO, B=NAMA, C=PANGGILAN, D=ATAS NAMA MANDARIN,
           E=PENYEBUTAN, F=DARI, G=KELUARGA, H=KETERANGAN.

    Args:
        excel_path:  Path ke file .xlsx
        tahun_lunar: Tahun lunar default untuk semua record.
        bulan_lunar: Bulan lunar default.
        hari_lunar:  Hari lunar default.
        db_path:     Path ke database.

    Returns:
        int: Jumlah record yang berhasil diimport.
    """
    try:
        import openpyxl
    except ImportError as e:
        raise RuntimeError(
            "Library openpyxl belum terinstal. "
            "Jalankan: pip install openpyxl"
        ) from e

    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True)
        ws = wb.worksheets[1]  # Sheet 2 (index 1)

        count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            # row: A=NO, B=NAMA, C=PANGGILAN,
            #      D=NAMA MANDARIN, E=PENYEBUTAN,
            #      F=DARI, G=KELUARGA, H=KETERANGAN
            panggilan = str(row[2] or "").strip() if len(row) > 2 else ""
            nama_mandarin = str(row[3] or "").strip() if len(row) > 3 else ""
            dari = str(row[5] or "").strip() if len(row) > 5 else ""

            # Skip baris kosong (tanpa panggilan DAN nama mandarin)
            if not panggilan and not nama_mandarin:
                continue

            insert_record(
                panggilan=panggilan,
                nama_mandarin=nama_mandarin,
                dari=dari,
                nama=str(row[1] or "").strip() if len(row) > 1 else "",
                penyebutan=str(row[4] or "").strip() if len(row) > 4 else "",
                keluarga=str(row[6] or "").strip() if len(row) > 6 else "",
                keterangan=str(row[7] or "").strip() if len(row) > 7 else "",
                tahun_lunar=tahun_lunar,
                bulan_lunar=bulan_lunar,
                hari_lunar=hari_lunar,
                db_path=db_path,
            )
            count += 1

        wb.close()
        return count

    except (OSError, KeyError) as e:
        raise RuntimeError(f"Gagal membaca file Excel: {e}") from e


# ============================================================
# Jika dijalankan langsung, inisialisasi database sebagai test.
# ============================================================
if __name__ == "__main__":
    init_db()
    print(f"Database berhasil diinisialisasi di: {_DB_PATH}")

    # Test import dari Excel
    xlsx = r"e:\harum semerbak produk\segel data4.xlsx"
    if os.path.isfile(xlsx):
        count = import_from_excel(
            xlsx,
            tahun_lunar="乙巳",
            bulan_lunar="正月",
            hari_lunar="十五",
        )
        print(f"Berhasil import {count} record dari Excel.")

    # Test get all
    records = get_all_records()
    for r in records[:5]:
        print(r)
