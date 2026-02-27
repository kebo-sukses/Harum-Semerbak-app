# -*- coding: utf-8 -*-
"""
database.py — Modul SQLite untuk menyimpan data formulir ritual.

Tabel utama: `ritual_forms`
Kolom:
  - id            : INTEGER PRIMARY KEY (auto-increment)
  - uuid          : TEXT UNIQUE (ID unik per-record, format UUID4)
  - nama_mendiang : TEXT (nama orang yang sudah meninggal)
  - nama_pengirim : TEXT (nama pengirim / 陽上)
  - tahun_lunar   : TEXT (tahun lunar, misal: 乙巳)
  - bulan_lunar   : TEXT (bulan lunar, misal: 正月)
  - hari_lunar    : TEXT (hari lunar, misal: 十五)
  - catatan       : TEXT (keterangan tambahan, opsional)
  - created_at    : TIMESTAMP (waktu pembuatan record)
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
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                uuid          TEXT    NOT NULL UNIQUE,
                nama_mendiang TEXT    NOT NULL,
                nama_pengirim TEXT    NOT NULL,
                tahun_lunar   TEXT    NOT NULL,
                bulan_lunar   TEXT    NOT NULL,
                hari_lunar    TEXT    NOT NULL,
                catatan       TEXT    DEFAULT '',
                created_at    TEXT    NOT NULL
            );
        """)
        conn.commit()
    except sqlite3.Error as e:
        raise RuntimeError(f"Gagal membuat tabel: {e}") from e
    finally:
        conn.close()


def insert_record(
    nama_mendiang: str,
    nama_pengirim: str,
    tahun_lunar: str,
    bulan_lunar: str,
    hari_lunar: str,
    catatan: str = "",
    db_path: str = _DB_PATH,
) -> str:
    """
    Menyimpan satu record baru ke tabel `ritual_forms`.

    Args:
        nama_mendiang: Nama orang yang sudah meninggal.
        nama_pengirim: Nama pengirim (陽上).
        tahun_lunar:   Tahun lunar (misal 乙巳).
        bulan_lunar:   Bulan lunar (misal 正月).
        hari_lunar:    Hari lunar (misal 十五).
        catatan:       Keterangan tambahan (opsional).
        db_path:       Path ke database (default: folder /database).

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
                (uuid, nama_mendiang, nama_pengirim,
                 tahun_lunar, bulan_lunar, hari_lunar,
                 catatan, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?);
            """,
            (
                record_uuid,
                nama_mendiang,
                nama_pengirim,
                tahun_lunar,
                bulan_lunar,
                hari_lunar,
                catatan,
                now,
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
# Jika dijalankan langsung, inisialisasi database sebagai test.
# ============================================================
if __name__ == "__main__":
    init_db()
    print(f"Database berhasil diinisialisasi di: {_DB_PATH}")

    # Test insert
    test_uuid = insert_record(
        nama_mendiang="蔡氏先人",
        nama_pengirim="蔡明志",
        tahun_lunar="乙巳",
        bulan_lunar="正月",
        hari_lunar="十五",
        catatan="Tes data awal",
    )
    print(f"Record test berhasil dibuat, UUID: {test_uuid}")

    # Test get all
    records = get_all_records()
    for r in records:
        print(r)
