# -*- coding: utf-8 -*-
"""
main.py — Aplikasi Desktop "Fill-in-the-Blanks" Formulir Ritual
Framework: CustomTkinter (modern Tkinter)

Fitur:
  1. Form Input  : Entry fields sesuai schema Excel Sheet 2
                    (Nama, Panggilan, Nama Mandarin, Penyebutan, Dari, Keluarga,
                     Keterangan, Tahun/Bulan/Hari Lunar).
  2. Tabel Preview: Treeview menampilkan data tersimpan di SQLite.
  3. Tombol Cetak : Generate PDF layer transparan & buka otomatis.
  4. Import Excel : Import data dari file Excel (Sheet 2).
  5. Calibration  : Offset X/Y untuk koreksi posisi cetak printer.
"""

import os
import sys
import subprocess
from tkinter import ttk, messagebox, filedialog

import customtkinter as ctk

# ============================================================
# Import modul internal
# ============================================================
# Pastikan root project ada di sys.path
_ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if _ROOT_DIR not in sys.path:
    sys.path.insert(0, _ROOT_DIR)

from database.database import (
    init_db, insert_record, get_all_records, delete_record, import_from_excel,
)
from modules.pdf_engine import generate_pdf, generate_calibration_pdf


# ============================================================
# Konstanta UI
# ============================================================
APP_TITLE = "Formulir Ritual — Fill-in-the-Blanks (F4)"
APP_WIDTH = 1100
APP_HEIGHT = 720

# Folder output PDF
_OUTPUT_DIR = os.path.join(_ROOT_DIR, "output")
os.makedirs(_OUTPUT_DIR, exist_ok=True)


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

        # --- Build UI ---
        self._build_input_frame()
        self._build_calibration_frame()
        self._build_button_frame()
        self._build_table_frame()

        # --- Load data awal ke tabel ---
        self._refresh_table()

    # ========================================================
    # Frame Input
    # ========================================================
    def _build_input_frame(self) -> None:
        """Membangun frame form input di bagian atas (schema Excel Sheet 2)."""
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(padx=15, pady=(15, 5), fill="x")

        ctk.CTkLabel(
            frame, text="📝 Data Formulir Ritual", font=ctk.CTkFont(size=16, weight="bold")
        ).grid(row=0, column=0, columnspan=6, padx=10, pady=(10, 5), sticky="w")

        # --- Baris 1: Nama & Panggilan & Nama Mandarin ---
        ctk.CTkLabel(frame, text="Nama:").grid(
            row=1, column=0, padx=(10, 5), pady=4, sticky="e"
        )
        self.entry_nama = ctk.CTkEntry(frame, width=180, placeholder_text="nama Indonesia")
        self.entry_nama.grid(row=1, column=1, padx=5, pady=4, sticky="w")

        ctk.CTkLabel(frame, text="Panggilan (稱呼):").grid(
            row=1, column=2, padx=(15, 5), pady=4, sticky="e"
        )
        self.entry_panggilan = ctk.CTkEntry(frame, width=180, placeholder_text="母親許門")
        self.entry_panggilan.grid(row=1, column=3, padx=5, pady=4, sticky="w")

        ctk.CTkLabel(frame, text="Nama Mandarin:").grid(
            row=1, column=4, padx=(15, 5), pady=4, sticky="e"
        )
        self.entry_mandarin = ctk.CTkEntry(frame, width=180, placeholder_text="梁氏橋玉")
        self.entry_mandarin.grid(row=1, column=5, padx=(5, 10), pady=4, sticky="w")

        # --- Baris 2: Penyebutan & Dari & Keluarga ---
        ctk.CTkLabel(frame, text="Penyebutan:").grid(
            row=2, column=0, padx=(10, 5), pady=4, sticky="e"
        )
        self.entry_penyebutan = ctk.CTkEntry(frame, width=180, placeholder_text="Nio Kiaw Gek")
        self.entry_penyebutan.grid(row=2, column=1, padx=5, pady=4, sticky="w")

        ctk.CTkLabel(frame, text="Dari (陽上):").grid(
            row=2, column=2, padx=(15, 5), pady=4, sticky="e"
        )
        self.entry_dari = ctk.CTkEntry(frame, width=180, placeholder_text="孝男")
        self.entry_dari.grid(row=2, column=3, padx=5, pady=4, sticky="w")

        ctk.CTkLabel(frame, text="Keluarga:").grid(
            row=2, column=4, padx=(15, 5), pady=4, sticky="e"
        )
        self.entry_keluarga = ctk.CTkEntry(frame, width=180, placeholder_text="Ibu Kandung")
        self.entry_keluarga.grid(row=2, column=5, padx=(5, 10), pady=4, sticky="w")

        # --- Baris 3: Keterangan & Tahun/Bulan/Hari Lunar ---
        ctk.CTkLabel(frame, text="Keterangan:").grid(
            row=3, column=0, padx=(10, 5), pady=4, sticky="e"
        )
        self.entry_keterangan = ctk.CTkEntry(frame, width=180, placeholder_text="合家敬奉")
        self.entry_keterangan.grid(row=3, column=1, padx=5, pady=4, sticky="w")

        ctk.CTkLabel(frame, text="Tahun Lunar:").grid(
            row=3, column=2, padx=(15, 5), pady=4, sticky="e"
        )
        self.entry_tahun = ctk.CTkEntry(frame, width=100, placeholder_text="乙巳")
        self.entry_tahun.grid(row=3, column=3, padx=5, pady=4, sticky="w")

        ctk.CTkLabel(frame, text="Bulan Lunar:").grid(
            row=3, column=4, padx=(15, 5), pady=4, sticky="e"
        )
        self.entry_bulan = ctk.CTkEntry(frame, width=100, placeholder_text="正月")
        self.entry_bulan.grid(row=3, column=5, padx=(5, 10), pady=4, sticky="w")

        ctk.CTkLabel(frame, text="Hari Lunar:").grid(
            row=4, column=0, padx=(10, 5), pady=(4, 10), sticky="e"
        )
        self.entry_hari = ctk.CTkEntry(frame, width=100, placeholder_text="十五")
        self.entry_hari.grid(row=4, column=1, padx=5, pady=(4, 10), sticky="w")

    # ========================================================
    # Frame Kalibrasi
    # ========================================================
    def _build_calibration_frame(self) -> None:
        """Membangun frame offset kalibrasi printer."""
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(padx=15, pady=5, fill="x")

        ctk.CTkLabel(
            frame, text="🔧 Calibration Offset (mm)", font=ctk.CTkFont(size=13, weight="bold")
        ).grid(row=0, column=0, columnspan=4, padx=10, pady=(8, 3), sticky="w")

        ctk.CTkLabel(frame, text="Offset X:").grid(
            row=1, column=0, padx=(10, 5), pady=5, sticky="e"
        )
        self.entry_offset_x = ctk.CTkEntry(frame, width=80, placeholder_text="0.0")
        self.entry_offset_x.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.entry_offset_x.insert(0, "0.0")

        ctk.CTkLabel(frame, text="Offset Y:").grid(
            row=1, column=2, padx=(20, 5), pady=5, sticky="e"
        )
        self.entry_offset_y = ctk.CTkEntry(frame, width=80, placeholder_text="0.0")
        self.entry_offset_y.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        self.entry_offset_y.insert(0, "0.0")

        # Tombol cetak kalibrasi
        ctk.CTkButton(
            frame,
            text="Cetak Grid Kalibrasi",
            width=180,
            command=self._on_calibration_print,
        ).grid(row=1, column=4, padx=15, pady=5, sticky="w")

    # ========================================================
    # Frame Tombol Aksi
    # ========================================================
    def _build_button_frame(self) -> None:
        """Membangun baris tombol aksi: Simpan, Cetak, Hapus, Import."""
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(padx=15, pady=5, fill="x")

        ctk.CTkButton(
            frame, text="💾 Simpan Data", width=140, command=self._on_save
        ).pack(side="left", padx=8, pady=10)

        ctk.CTkButton(
            frame, text="🖨️ Cetak PDF", width=140, fg_color="green", command=self._on_print
        ).pack(side="left", padx=8, pady=10)

        ctk.CTkButton(
            frame, text="🗑️ Hapus Terpilih", width=140, fg_color="red", command=self._on_delete
        ).pack(side="left", padx=8, pady=10)

        ctk.CTkButton(
            frame, text="📂 Import Excel", width=140, fg_color="#E67E22", command=self._on_import_excel
        ).pack(side="left", padx=8, pady=10)

        ctk.CTkButton(
            frame, text="🔄 Refresh", width=120, command=self._refresh_table
        ).pack(side="right", padx=8, pady=10)

    # ========================================================
    # Tabel Preview (Treeview)
    # ========================================================
    def _build_table_frame(self) -> None:
        """Membangun tabel preview data menggunakan ttk.Treeview."""
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(padx=15, pady=(5, 15), fill="both", expand=True)

        ctk.CTkLabel(
            frame, text="📋 Data Tersimpan", font=ctk.CTkFont(size=14, weight="bold")
        ).pack(padx=10, pady=(10, 5), anchor="w")

        # Kolom tabel
        columns = (
            "uuid", "nama", "panggilan", "mandarin",
            "penyebutan", "dari", "keluarga", "keterangan",
            "tahun", "bulan", "hari", "tanggal",
        )
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=10)

        # Header kolom
        self.tree.heading("uuid", text="ID")
        self.tree.heading("nama", text="Nama")
        self.tree.heading("panggilan", text="Panggilan")
        self.tree.heading("mandarin", text="Mandarin")
        self.tree.heading("penyebutan", text="Penyebutan")
        self.tree.heading("dari", text="Dari")
        self.tree.heading("keluarga", text="Keluarga")
        self.tree.heading("keterangan", text="Keterangan")
        self.tree.heading("tahun", text="Tahun")
        self.tree.heading("bulan", text="Bulan")
        self.tree.heading("hari", text="Hari")
        self.tree.heading("tanggal", text="Dibuat")

        # Lebar kolom
        self.tree.column("uuid", width=65, anchor="center")
        self.tree.column("nama", width=90, anchor="center")
        self.tree.column("panggilan", width=90, anchor="center")
        self.tree.column("mandarin", width=90, anchor="center")
        self.tree.column("penyebutan", width=90, anchor="center")
        self.tree.column("dari", width=80, anchor="center")
        self.tree.column("keluarga", width=80, anchor="center")
        self.tree.column("keterangan", width=80, anchor="center")
        self.tree.column("tahun", width=50, anchor="center")
        self.tree.column("bulan", width=50, anchor="center")
        self.tree.column("hari", width=50, anchor="center")
        self.tree.column("tanggal", width=110, anchor="center")

        # Scrollbar vertikal
        scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_y.set)

        # Scrollbar horizontal (banyak kolom)
        scrollbar_x = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=scrollbar_x.set)

        self.tree.pack(side="top", fill="both", expand=True, padx=(10, 0), pady=(0, 0))
        scrollbar_y.pack(side="right", fill="y", padx=(0, 10), pady=(0, 0))
        scrollbar_x.pack(side="bottom", fill="x", padx=10, pady=(0, 10))

    # ========================================================
    # Event Handlers
    # ========================================================
    def _on_save(self) -> None:
        """Handler tombol Simpan: Validasi input lalu simpan ke database."""
        panggilan = self.entry_panggilan.get().strip()
        mandarin = self.entry_mandarin.get().strip()
        dari = self.entry_dari.get().strip()
        nama = self.entry_nama.get().strip()
        penyebutan = self.entry_penyebutan.get().strip()
        keluarga = self.entry_keluarga.get().strip()
        keterangan = self.entry_keterangan.get().strip()
        tahun = self.entry_tahun.get().strip()
        bulan = self.entry_bulan.get().strip()
        hari = self.entry_hari.get().strip()

        # Validasi: Field wajib tidak boleh kosong
        if not all([panggilan, mandarin, dari]):
            messagebox.showwarning(
                "Input Tidak Lengkap",
                "Harap isi field wajib:\n"
                "• Panggilan (稱呼)\n• Nama Mandarin\n• Dari (陽上)",
            )
            return

        try:
            record_uuid = insert_record(
                panggilan=panggilan,
                nama_mandarin=mandarin,
                dari=dari,
                nama=nama,
                penyebutan=penyebutan,
                keluarga=keluarga,
                keterangan=keterangan,
                tahun_lunar=tahun,
                bulan_lunar=bulan,
                hari_lunar=hari,
            )
            messagebox.showinfo("Berhasil", f"Data tersimpan.\nUUID: {record_uuid[:8]}...")
            self._clear_inputs()
            self._refresh_table()
        except RuntimeError as e:
            messagebox.showerror("Gagal Menyimpan", str(e))

    def _on_print(self) -> None:
        """Handler tombol Cetak: Ambil record terpilih, generate PDF, buka otomatis."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Tidak Ada Pilihan", "Pilih satu baris data di tabel untuk dicetak.")
            return

        # Ambil data dari baris terpilih (urutan kolom sesuai Treeview)
        item = self.tree.item(selected[0])
        values = item["values"]
        # Kolom: 0=uuid, 1=nama, 2=panggilan, 3=mandarin, 4=penyebutan,
        #         5=dari, 6=keluarga, 7=keterangan, 8=tahun, 9=bulan, 10=hari, 11=tanggal

        data = {
            "nama": values[1],
            "panggilan": values[2],
            "nama_mandarin": values[3],
            "penyebutan": values[4],
            "dari": values[5],
            "keluarga": values[6],
            "keterangan": values[7],
            "tahun_lunar": values[8],
            "bulan_lunar": values[9],
            "hari_lunar": values[10],
        }

        # Ambil offset kalibrasi
        try:
            offset_x = float(self.entry_offset_x.get() or "0.0")
            offset_y = float(self.entry_offset_y.get() or "0.0")
        except ValueError:
            messagebox.showwarning("Offset Tidak Valid", "Offset X dan Y harus berupa angka (misal: 1.5).")
            return

        # Nama file: ritual_<uuid-pendek>.pdf
        short_uuid = str(values[0])[:8]
        filename = f"ritual_{short_uuid}.pdf"
        output_path = os.path.join(_OUTPUT_DIR, filename)

        try:
            result_path = generate_pdf(data, output_path, offset_x=offset_x, offset_y=offset_y)
            messagebox.showinfo("PDF Berhasil", f"File PDF dibuat:\n{result_path}")
            # Buka PDF otomatis di aplikasi default OS
            self._open_file(result_path)
        except FileNotFoundError as e:
            messagebox.showerror("Font Tidak Ditemukan", str(e))
        except RuntimeError as e:
            messagebox.showerror("Gagal Membuat PDF", str(e))

    def _on_delete(self) -> None:
        """Handler tombol Hapus: Hapus record terpilih dari database."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Tidak Ada Pilihan", "Pilih satu baris data di tabel untuk dihapus.")
            return

        item = self.tree.item(selected[0])
        record_uuid = item["values"][0]

        confirm = messagebox.askyesno(
            "Konfirmasi Hapus",
            f"Apakah Anda yakin ingin menghapus record:\nUUID: {record_uuid}?",
        )
        if not confirm:
            return

        try:
            success = delete_record(str(record_uuid))
            if success:
                messagebox.showinfo("Berhasil", "Record berhasil dihapus.")
                self._refresh_table()
            else:
                messagebox.showwarning("Tidak Ditemukan", "Record tidak ditemukan di database.")
        except RuntimeError as e:
            messagebox.showerror("Gagal Menghapus", str(e))

    def _on_calibration_print(self) -> None:
        """Handler tombol Cetak Grid Kalibrasi."""
        output_path = os.path.join(_OUTPUT_DIR, "calibration_grid.pdf")
        try:
            result_path = generate_calibration_pdf(output_path)
            messagebox.showinfo("Kalibrasi", f"PDF kalibrasi dibuat:\n{result_path}")
            self._open_file(result_path)
        except RuntimeError as e:
            messagebox.showerror("Gagal", str(e))

    # ========================================================
    # Utility Methods
    # ========================================================
    def _refresh_table(self) -> None:
        """Memuat ulang seluruh data dari database ke tabel Treeview."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            records = get_all_records()
            for r in records:
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        r["uuid"],
                        r["nama"],
                        r["panggilan"],
                        r["nama_mandarin"],
                        r["penyebutan"],
                        r["dari"],
                        r["keluarga"],
                        r["keterangan"],
                        r["tahun_lunar"],
                        r["bulan_lunar"],
                        r["hari_lunar"],
                        r["created_at"],
                    ),
                )
        except RuntimeError as e:
            messagebox.showerror("Gagal Memuat Data", str(e))

    def _clear_inputs(self) -> None:
        """Mengosongkan semua field input setelah simpan berhasil."""
        self.entry_nama.delete(0, "end")
        self.entry_panggilan.delete(0, "end")
        self.entry_mandarin.delete(0, "end")
        self.entry_penyebutan.delete(0, "end")
        self.entry_dari.delete(0, "end")
        self.entry_keluarga.delete(0, "end")
        self.entry_keterangan.delete(0, "end")
        self.entry_tahun.delete(0, "end")
        self.entry_bulan.delete(0, "end")
        self.entry_hari.delete(0, "end")

    def _on_import_excel(self) -> None:
        """Handler tombol Import Excel: Pilih file .xlsx, import Sheet 2 ke DB."""
        filepath = filedialog.askopenfilename(
            title="Pilih File Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )
        if not filepath:
            return

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
