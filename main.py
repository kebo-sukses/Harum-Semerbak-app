# -*- coding: utf-8 -*-
"""
main.py — Aplikasi Desktop "Fill-in-the-Blanks" Formulir Ritual
Framework: CustomTkinter (modern Tkinter)

Fitur:
  1. Form Input  : Entry fields untuk Nama Mendiang, Pengirim, Tanggal Lunar.
  2. Tabel Preview: Treeview menampilkan data tersimpan di SQLite.
  3. Tombol Cetak : Generate PDF layer transparan & buka otomatis.
  4. Calibration  : Offset X/Y untuk koreksi posisi cetak printer.
  5. Tombol Kalibrasi: Cetak halaman grid kalibrasi.
"""

import os
import sys
import subprocess
from tkinter import ttk, messagebox

import customtkinter as ctk

# ============================================================
# Import modul internal
# ============================================================
# Pastikan root project ada di sys.path
_ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if _ROOT_DIR not in sys.path:
    sys.path.insert(0, _ROOT_DIR)

from database.database import init_db, insert_record, get_all_records, delete_record
from modules.pdf_engine import generate_pdf, generate_calibration_pdf


# ============================================================
# Konstanta UI
# ============================================================
APP_TITLE = "Formulir Ritual — Fill-in-the-Blanks (F4)"
APP_WIDTH = 960
APP_HEIGHT = 680

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
        """Membangun frame form input di bagian atas."""
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(padx=15, pady=(15, 5), fill="x")

        title_label = ctk.CTkLabel(
            frame, text="📝 Data Formulir Ritual", font=ctk.CTkFont(size=16, weight="bold")
        )
        title_label.grid(row=0, column=0, columnspan=4, padx=10, pady=(10, 5), sticky="w")

        # --- Baris 1: Nama Mendiang & Nama Pengirim ---
        ctk.CTkLabel(frame, text="Nama Mendiang (往生者):").grid(
            row=1, column=0, padx=(10, 5), pady=5, sticky="e"
        )
        self.entry_mendiang = ctk.CTkEntry(frame, width=250, placeholder_text="contoh: 蔡氏先人")
        self.entry_mendiang.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ctk.CTkLabel(frame, text="Nama Pengirim (陽上):").grid(
            row=1, column=2, padx=(20, 5), pady=5, sticky="e"
        )
        self.entry_pengirim = ctk.CTkEntry(frame, width=250, placeholder_text="contoh: 蔡明志")
        self.entry_pengirim.grid(row=1, column=3, padx=5, pady=5, sticky="w")

        # --- Baris 2: Tahun, Bulan, Hari Lunar ---
        ctk.CTkLabel(frame, text="Tahun Lunar (太歲):").grid(
            row=2, column=0, padx=(10, 5), pady=5, sticky="e"
        )
        self.entry_tahun = ctk.CTkEntry(frame, width=120, placeholder_text="contoh: 乙巳")
        self.entry_tahun.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        ctk.CTkLabel(frame, text="Bulan Lunar:").grid(
            row=2, column=2, padx=(20, 5), pady=5, sticky="e"
        )
        self.entry_bulan = ctk.CTkEntry(frame, width=120, placeholder_text="contoh: 正月")
        self.entry_bulan.grid(row=2, column=3, padx=5, pady=5, sticky="w")

        ctk.CTkLabel(frame, text="Hari Lunar:").grid(
            row=3, column=0, padx=(10, 5), pady=5, sticky="e"
        )
        self.entry_hari = ctk.CTkEntry(frame, width=120, placeholder_text="contoh: 十五")
        self.entry_hari.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # --- Baris 3: Catatan (opsional) ---
        ctk.CTkLabel(frame, text="Catatan:").grid(
            row=3, column=2, padx=(20, 5), pady=5, sticky="e"
        )
        self.entry_catatan = ctk.CTkEntry(frame, width=250, placeholder_text="opsional")
        self.entry_catatan.grid(row=3, column=3, padx=5, pady=5, sticky="w")

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
        """Membangun baris tombol aksi: Simpan, Cetak, Hapus."""
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(padx=15, pady=5, fill="x")

        ctk.CTkButton(
            frame, text="💾 Simpan Data", width=160, command=self._on_save
        ).pack(side="left", padx=10, pady=10)

        ctk.CTkButton(
            frame, text="🖨️ Cetak PDF", width=160, fg_color="green", command=self._on_print
        ).pack(side="left", padx=10, pady=10)

        ctk.CTkButton(
            frame, text="🗑️ Hapus Terpilih", width=160, fg_color="red", command=self._on_delete
        ).pack(side="left", padx=10, pady=10)

        ctk.CTkButton(
            frame, text="🔄 Refresh Tabel", width=160, command=self._refresh_table
        ).pack(side="right", padx=10, pady=10)

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
        columns = ("uuid", "mendiang", "pengirim", "tahun", "bulan", "hari", "catatan", "tanggal")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=10)

        # Header kolom
        self.tree.heading("uuid", text="ID")
        self.tree.heading("mendiang", text="Nama Mendiang")
        self.tree.heading("pengirim", text="Nama Pengirim")
        self.tree.heading("tahun", text="Tahun")
        self.tree.heading("bulan", text="Bulan")
        self.tree.heading("hari", text="Hari")
        self.tree.heading("catatan", text="Catatan")
        self.tree.heading("tanggal", text="Dibuat")

        # Lebar kolom
        self.tree.column("uuid", width=80, anchor="center")
        self.tree.column("mendiang", width=120, anchor="center")
        self.tree.column("pengirim", width=120, anchor="center")
        self.tree.column("tahun", width=60, anchor="center")
        self.tree.column("bulan", width=60, anchor="center")
        self.tree.column("hari", width=60, anchor="center")
        self.tree.column("catatan", width=140, anchor="w")
        self.tree.column("tanggal", width=130, anchor="center")

        # Scrollbar
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=(0, 10))
        scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=(0, 10))

    # ========================================================
    # Event Handlers
    # ========================================================
    def _on_save(self) -> None:
        """Handler tombol Simpan: Validasi input lalu simpan ke database."""
        mendiang = self.entry_mendiang.get().strip()
        pengirim = self.entry_pengirim.get().strip()
        tahun = self.entry_tahun.get().strip()
        bulan = self.entry_bulan.get().strip()
        hari = self.entry_hari.get().strip()
        catatan = self.entry_catatan.get().strip()

        # Validasi: Field wajib tidak boleh kosong
        if not all([mendiang, pengirim, tahun, bulan, hari]):
            messagebox.showwarning(
                "Input Tidak Lengkap",
                "Harap isi semua field wajib:\n"
                "• Nama Mendiang\n• Nama Pengirim\n"
                "• Tahun Lunar\n• Bulan Lunar\n• Hari Lunar",
            )
            return

        try:
            record_uuid = insert_record(
                nama_mendiang=mendiang,
                nama_pengirim=pengirim,
                tahun_lunar=tahun,
                bulan_lunar=bulan,
                hari_lunar=hari,
                catatan=catatan,
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

        # Ambil data dari baris terpilih
        item = self.tree.item(selected[0])
        values = item["values"]

        data = {
            "nama_mendiang": values[1],     # Kolom "mendiang"
            "nama_pengirim": values[2],     # Kolom "pengirim"
            "tahun_lunar": values[3],       # Kolom "tahun"
            "bulan_lunar": values[4],       # Kolom "bulan"
            "hari_lunar": values[5],        # Kolom "hari"
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
        # Hapus semua baris lama
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
                        r["nama_mendiang"],
                        r["nama_pengirim"],
                        r["tahun_lunar"],
                        r["bulan_lunar"],
                        r["hari_lunar"],
                        r["catatan"],
                        r["created_at"],
                    ),
                )
        except RuntimeError as e:
            messagebox.showerror("Gagal Memuat Data", str(e))

    def _clear_inputs(self) -> None:
        """Mengosongkan semua field input setelah simpan berhasil."""
        self.entry_mendiang.delete(0, "end")
        self.entry_pengirim.delete(0, "end")
        self.entry_tahun.delete(0, "end")
        self.entry_bulan.delete(0, "end")
        self.entry_hari.delete(0, "end")
        self.entry_catatan.delete(0, "end")

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
