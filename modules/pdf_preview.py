# -*- coding: utf-8 -*-
"""
pdf_preview.py — Jendela preview PDF dan cetak langsung ke printer.

Menampilkan Toplevel window berisi:
  - Render visual halaman PDF (menggunakan PyMuPDF/fitz)
  - Tombol Cetak ke printer default / pilih printer
  - Tombol tutup

Tidak menyimpan file ke folder output — menggunakan file temporer.
"""

import os
import tempfile
import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


class PDFPreviewWindow(ctk.CTkToplevel):
    """Jendela preview PDF dengan opsi cetak langsung ke printer."""

    def __init__(
        self,
        master,
        pdf_bytes: bytes,
        title_text: str = "Preview PDF",
    ) -> None:
        """Inisialisasi jendela preview.

        Args:
            master:      Parent widget.
            pdf_bytes:   Data PDF dalam bytes (hasil generate di memori).
            title_text:  Judul jendela.
        """
        super().__init__(master)
        self.title(title_text)
        self.resizable(True, True)

        self._pdf_bytes = pdf_bytes
        self._photo_image = None
        self._temp_path = None

        self._build_ui()
        self._render_preview()

        # Fokuskan window
        self.after(100, self.focus_force)
        self.after(100, self.lift)

    def _build_ui(self) -> None:
        """Bangun layout UI: preview area + tombol."""
        # --- Frame atas: tombol ---
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(10, 5))

        ctk.CTkButton(
            btn_frame,
            text="🖨️  Cetak",
            width=140,
            fg_color="#1565C0",
            hover_color="#0D47A1",
            command=self._on_print,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_frame,
            text="💾  Simpan PDF",
            width=140,
            fg_color="#2E7D32",
            hover_color="#1B5E20",
            command=self._on_save,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_frame,
            text="✖  Tutup",
            width=100,
            fg_color="#757575",
            hover_color="#616161",
            command=self.destroy,
        ).pack(side="right")

        # --- Frame bawah: scroll canvas untuk preview ---
        preview_frame = ctk.CTkFrame(self)
        preview_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        self._canvas = tk.Canvas(
            preview_frame,
            bg="#E0E0E0",
            highlightthickness=0,
        )
        v_scroll = ttk.Scrollbar(
            preview_frame, orient="vertical", command=self._canvas.yview,
        )
        h_scroll = ttk.Scrollbar(
            preview_frame, orient="horizontal", command=self._canvas.xview,
        )
        self._canvas.configure(
            yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set,
        )

        self._canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

    def _render_preview(self) -> None:
        """Render halaman pertama PDF sebagai gambar di canvas."""
        if fitz is None:
            self._canvas.create_text(
                300, 200,
                text="PyMuPDF (fitz) tidak terinstall.\n"
                     "Jalankan: pip install PyMuPDF",
                font=("Arial", 14),
                fill="red",
            )
            self.geometry("650x450")
            return

        try:
            doc = fitz.open(stream=self._pdf_bytes, filetype="pdf")
            page = doc[0]

            # Render di 1.5x zoom untuk preview jelas
            zoom = 1.5
            matrix = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix)
            doc.close()

            # Konversi ke PIL Image lalu ke PhotoImage
            if Image is not None and ImageTk is not None:
                img = Image.frombytes(
                    "RGB", (pix.width, pix.height), pix.samples,
                )
                self._photo_image = ImageTk.PhotoImage(img)
            else:
                # Fallback tanpa Pillow — gunakan PPM
                ppm_data = pix.tobytes("ppm")
                self._photo_image = tk.PhotoImage(data=ppm_data)

            self._canvas.create_image(
                0, 0, anchor="nw", image=self._photo_image,
            )
            self._canvas.configure(
                scrollregion=(0, 0, pix.width, pix.height),
            )

            # Sesuaikan ukuran window
            win_w = min(pix.width + 40, 960)
            win_h = min(pix.height + 100, 820)
            self.geometry(f"{win_w}x{win_h}")

        except Exception as e:
            self._canvas.create_text(
                300, 200,
                text=f"Gagal render preview:\n{e}",
                font=("Arial", 12),
                fill="red",
            )
            self.geometry("650x450")

    def _get_temp_path(self) -> str:
        """Buat file PDF sementara untuk dicetak/disimpan."""
        if self._temp_path and os.path.isfile(self._temp_path):
            return self._temp_path

        fd, path = tempfile.mkstemp(suffix=".pdf", prefix="ritual_preview_")
        os.write(fd, self._pdf_bytes)
        os.close(fd)
        self._temp_path = path
        return path

    def _on_print(self) -> None:
        """Cetak PDF ke printer — multi-strategy untuk kompatibilitas Windows."""
        if os.name != "nt":
            self._print_unix()
            return

        # --- Strategy 1: langsung cetak PDF via ShellExecute "print" ---
        try:
            temp_path = self._get_temp_path()
            os.startfile(temp_path, "print")
            return  # sukses — dialog cetak muncul
        except OSError:
            pass  # WinError 1155 — tidak ada app terdaftar

        # --- Strategy 2: render ke gambar lalu cetak (Windows selalu bisa
        #     mencetak gambar melalui Photos / Photo Viewer) ---
        try:
            img_path = self._render_print_image()
            os.startfile(img_path, "print")
            return
        except OSError:
            pass

        # --- Strategy 3: buka PDF di aplikasi default, user cetak manual ---
        try:
            temp_path = self._get_temp_path()
            os.startfile(temp_path)
            messagebox.showinfo(
                "Cetak Manual",
                "Tidak ada aplikasi yang bisa mencetak PDF secara otomatis.\n\n"
                "File telah dibuka di aplikasi default.\n"
                "Gunakan  Ctrl + P  untuk mencetak.",
                parent=self,
            )
        except OSError as exc:
            messagebox.showerror(
                "Gagal Mencetak",
                f"Tidak dapat membuka file untuk dicetak:\n{exc}",
                parent=self,
            )

    def _print_unix(self) -> None:
        """Cetak di Linux/macOS menggunakan lp."""
        import subprocess

        try:
            temp_path = self._get_temp_path()
            subprocess.run(["lp", temp_path], check=True)
            messagebox.showinfo(
                "Cetak", "Dokumen dikirim ke printer.", parent=self,
            )
        except Exception as exc:
            messagebox.showerror(
                "Gagal Mencetak", f"Error: {exc}", parent=self,
            )

    def _render_print_image(self) -> str:
        """Render halaman PDF ke PNG 300 DPI untuk dicetak sebagai gambar.

        Returns:
            Path ke file gambar sementara.

        Raises:
            RuntimeError: Jika fitz tidak tersedia atau render gagal.
        """
        if fitz is None:
            raise RuntimeError("PyMuPDF tidak terinstall.")

        doc = fitz.open(stream=self._pdf_bytes, filetype="pdf")
        page = doc[0]
        # 300 DPI ÷ 72 PDF-pt/inch ≈ 4.17× zoom → kualitas cetak bagus
        zoom = 300.0 / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        doc.close()

        fd, img_path = tempfile.mkstemp(suffix=".png", prefix="ritual_print_")
        os.close(fd)
        pix.save(img_path)
        self._temp_print_img = img_path
        return img_path

    def _on_save(self) -> None:
        """Simpan PDF ke lokasi yang dipilih user."""
        from tkinter import filedialog

        file_path = filedialog.asksaveasfilename(
            parent=self,
            title="Simpan PDF",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
        )
        if not file_path:
            return

        try:
            with open(file_path, "wb") as f:
                f.write(self._pdf_bytes)
            messagebox.showinfo(
                "Berhasil",
                f"PDF disimpan:\n{file_path}",
                parent=self,
            )
        except Exception as e:
            messagebox.showerror(
                "Gagal Menyimpan", f"Error: {e}", parent=self,
            )

    def destroy(self) -> None:
        """Bersihkan file temporer saat window ditutup."""
        for path in (self._temp_path, getattr(self, "_temp_print_img", None)):
            if path and os.path.isfile(path):
                try:
                    os.remove(path)
                except OSError:
                    pass
        super().destroy()


# Perlu import ttk untuk Scrollbar
from tkinter import ttk  # noqa: E402
