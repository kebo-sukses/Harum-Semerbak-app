# -*- coding: utf-8 -*-
"""
calibration_preview.py — Preview visual kalibrasi posisi teks pada kertas A4.

Menampilkan Toplevel window berisi:
  - Canvas menggambarkan kertas A4 (210×297mm) berskala
  - Grid titik setiap 10mm, label koordinat setiap 50mm
  - Marker posisi teks (berwarna) untuk setiap field formulir
  - Kontrol offset X/Y real-time dengan tombol +/−
  - Tombol cetak PDF kalibrasi

Koordinat field disinkronkan dengan modules/pdf_engine.py.
"""

import tkinter as tk
import customtkinter as ctk


# ============================================================
# Konstanta Kertas A4 (mm)
# ============================================================
A4_W = 210
A4_H = 297


# ============================================================
# Posisi field teks — HARUS SINKRON dengan pdf_engine.py
# ============================================================
FIELDS = [
    {
        "label": "Panggilan + Mandarin",
        "desc": "X=68, Y=200, vertikal ↓12mm",
        "sample": "母親許門梁氏橋玉",
        "x": 68, "y": 200,
        "vertical": True,
        "spacing": 12,
        "color": "#1565C0",
    },
    {
        "label": "Dari + Keterangan",
        "desc": "X=42, Y=170, vertikal ↓12mm",
        "sample": "孝男合家敬奉",
        "x": 42, "y": 170,
        "vertical": True,
        "spacing": 12,
        "color": "#2E7D32",
    },
]


# ============================================================
# Kelas Preview Kalibrasi
# ============================================================
class CalibrationPreview(ctk.CTkToplevel):
    """Window preview visual posisi teks pada kertas formulir ritual A4."""

    SCALE = 1.5     # pixel per mm
    PAD = 30        # padding sekeliling kertas (px)

    def __init__(
        self,
        master,
        offset_x: float = 0.0,
        offset_y: float = 0.0,
        on_print_calibration=None,
    ) -> None:
        """
        Args:
            master:              Parent widget.
            offset_x:            Offset X awal (mm).
            offset_y:            Offset Y awal (mm).
            on_print_calibration: Callback(offset_x, offset_y) untuk cetak PDF.
        """
        super().__init__(master)
        self.title("Kalibrasi Visual - Layout Kertas A4 (210x297mm)")
        self.resizable(False, False)

        self._offset_x = offset_x
        self._offset_y = offset_y
        self._on_print_calibration = on_print_calibration

        self._build_canvas()
        self._build_right_panel()
        self._draw()

        # Fokuskan window ini
        self.after(100, self.focus_force)

    # --------------------------------------------------------
    # Properti ukuran canvas
    # --------------------------------------------------------
    @property
    def _canvas_w(self) -> int:
        return int(A4_W * self.SCALE) + 2 * self.PAD

    @property
    def _canvas_h(self) -> int:
        return int(A4_H * self.SCALE) + 2 * self.PAD

    # --------------------------------------------------------
    # Konversi koordinat mm (origin bawah-kiri, seperti ReportLab)
    # -> pixel canvas (origin atas-kiri, seperti Tkinter)
    # --------------------------------------------------------
    def _mm_to_px(self, x_mm: float, y_mm: float) -> tuple[float, float]:
        px = self.PAD + x_mm * self.SCALE
        py = self.PAD + (A4_H - y_mm) * self.SCALE
        return px, py

    # ========================================================
    # Build: Canvas (sisi kiri)
    # ========================================================
    def _build_canvas(self) -> None:
        self._cv = tk.Canvas(
            self,
            width=self._canvas_w,
            height=self._canvas_h,
            bg="#F5F5F5",
            highlightthickness=1,
            highlightbackground="#999",
        )
        self._cv.pack(side="left", padx=(10, 5), pady=10)

    # ========================================================
    # Build: Panel kontrol (sisi kanan)
    # ========================================================
    def _build_right_panel(self) -> None:
        panel = ctk.CTkFrame(self, width=255)
        panel.pack(side="right", fill="y", padx=(5, 10), pady=10)
        panel.pack_propagate(False)

        # --- Judul ---
        ctk.CTkLabel(
            panel,
            text="Kalibrasi Offset",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(padx=10, pady=(10, 8))

        # --- Legenda warna ---
        ctk.CTkLabel(
            panel, text="Legenda Field:",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(padx=10, pady=(5, 3), anchor="w")

        for field in FIELDS:
            row = ctk.CTkFrame(panel, fg_color="transparent")
            row.pack(fill="x", padx=10, pady=2)
            # Kotak warna
            box = tk.Canvas(row, width=14, height=14, highlightthickness=0)
            box.pack(side="left", padx=(0, 6))
            box.create_rectangle(1, 1, 13, 13, fill=field["color"], outline="")
            # Label
            ctk.CTkLabel(
                row,
                text=field["label"],
                font=ctk.CTkFont(size=11),
            ).pack(side="left")

        # --- Separator ---
        sep = ctk.CTkFrame(panel, height=2, fg_color="gray60")
        sep.pack(fill="x", padx=10, pady=8)

        # --- Offset X ---
        ctk.CTkLabel(
            panel, text="Offset X (mm):", font=ctk.CTkFont(size=12),
        ).pack(padx=10, pady=(2, 2), anchor="w")

        ox_row = ctk.CTkFrame(panel, fg_color="transparent")
        ox_row.pack(fill="x", padx=10, pady=2)
        ctk.CTkButton(
            ox_row, text="-1", width=36, command=lambda: self._nudge("x", -1),
        ).pack(side="left", padx=(0, 2))
        ctk.CTkButton(
            ox_row, text="-.5", width=36, command=lambda: self._nudge("x", -0.5),
        ).pack(side="left", padx=2)
        self._entry_ox = ctk.CTkEntry(ox_row, width=65, justify="center")
        self._entry_ox.pack(side="left", padx=4)
        self._entry_ox.insert(0, f"{self._offset_x:.1f}")
        ctk.CTkButton(
            ox_row, text="+.5", width=36, command=lambda: self._nudge("x", 0.5),
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            ox_row, text="+1", width=36, command=lambda: self._nudge("x", 1),
        ).pack(side="left", padx=(2, 0))

        # --- Offset Y ---
        ctk.CTkLabel(
            panel, text="Offset Y (mm):", font=ctk.CTkFont(size=12),
        ).pack(padx=10, pady=(8, 2), anchor="w")

        oy_row = ctk.CTkFrame(panel, fg_color="transparent")
        oy_row.pack(fill="x", padx=10, pady=2)
        ctk.CTkButton(
            oy_row, text="-1", width=36, command=lambda: self._nudge("y", -1),
        ).pack(side="left", padx=(0, 2))
        ctk.CTkButton(
            oy_row, text="-.5", width=36, command=lambda: self._nudge("y", -0.5),
        ).pack(side="left", padx=2)
        self._entry_oy = ctk.CTkEntry(oy_row, width=65, justify="center")
        self._entry_oy.pack(side="left", padx=4)
        self._entry_oy.insert(0, f"{self._offset_y:.1f}")
        ctk.CTkButton(
            oy_row, text="+.5", width=36, command=lambda: self._nudge("y", 0.5),
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            oy_row, text="+1", width=36, command=lambda: self._nudge("y", 1),
        ).pack(side="left", padx=(2, 0))

        # --- Tombol Update ---
        ctk.CTkButton(
            panel, text="Update Preview", width=220,
            command=self._on_update,
        ).pack(padx=10, pady=(15, 5))

        # --- Tombol Cetak PDF ---
        ctk.CTkButton(
            panel, text="Cetak PDF Kalibrasi", width=220,
            fg_color="green", command=self._on_print_click,
        ).pack(padx=10, pady=5)

        # --- Info ---
        ctk.CTkLabel(
            panel,
            text=(
                "* Grid titik setiap 10mm\n"
                "* Label koordinat tiap 50mm\n"
                "* Klik +/-0.5 atau +/-1 untuk geser\n"
                "* Update Preview untuk refresh\n"
                "* Cetak PDF untuk print grid"
            ),
            font=ctk.CTkFont(size=10),
            text_color="gray50",
            justify="left",
        ).pack(padx=10, pady=(12, 5), anchor="w")

    # ========================================================
    # Event: Tombol nudge offset
    # ========================================================
    def _nudge(self, axis: str, delta: float) -> None:
        entry = self._entry_ox if axis == "x" else self._entry_oy
        try:
            val = float(entry.get() or "0")
        except ValueError:
            val = 0.0
        val += delta
        entry.delete(0, "end")
        entry.insert(0, f"{val:.1f}")
        self._on_update()

    # ========================================================
    # Event: Update preview
    # ========================================================
    def _on_update(self) -> None:
        try:
            self._offset_x = float(self._entry_ox.get() or "0")
            self._offset_y = float(self._entry_oy.get() or "0")
        except ValueError:
            return
        self._draw()

    # ========================================================
    # Event: Cetak PDF
    # ========================================================
    def _on_print_click(self) -> None:
        self._on_update()  # pastikan offset terbaru
        if self._on_print_calibration:
            self._on_print_calibration(self._offset_x, self._offset_y)

    # ========================================================
    # Getter offset
    # ========================================================
    def get_offsets(self) -> tuple[float, float]:
        """Return (offset_x, offset_y) terkini."""
        return self._offset_x, self._offset_y

    # ========================================================
    # Render Canvas
    # ========================================================
    def _draw(self) -> None:
        """Menggambar ulang seluruh canvas: kertas, grid, dan marker field."""
        cv = self._cv
        cv.delete("all")

        ox = self._offset_x
        oy = self._offset_y

        # --- 1. Outline kertas A4 (latar merah muda = simulasi kertas merah) ---
        x0, y0 = self._mm_to_px(0, A4_H)
        x1, y1 = self._mm_to_px(A4_W, 0)
        cv.create_rectangle(x0, y0, x1, y1, fill="#FFEBEE", outline="#333", width=2)

        # --- 2. Grid titik setiap 10mm ---
        for gx in range(0, A4_W + 1, 10):
            for gy in range(0, A4_H + 1, 10):
                px, py = self._mm_to_px(gx, gy)
                cv.create_oval(px - 1, py - 1, px + 1, py + 1, fill="#BDBDBD", outline="")

                # Label setiap 50mm
                if gx % 50 == 0 and gy % 50 == 0:
                    cv.create_text(
                        px + 8, py - 6,
                        text=f"{gx},{gy}",
                        font=("Consolas", 6),
                        fill="#9E9E9E",
                    )

        # --- 3. Garis bantu sumbu di tepi kertas (setiap 50mm) ---
        for gx in range(0, A4_W + 1, 50):
            px_top, py_top = self._mm_to_px(gx, A4_H)
            px_bot, py_bot = self._mm_to_px(gx, 0)
            cv.create_line(px_top, py_top, px_bot, py_bot, fill="#E0E0E0", dash=(1, 4))

        for gy in range(0, A4_H + 1, 50):
            px_l, py_l = self._mm_to_px(0, gy)
            px_r, py_r = self._mm_to_px(A4_W, gy)
            cv.create_line(px_l, py_l, px_r, py_r, fill="#E0E0E0", dash=(1, 4))

        # --- 4. Marker posisi setiap field ---
        for field in FIELDS:
            self._draw_field(cv, field, ox, oy)

        # --- 5. Judul atas ---
        cv.create_text(
            self._canvas_w // 2, 12,
            text=f"Kertas A4 (210x297mm)  |  Offset X={ox:+.1f}mm  Y={oy:+.1f}mm",
            font=("Consolas", 9, "bold"),
            fill="#333",
        )

    # --------------------------------------------------------
    # Menggambar satu field (marker + sample teks)
    # --------------------------------------------------------
    def _draw_field(
        self,
        cv: tk.Canvas,
        field: dict,
        ox: float,
        oy: float,
    ) -> None:
        """Menggambar marker posisi dan sample teks untuk satu field."""
        fx = field["x"] + ox
        fy = field["y"] + oy
        color = field["color"]
        sample = field["sample"]
        spacing = field["spacing"]

        # Font Mandarin
        cjk_font = ("SimSun", 9)

        half = 5.5 * self.SCALE  # setengah sisi kotak karakter

        if field["vertical"]:
            # --- Teks vertikal: karakter atas-ke-bawah ---
            for i, char in enumerate(sample):
                cx, cy = self._mm_to_px(fx, fy - i * spacing)
                # Kotak karakter (garis putus-putus)
                cv.create_rectangle(
                    cx - half, cy - half, cx + half, cy + half,
                    outline=color, width=1, dash=(3, 2),
                )
                # Karakter
                cv.create_text(cx, cy, text=char, font=cjk_font, fill=color)

            # Label field di kanan atas karakter pertama
            lx, ly = self._mm_to_px(fx, fy)
            cv.create_text(
                lx + 10 * self.SCALE, ly - 2,
                text=field["label"],
                font=("Arial", 7, "bold"),
                fill=color, anchor="w",
            )
        else:
            # --- Teks horizontal: karakter kiri-ke-kanan ---
            for i, char in enumerate(sample):
                cx, cy = self._mm_to_px(fx + i * spacing, fy)
                cv.create_rectangle(
                    cx - half, cy - half, cx + half, cy + half,
                    outline=color, width=1, dash=(3, 2),
                )
                cv.create_text(cx, cy, text=char, font=cjk_font, fill=color)

            # Label field di atas karakter pertama
            lx, ly = self._mm_to_px(fx, fy)
            cv.create_text(
                lx, ly - 8 * self.SCALE,
                text=field["label"],
                font=("Arial", 7, "bold"),
                fill=color, anchor="w",
            )
