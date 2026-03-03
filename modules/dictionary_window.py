# -*- coding: utf-8 -*-
"""
dictionary_window.py — Kamus Mini Aksara Mandarin

Menampilkan daftar lengkap:
  1. Penerima (Panggilan / 稱呼) — aksara mandarin, pinyin, terjemahan Indonesia
  2. Pemberi (Dari / 陽上)       — aksara mandarin, pinyin, terjemahan Indonesia
  3. Angka Mandarin 1–9

Setiap aksara mandarin bisa di-copy ke clipboard dengan satu klik.
Pengguna dapat menambahkan entri kustom yang tersimpan di file JSON.
"""

import json
import os
import tkinter as tk
from tkinter import ttk, messagebox

import customtkinter as ctk
from pypinyin import pinyin, Style as PinyinStyle

# ============================================================
# Path file JSON untuk entri kustom
# ============================================================
_CUSTOM_DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "database")
_CUSTOM_DATA_FILE = os.path.join(_CUSTOM_DATA_DIR, "custom_dictionary.json")


def _load_custom_data() -> dict[str, list[tuple[str, str]]]:
    """Muat entri kustom dari file JSON."""
    default = {"panggilan": [], "dari": [], "aksara_umum": []}
    if not os.path.exists(_CUSTOM_DATA_FILE):
        return default
    try:
        with open(_CUSTOM_DATA_FILE, "r", encoding="utf-8") as fh:
            raw = json.load(fh)
        result = {}
        for key in ("panggilan", "dari", "aksara_umum"):
            result[key] = [
                (item[0], item[1])
                for item in raw.get(key, [])
                if isinstance(item, list) and len(item) == 2
            ]
        return result
    except (json.JSONDecodeError, OSError):
        return default


def _save_custom_data(data: dict[str, list[tuple[str, str]]]) -> None:
    """Simpan entri kustom ke file JSON."""
    os.makedirs(_CUSTOM_DATA_DIR, exist_ok=True)
    serializable = {
        key: [list(pair) for pair in entries]
        for key, entries in data.items()
    }
    with open(_CUSTOM_DATA_FILE, "w", encoding="utf-8") as fh:
        json.dump(serializable, fh, ensure_ascii=False, indent=2)


# ============================================================
# Data Penerima (Panggilan)
# ============================================================
_PANGGILAN_DATA: list[tuple[str, str]] = [
    ("Ayah", "父親"),
    ("Kakek (Pihak Ayah)", "祖父"),
    ("Nenek (Pihak Ayah)", "祖母"),
    ("Kakak Laki-laki Ayah", "伯父"),
    ("Istri Kakak Laki-laki Ayah", "伯母"),
    ("Adik Laki-laki Ayah", "叔父"),
    ("Istri Adik Laki-laki Ayah", "嬸母"),
    ("Saudara Perempuan Ayah (Bibi)", "姑母"),
    ("Suami Saudara Perempuan Ayah", "姑丈"),
    ("Kakak Dari Kakek", "太伯祖考"),
    ("Ibu", "母親"),
    ("Kakek (Pihak Ibu)", "外祖父"),
    ("Nenek (Pihak Ibu)", "外祖母"),
    ("Saudara Laki-laki Ibu", "舅父"),
    ("Istri Saudara Laki-laki Ibu", "舅母"),
    ("Saudara Perempuan Ibu", "姨母"),
    ("Suami Saudara Perempuan Ibu", "姨丈"),
    ("Kakak Laki-laki (長)", "亡胞長兄"),
    ("Kakak Laki-laki (二)", "亡胞二兄"),
    ("Kakak Laki-laki (三)", "亡胞三兄"),
    ("Kakak Perempuan (長)", "亡胞長姊"),
    ("Kakak Perempuan (二)", "亡胞二姊"),
    ("Kakak Perempuan (三)", "亡胞三姊"),
    ("Adik Laki-laki (大)", "亡胞大弟"),
    ("Adik Laki-laki (二)", "亡胞二弟"),
    ("Adik Laki-laki (三)", "亡胞三弟"),
    ("Adik Perempuan (大)", "亡胞大妹"),
    ("Adik Perempuan (二)", "亡胞二妹"),
    ("Adik Perempuan (三)", "亡胞三妹"),
    ("Kakak/Adik Ipar Laki-laki", "姊夫"),
    ("Saudara Laki-laki Istri", "内兄"),
    ("Suami Saudara Perempuan Istri", "襟兄"),
    ("Kakak/Adik Ipar Perempuan", "嫂嫂"),
    ("Saudara Perempuan Istri", "内姊"),
    ("Istri Saudara Laki-laki Istri", "内嫂"),
    ("Ayah Angkat", "養父"),
    ("Ibu Angkat", "養母"),
    ("Kakek Angkat (Ayah)", "養祖父"),
    ("Nenek Angkat (Ayah)", "養祖母"),
    ("Kakek Angkat (Ibu)", "養外祖父"),
    ("Nenek Angkat (Ibu)", "養外祖母"),
    ("Menantu Laki-laki", "女婿"),
    ("Menantu Perempuan", "媳妇"),
]

# ============================================================
# Data Pemberi (Dari)
# ============================================================
_DARI_DATA: list[tuple[str, str]] = [
    ("Anak Lk dan Pr", "众孝眷"),
    ("Anak Ke-1 (Lk)", "孝長子"),
    ("Anak Ke-2 (Lk)", "孝次子"),
    ("Anak Ke-3 (Lk)", "孝三子"),
    ("Anak Ke-1 (Pr)", "孝長女"),
    ("Anak Ke-2 (Pr)", "孝次女"),
    ("Anak Ke-3 (Pr)", "孝三女"),
    ("Anak Bungsu (Lk)", "孝幼子"),
    ("Anak Bungsu (Pr)", "孝幼女"),
    ("Adik 1 (Lk)", "愚大弟"),
    ("Adik 2 (Lk)", "愚二弟"),
    ("Adik 3 (Lk)", "愚三弟"),
    ("Adik 1 (Pr)", "愚大妹"),
    ("Adik 2 (Pr)", "愚二妹"),
    ("Adik 3 (Pr)", "愚三妹"),
    ("Kakak 1 (Lk)", "胞長兄"),
    ("Kakak 2 (Lk)", "胞二兄"),
    ("Kakak 3 (Lk)", "胞三兄"),
    ("Kakak 1 (Pr)", "胞長姊"),
    ("Kakak 2 (Pr)", "胞二姊"),
    ("Kakak 3 (Pr)", "胞三姊"),
    ("Anak Angkat (Lk)", "孝養子"),
    ("Anak Angkat (Pr)", "孝養女"),
    ("Anak Tiri (Lk)", "孝繼子"),
    ("Anak Tiri (Pr)", "孝繼女"),
    ("Menantu (Lk)", "孝女婿"),
    ("Menantu (Pr)", "孝兒媳"),
    ("Cucu Kandung (Lk)", "孝孫"),
    ("Cucu Kandung (Pr)", "孝孫女"),
    ("Cucu Luar (Lk)", "孝外孫"),
    ("Cucu Luar (Pr)", "孝外孫女"),
    ("Keponakan Sdr Lk (Lk)", "孝姪"),
    ("Keponakan Sdr Lk (Pr)", "孝姪女"),
    ("Keponakan Sdr Pr (Lk)", "孝外甥"),
    ("Keponakan Sdr Pr (Pr)", "孝外甥女"),
    ("Saudara Ipar Adik (Lk)", "孝內弟"),
    ("Saudara Ipar Adik (Pr)", "孝內姊"),
    ("Saudara Ipar Kakak (Lk)", "内兄"),
    ("Saudara Ipar Kakak (Pr)", "内妹"),
    ("Frasa: 众孝眷 偕 合家敬奉", "众孝眷 偕 合家敬奉"),
    ("Frasa: 孝子贤孙 偕 合家敬奉", "孝子贤孙 偕 合家敬奉"),
    ("Frasa: 合家敬奉 叩首", "合家敬奉 叩首"),
]

# ============================================================
# Angka Mandarin 1–9
# ============================================================
_ANGKA_DATA: list[tuple[str, str]] = [
    ("1", "一"),
    ("2", "二"),
    ("3", "三"),
    ("4", "四"),
    ("5", "五"),
    ("6", "六"),
    ("7", "七"),
    ("8", "八"),
    ("9", "九"),
]

# Angka urutan (untuk anak, kakak, adik)
_ANGKA_URUTAN_DATA: list[tuple[str, str]] = [
    ("Ke-1 / Pertama", "長"),
    ("Ke-2 / Kedua", "次"),
    ("Ke-2 (umum)", "二"),
    ("Ke-3", "三"),
    ("Ke-4", "四"),
    ("Ke-5", "五"),
    ("Ke-6", "六"),
    ("Ke-7", "七"),
    ("Ke-8", "八"),
    ("Ke-9", "九"),
    ("Besar/Pertama (Adik)", "大"),
]

# Aksara umum tambahan
_AKSARA_UMUM: list[tuple[str, str]] = [
    ("Laki-laki", "男"),
    ("Perempuan", "女"),
    ("Anak Lk", "子"),
    ("Ayah", "父"),
    ("Ibu", "母"),
    ("Kakek", "祖"),
    ("Nenek", "祖母"),
    ("Suami", "夫"),
    ("Istri", "妻"),
    ("Hormat (Jing Feng)", "敬奉"),
    ("Sujud (Kou Shou)", "叩首"),
    ("Almarhum (Wang)", "亡"),
    ("Saudara Kandung (Bao)", "胞"),
    ("Bakti (Xiao)", "孝"),
]


def _get_pinyin(mandarin: str) -> str:
    """Konversi aksara Mandarin ke Hanyu Pinyin dengan nada."""
    if not mandarin:
        return ""
    # Filter spasi dan karakter non-CJK
    result_parts = []
    for char in mandarin:
        if '\u4e00' <= char <= '\u9fff':
            py = pinyin(char, style=PinyinStyle.TONE)
            if py and py[0]:
                result_parts.append(py[0][0])
        elif char == ' ':
            result_parts.append(' ')
    return ' '.join(result_parts) if result_parts else mandarin


class DictionaryWindow(ctk.CTkToplevel):
    """Jendela Kamus Mini Aksara Mandarin."""

    def __init__(self, master=None):
        super().__init__(master)
        self.title("📖 Kamus Mini — Aksara Mandarin")
        self.geometry("880x680")
        self.minsize(750, 500)
        self.transient(master)
        self.grab_set()
        self.lift()
        self.focus_force()

        # Load custom entries
        self._custom_data = _load_custom_data()

        # Notification label (untuk feedback copy)
        self._notif_label = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=12),
            text_color="#4CAF50", height=20,
        )
        self._notif_label.pack(padx=10, pady=(8, 0))

        # Search bar
        search_frame = ctk.CTkFrame(self, fg_color="transparent")
        search_frame.pack(padx=15, pady=(5, 5), fill="x")
        ctk.CTkLabel(
            search_frame, text="🔍 Cari:",
            font=ctk.CTkFont(size=13),
        ).pack(side="left", padx=(0, 8))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._on_search())
        ctk.CTkEntry(
            search_frame, textvariable=self._search_var,
            width=300, placeholder_text="Ketik Indonesia / Mandarin / Pinyin...",
        ).pack(side="left", fill="x", expand=True)

        # Tabview
        self._tabview = ctk.CTkTabview(self, corner_radius=8)
        self._tabview.pack(padx=15, pady=(5, 15), fill="both", expand=True)

        self._tabview.add("Penerima (稱呼)")
        self._tabview.add("Pemberi (陽上)")
        self._tabview.add("Angka & Aksara")

        # Build tables with add/delete bar
        self._tree_panggilan = self._build_table(
            self._tabview.tab("Penerima (稱呼)"),
            self._merged_data("panggilan"),
            "panggilan",
        )
        self._build_action_bar(
            self._tabview.tab("Penerima (稱呼)"), "panggilan"
        )

        self._tree_dari = self._build_table(
            self._tabview.tab("Pemberi (陽上)"),
            self._merged_data("dari"),
            "dari",
        )
        self._build_action_bar(
            self._tabview.tab("Pemberi (陽上)"), "dari"
        )

        self._build_angka_tab(self._tabview.tab("Angka & Aksara"))

        # Store built-in counts for custom entry detection
        self._builtin_counts = {
            "panggilan": len(_PANGGILAN_DATA),
            "dari": len(_DARI_DATA),
        }

    # ============================================================
    # Data helpers
    # ============================================================
    def _merged_data(self, category: str) -> list[tuple[str, str]]:
        """Gabungkan data bawaan + kustom."""
        builtin = {
            "panggilan": _PANGGILAN_DATA,
            "dari": _DARI_DATA,
            "aksara_umum": _AKSARA_UMUM,
        }.get(category, [])
        custom = self._custom_data.get(category, [])
        return list(builtin) + list(custom)

    def _build_table(self, parent, data, tag_prefix):
        """Build a Treeview table with columns: No, Indonesia, Mandarin, Pinyin, Copy."""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=5, pady=5)

        columns = ("no", "indonesia", "mandarin", "pinyin", "copy")
        tree = ttk.Treeview(
            frame, columns=columns, show="headings",
            height=20, selectmode="browse",
        )

        tree.heading("no", text="No")
        tree.heading("indonesia", text="Bahasa Indonesia")
        tree.heading("mandarin", text="Aksara Mandarin")
        tree.heading("pinyin", text="Hanyu Pinyin")
        tree.heading("copy", text="📋 Copy")

        tree.column("no", width=40, anchor="center", stretch=False)
        tree.column("indonesia", width=250, anchor="w")
        tree.column("mandarin", width=150, anchor="center")
        tree.column("pinyin", width=200, anchor="w")
        tree.column("copy", width=70, anchor="center", stretch=False)

        # Scrollbar
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Style
        style = ttk.Style()
        style.configure(
            f"{tag_prefix}.Treeview",
            rowheight=28,
            font=("Microsoft YaHei UI", 11),
        )
        style.configure(
            f"{tag_prefix}.Treeview.Heading",
            font=("Segoe UI", 11, "bold"),
        )
        tree.configure(style=f"{tag_prefix}.Treeview")

        # Populate
        self._populate_tree(tree, data)

        # Bind click to copy
        tree.bind("<ButtonRelease-1>", lambda e: self._on_tree_click(e, tree))

        return tree

    def _build_action_bar(self, parent, category: str):
        """Build the add / delete action bar below a Treeview tab."""
        bar = ctk.CTkFrame(parent, fg_color="transparent")
        bar.pack(fill="x", padx=5, pady=(5, 0))

        # --- Input fields ---
        ctk.CTkLabel(
            bar, text="Indonesia:", font=ctk.CTkFont(size=12),
        ).pack(side="left", padx=(0, 4))
        entry_indo = ctk.CTkEntry(
            bar, width=200, placeholder_text="Nama / arti Indonesia",
        )
        entry_indo.pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            bar, text="Mandarin:", font=ctk.CTkFont(size=12),
        ).pack(side="left", padx=(0, 4))
        entry_mandarin = ctk.CTkEntry(
            bar, width=160, placeholder_text="Aksara Mandarin",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
        )
        entry_mandarin.pack(side="left", padx=(0, 8))

        # --- Tambah button ---
        btn_add = ctk.CTkButton(
            bar, text="➕ Tambah", width=100,
            fg_color="#43A047", hover_color="#2E7D32",
            command=lambda: self._add_custom_entry(
                category, entry_indo, entry_mandarin
            ),
        )
        btn_add.pack(side="left", padx=(0, 6))

        # --- Hapus button ---
        btn_del = ctk.CTkButton(
            bar, text="🗑️ Hapus", width=100,
            fg_color="#E53935", hover_color="#B71C1C",
            command=lambda: self._delete_custom_entry(category),
        )
        btn_del.pack(side="left")

    def _add_custom_entry(
        self, category: str, entry_indo: ctk.CTkEntry, entry_mandarin: ctk.CTkEntry
    ):
        """Tambah entri kustom baru ke kategori tertentu, simpan ke JSON."""
        indo = entry_indo.get().strip()
        mandarin = entry_mandarin.get().strip()

        if not indo or not mandarin:
            messagebox.showwarning(
                "Input Tidak Lengkap",
                "Kolom Indonesia dan Mandarin harus diisi.",
                parent=self,
            )
            return

        # Cek duplikat di data gabungan
        merged = self._merged_data(category)
        for existing_indo, existing_mand in merged:
            if existing_indo.lower() == indo.lower() and existing_mand == mandarin:
                messagebox.showinfo(
                    "Sudah Ada",
                    f'Entri "{indo}" ({mandarin}) sudah ada di kamus.',
                    parent=self,
                )
                return

        # Tambah & simpan
        self._custom_data.setdefault(category, []).append((indo, mandarin))
        _save_custom_data(self._custom_data)

        # Refresh tree
        self._refresh_tree(category)

        # Bersihkan input
        entry_indo.delete(0, "end")
        entry_mandarin.delete(0, "end")

        self._notif_label.configure(
            text=f"✅ Ditambahkan: {indo} ({mandarin})",
            text_color="#43A047",
        )
        self.after(2500, lambda: self._notif_label.configure(text=""))

    def _delete_custom_entry(self, category: str):
        """Hapus entri kustom yang dipilih dari Treeview."""
        tree = self._tree_panggilan if category == "panggilan" else self._tree_dari
        selected = tree.selection()
        if not selected:
            messagebox.showinfo(
                "Pilih Baris",
                "Klik baris yang ingin dihapus terlebih dahulu.",
                parent=self,
            )
            return

        item_id = selected[0]
        tags = tree.item(item_id, "tags")
        if "custom" not in tags:
            messagebox.showwarning(
                "Tidak Bisa Dihapus",
                "Hanya entri kustom (berwarna kuning) yang bisa dihapus.\n"
                "Entri bawaan tidak dapat dihapus.",
                parent=self,
            )
            return

        values = tree.item(item_id, "values")
        indo = values[1]
        mandarin = values[2]

        confirm = messagebox.askyesno(
            "Konfirmasi Hapus",
            f'Hapus entri kustom:\n"{indo}" ({mandarin})?',
            parent=self,
        )
        if not confirm:
            return

        # Hapus dari custom data
        custom_list = self._custom_data.get(category, [])
        self._custom_data[category] = [
            (i, m) for i, m in custom_list
            if not (i == indo and m == mandarin)
        ]
        _save_custom_data(self._custom_data)

        # Refresh tree
        self._refresh_tree(category)

        self._notif_label.configure(
            text=f"🗑️ Dihapus: {indo} ({mandarin})",
            text_color="#E53935",
        )
        self.after(2500, lambda: self._notif_label.configure(text=""))

    def _refresh_tree(self, category: str):
        """Refresh Treeview setelah perubahan data kustom."""
        tree = self._tree_panggilan if category == "panggilan" else self._tree_dari
        builtin_count = {
            "panggilan": len(_PANGGILAN_DATA),
            "dari": len(_DARI_DATA),
        }.get(category, 0)
        data = self._merged_data(category)
        self._populate_tree(tree, data, builtin_count=builtin_count)

    def _populate_tree(self, tree, data, builtin_count=None):
        """Populate tree with data rows. Custom entries shown in different color."""
        for item in tree.get_children():
            tree.delete(item)

        for idx, (indo, mandarin) in enumerate(data, start=1):
            py = _get_pinyin(mandarin)
            is_custom = builtin_count is not None and idx > builtin_count
            tag = "custom" if is_custom else ("even" if idx % 2 == 0 else "odd")
            tree.insert(
                "", "end",
                values=(idx, indo, mandarin, py, "📋 Copy"),
                tags=(tag,),
            )

        # Row colors
        tree.tag_configure("even", background="#F5F5F5")
        tree.tag_configure("odd", background="#FFFFFF")
        tree.tag_configure("custom", background="#FFF8E1", foreground="#E65100")

    def _build_angka_tab(self, parent):
        """Build the numbers & common characters tab."""
        container = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=5, pady=5)

        # --- Angka Dasar 1–9 ---
        self._build_card_section(
            container,
            "🔢 Angka Dasar (1–9)",
            _ANGKA_DATA,
        )

        # --- Angka Urutan ---
        self._build_card_section(
            container,
            "🔢 Angka Urutan (untuk Anak/Kakak/Adik)",
            _ANGKA_URUTAN_DATA,
        )

        # --- Aksara Umum ---
        self._build_card_section(
            container,
            "📝 Aksara Umum",
            _AKSARA_UMUM,
        )

        # --- Aksara Kustom ---
        self._aksara_custom_frame = ctk.CTkFrame(container, fg_color="transparent")
        self._aksara_custom_frame.pack(fill="x", padx=0, pady=(0, 5))
        self._rebuild_custom_aksara_section()

        # --- Input tambah aksara kustom ---
        add_frame = ctk.CTkFrame(container, fg_color="transparent")
        add_frame.pack(fill="x", padx=10, pady=(5, 10))

        ctk.CTkLabel(
            add_frame, text="Tambah Aksara Baru:",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(anchor="w", pady=(0, 5))

        input_row = ctk.CTkFrame(add_frame, fg_color="transparent")
        input_row.pack(fill="x")

        ctk.CTkLabel(input_row, text="Indonesia:", font=ctk.CTkFont(size=12)).pack(
            side="left", padx=(0, 4)
        )
        self._aksara_indo_entry = ctk.CTkEntry(
            input_row, width=180, placeholder_text="Nama / Arti",
        )
        self._aksara_indo_entry.pack(side="left", padx=(0, 8))

        ctk.CTkLabel(input_row, text="Mandarin:", font=ctk.CTkFont(size=12)).pack(
            side="left", padx=(0, 4)
        )
        self._aksara_mand_entry = ctk.CTkEntry(
            input_row, width=140, placeholder_text="Aksara",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13),
        )
        self._aksara_mand_entry.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            input_row, text="➕ Tambah", width=100,
            fg_color="#43A047", hover_color="#2E7D32",
            command=self._add_custom_aksara,
        ).pack(side="left")

    def _rebuild_custom_aksara_section(self):
        """Rebuild custom aksara card section."""
        for widget in self._aksara_custom_frame.winfo_children():
            widget.destroy()

        custom_aksara = self._custom_data.get("aksara_umum", [])
        if not custom_aksara:
            return

        ctk.CTkLabel(
            self._aksara_custom_frame,
            text="✨ Aksara Kustom (klik kanan untuk hapus)",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#E65100",
        ).pack(padx=10, pady=(12, 5), anchor="w")

        cards_frame = ctk.CTkFrame(self._aksara_custom_frame, fg_color="transparent")
        cards_frame.pack(padx=10, pady=(0, 5), fill="x")

        for idx, (label, mandarin) in enumerate(custom_aksara):
            py = _get_pinyin(mandarin)
            card = ctk.CTkFrame(
                cards_frame, corner_radius=8,
                fg_color="#FFF8E1", border_width=1, border_color="#FFB74D",
                width=120, height=90,
            )
            card.grid(row=idx // 6, column=idx % 6, padx=4, pady=4, sticky="nsew")
            card.grid_propagate(False)

            lbl_mandarin = ctk.CTkLabel(
                card, text=mandarin,
                font=ctk.CTkFont(family="Microsoft YaHei UI", size=20, weight="bold"),
                text_color="#E65100",
            )
            lbl_mandarin.pack(pady=(8, 0))

            ctk.CTkLabel(
                card, text=py,
                font=ctk.CTkFont(size=10),
                text_color="#666",
            ).pack(pady=0)

            ctk.CTkLabel(
                card, text=label,
                font=ctk.CTkFont(size=9),
                text_color="#999",
            ).pack(pady=(0, 2))

            # Left-click → copy; Right-click → delete
            for widget in (card, lbl_mandarin):
                widget.bind(
                    "<Button-1>",
                    lambda e, m=mandarin, l=label: self._copy_to_clipboard(m, l),
                )
                widget.bind(
                    "<Button-3>",
                    lambda e, m=mandarin, l=label: self._delete_custom_aksara(l, m),
                )
                widget.configure(cursor="hand2")

        for c in range(6):
            cards_frame.columnconfigure(c, weight=1)

    def _add_custom_aksara(self):
        """Tambah aksara kustom baru ke tab Angka & Aksara."""
        indo = self._aksara_indo_entry.get().strip()
        mandarin = self._aksara_mand_entry.get().strip()

        if not indo or not mandarin:
            messagebox.showwarning(
                "Input Tidak Lengkap",
                "Kolom Indonesia dan Mandarin harus diisi.",
                parent=self,
            )
            return

        # Cek duplikat
        all_aksara = _AKSARA_UMUM + self._custom_data.get("aksara_umum", [])
        for ei, em in all_aksara:
            if ei.lower() == indo.lower() and em == mandarin:
                messagebox.showinfo(
                    "Sudah Ada",
                    f'Aksara "{indo}" ({mandarin}) sudah ada di kamus.',
                    parent=self,
                )
                return

        self._custom_data.setdefault("aksara_umum", []).append((indo, mandarin))
        _save_custom_data(self._custom_data)

        self._aksara_indo_entry.delete(0, "end")
        self._aksara_mand_entry.delete(0, "end")

        self._rebuild_custom_aksara_section()

        self._notif_label.configure(
            text=f"✅ Ditambahkan: {indo} ({mandarin})",
            text_color="#43A047",
        )
        self.after(2500, lambda: self._notif_label.configure(text=""))

    def _delete_custom_aksara(self, label: str, mandarin: str):
        """Hapus aksara kustom (dipanggil via klik kanan pada card)."""
        confirm = messagebox.askyesno(
            "Konfirmasi Hapus",
            f'Hapus aksara kustom:\n"{label}" ({mandarin})?',
            parent=self,
        )
        if not confirm:
            return

        custom_list = self._custom_data.get("aksara_umum", [])
        self._custom_data["aksara_umum"] = [
            (i, m) for i, m in custom_list
            if not (i == label and m == mandarin)
        ]
        _save_custom_data(self._custom_data)
        self._rebuild_custom_aksara_section()

        self._notif_label.configure(
            text=f"🗑️ Dihapus: {label} ({mandarin})",
            text_color="#E53935",
        )
        self.after(2500, lambda: self._notif_label.configure(text=""))

    def _build_card_section(self, parent, title, data):
        """Build a section of clickable character cards."""
        ctk.CTkLabel(
            parent, text=title,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(padx=10, pady=(12, 5), anchor="w")

        cards_frame = ctk.CTkFrame(parent, fg_color="transparent")
        cards_frame.pack(padx=10, pady=(0, 5), fill="x")

        for idx, (label, mandarin) in enumerate(data):
            py = _get_pinyin(mandarin)
            card = ctk.CTkFrame(
                cards_frame, corner_radius=8,
                fg_color="#E3F2FD", border_width=1, border_color="#90CAF9",
                width=120, height=90,
            )
            card.grid(row=idx // 6, column=idx % 6, padx=4, pady=4, sticky="nsew")
            card.grid_propagate(False)

            # Aksara besar
            lbl_mandarin = ctk.CTkLabel(
                card, text=mandarin,
                font=ctk.CTkFont(family="Microsoft YaHei UI", size=20, weight="bold"),
                text_color="#1565C0",
            )
            lbl_mandarin.pack(pady=(8, 0))

            # Pinyin
            ctk.CTkLabel(
                card, text=py,
                font=ctk.CTkFont(size=10),
                text_color="#666",
            ).pack(pady=0)

            # Label Indonesia
            ctk.CTkLabel(
                card, text=label,
                font=ctk.CTkFont(size=9),
                text_color="#999",
            ).pack(pady=(0, 2))

            # Bind click → copy
            for widget in (card, lbl_mandarin):
                widget.bind(
                    "<Button-1>",
                    lambda e, m=mandarin, l=label: self._copy_to_clipboard(m, l),
                )
                widget.configure(cursor="hand2")

        # Configure grid columns
        for c in range(6):
            cards_frame.columnconfigure(c, weight=1)

    # ============================================================
    # Event handlers
    # ============================================================
    def _on_tree_click(self, event, tree):
        """Handle click on Treeview row — copy mandarin to clipboard."""
        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = tree.identify_column(event.x)
        row_id = tree.identify_row(event.y)
        if not row_id:
            return

        values = tree.item(row_id, "values")
        if not values:
            return

        # Col #3 = mandarin (index 2), Col #5 = copy button (index 4)
        # Klik di kolom mandarin atau copy → copy aksara
        if col in ("#3", "#5"):
            mandarin = values[2]
            indo = values[1]
            self._copy_to_clipboard(mandarin, indo)

    def _copy_to_clipboard(self, mandarin: str, label: str = ""):
        """Copy aksara mandarin ke clipboard dan tampilkan notifikasi."""
        self.clipboard_clear()
        self.clipboard_append(mandarin)
        self.update()

        display = f"{mandarin} ({label})" if label else mandarin
        self._notif_label.configure(
            text=f"✅ Tersalin: {display}",
            text_color="#4CAF50",
        )
        # Auto-hide after 2s
        self.after(2000, lambda: self._notif_label.configure(text=""))

    def _on_search(self):
        """Filter semua tabel berdasarkan pencarian."""
        query = self._search_var.get().strip().lower()

        # Filter Panggilan (built-in + custom)
        all_panggilan = self._merged_data("panggilan")
        filtered_p = [
            (indo, mand) for indo, mand in all_panggilan
            if query in indo.lower()
            or query in mand
            or query in _get_pinyin(mand).lower()
        ] if query else all_panggilan
        self._populate_tree(
            self._tree_panggilan, filtered_p,
            builtin_count=len(_PANGGILAN_DATA) if not query else None,
        )

        # Filter Dari (built-in + custom)
        all_dari = self._merged_data("dari")
        filtered_d = [
            (indo, mand) for indo, mand in all_dari
            if query in indo.lower()
            or query in mand
            or query in _get_pinyin(mand).lower()
        ] if query else all_dari
        self._populate_tree(
            self._tree_dari, filtered_d,
            builtin_count=len(_DARI_DATA) if not query else None,
        )
