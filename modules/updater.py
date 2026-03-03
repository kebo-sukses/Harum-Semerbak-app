# -*- coding: utf-8 -*-
"""
updater.py — Auto-update checker via GitHub Releases

Alur:
  1. Saat app start, cek GitHub API untuk release terbaru.
  2. Bandingkan versi cloud vs versi lokal.
  3. Jika ada versi baru → tampilkan dialog notifikasi.
  4. User klik "Download & Install" → download installer, jalankan, tutup app.
"""

import os
import sys
import json
import tempfile
import threading
import subprocess
from urllib.request import urlopen, Request
from urllib.error import URLError
from packaging.version import Version

# ============================================================
# Konfigurasi
# ============================================================
GITHUB_OWNER = "kebo-sukses"
GITHUB_REPO = "Harum-Semerbak-app"
RELEASES_API = (
    f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
)
USER_AGENT = "FormulirRitual-Updater/1.0"


# ============================================================
# Fungsi utilitas
# ============================================================

def _fetch_latest_release() -> dict | None:
    """Ambil info release terbaru dari GitHub API.

    Returns:
        dict dengan keys 'tag_name', 'name', 'body', 'assets', dsb.
        None jika gagal (offline, rate-limited, dll).
    """
    try:
        req = Request(RELEASES_API, headers={"User-Agent": USER_AGENT})
        with urlopen(req, timeout=8) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except (URLError, OSError, json.JSONDecodeError):
        return None


def _parse_version(tag: str) -> Version:
    """Parse tag seperti 'v1.0.1' atau '1.0.1' menjadi Version object."""
    tag = tag.strip().lstrip("vV")
    return Version(tag)


def _find_installer_asset(assets: list[dict]) -> dict | None:
    """Cari asset .exe installer dari daftar release assets."""
    for asset in assets:
        name = asset.get("name", "").lower()
        if name.endswith(".exe") and "setup" in name:
            return asset
    # Fallback: cari file .exe apapun
    for asset in assets:
        name = asset.get("name", "").lower()
        if name.endswith(".exe"):
            return asset
    return None


def _download_file(url: str, dest_path: str,
                   progress_callback=None) -> bool:
    """Download file dari URL ke dest_path.

    Args:
        url: URL download (browser_download_url dari GitHub asset).
        dest_path: Path tujuan file.
        progress_callback: Optional callable(downloaded_bytes, total_bytes).

    Returns:
        True jika berhasil, False jika gagal.
    """
    try:
        req = Request(url, headers={"User-Agent": USER_AGENT})
        with urlopen(req, timeout=60) as resp:
            total = int(resp.headers.get("Content-Length", 0))
            downloaded = 0
            chunk_size = 64 * 1024  # 64 KB

            with open(dest_path, "wb") as f:
                while True:
                    chunk = resp.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback:
                        progress_callback(downloaded, total)
        return True
    except (URLError, OSError):
        return False


# ============================================================
# Kelas utama
# ============================================================

class UpdateChecker:
    """Mengecek dan menjalankan update dari GitHub Releases.

    Penggunaan:
        checker = UpdateChecker(current_version="1.0.0")
        checker.check_in_background(on_update_available=callback)
    """

    def __init__(self, current_version: str):
        self.current_version = _parse_version(current_version)
        self.latest_release: dict | None = None
        self.latest_version: Version | None = None
        self.installer_asset: dict | None = None

    # --------------------------------------------------------
    # Background check
    # --------------------------------------------------------
    def check_in_background(self, on_update_available=None,
                            on_no_update=None):
        """Cek update di background thread, panggil callback jika ada.

        Args:
            on_update_available: callable(version_str, release_name, release_notes)
                Dipanggil di *main thread* jika ada versi baru.
            on_no_update: callable() — jika sudah terbaru / gagal cek.
        """
        def _worker():
            has_update = self._check()
            if has_update and on_update_available:
                on_update_available(
                    str(self.latest_version),
                    self.latest_release.get("name", ""),
                    self.latest_release.get("body", ""),
                )
            elif on_no_update:
                on_no_update()

        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def _check(self) -> bool:
        """Cek apakah ada versi baru. Return True jika ada."""
        release = _fetch_latest_release()
        if not release:
            return False

        tag = release.get("tag_name", "")
        if not tag:
            return False

        try:
            remote_ver = _parse_version(tag)
        except Exception:
            return False

        self.latest_release = release
        self.latest_version = remote_ver
        self.installer_asset = _find_installer_asset(
            release.get("assets", [])
        )

        return remote_ver > self.current_version

    # --------------------------------------------------------
    # Download & Install
    # --------------------------------------------------------
    def download_and_install(self, progress_callback=None,
                             on_done=None, on_error=None):
        """Download installer lalu jalankan. Harus dipanggil setelah check.

        Args:
            progress_callback: callable(downloaded, total) — progress bar.
            on_done: callable(installer_path) — setelah download selesai.
            on_error: callable(error_msg) — jika gagal.
        """
        if not self.installer_asset:
            if on_error:
                on_error("Tidak ditemukan file installer di release terbaru.")
            return

        def _worker():
            url = self.installer_asset["browser_download_url"]
            filename = self.installer_asset["name"]
            dest = os.path.join(tempfile.gettempdir(), filename)

            ok = _download_file(url, dest, progress_callback)
            if ok and os.path.exists(dest):
                if on_done:
                    on_done(dest)
            else:
                if on_error:
                    on_error("Gagal mendownload installer.")

        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    @staticmethod
    def launch_installer_and_exit(installer_path: str):
        """Jalankan installer dan tutup aplikasi saat ini."""
        try:
            subprocess.Popen(
                [installer_path],
                creationflags=subprocess.DETACHED_PROCESS
                if sys.platform == "win32" else 0,
            )
        except OSError:
            pass
        sys.exit(0)
