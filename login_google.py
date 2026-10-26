"""
Script Bantu Login Google Tanpa Terdeteksi Robot / Loop Captcha
Membuka Google Chrome asli dengan profil 'data/browser_profile'.
Karena dibuka sebagai Chrome murni tanpa otomasi Playwright, Google tidak akan memblokir login Anda.
"""

import os
import subprocess
import time
from pathlib import Path

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
PROFILE_DIR = DATA_DIR / "browser_profile"
PROFILE_DIR.mkdir(parents=True, exist_ok=True)

CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe")
]

chrome_exe = None
for p in CHROME_PATHS:
    if os.path.exists(p):
        chrome_exe = p
        break

if not chrome_exe:
    print("❌ Google Chrome tidak ditemukan di komputer Anda!")
    exit(1)

print("=" * 65)
print("🔑 MEMBUKA GOOGLE CHROME ASLI UNTUK LOGIN GOOGLE")
print("=" * 65)
print(f"📁 Profil disimpan di: {PROFILE_DIR}")
print("🌐 Membuka halaman login Google...")
print("=" * 65)

cmd = [
    chrome_exe,
    f"--user-data-dir={PROFILE_DIR.resolve()}",
    "--no-first-run",
    "--no-default-browser-check",
    "https://accounts.google.com/signin"
]

proc = subprocess.Popen(cmd)

print("\nSilakan login ke akun Google Anda di jendela Google Chrome yang terbuka.")
print("Setelah berhasil login (masuk ke akun Google):")
input("👉 Tekan [ENTER] di terminal ini untuk menutup dan menyimpan sesi login...")

try:
    proc.terminate()
except Exception:
    pass

# Buat flag login selesai
flag_file = PROFILE_DIR / "login_completed.flag"
flag_file.write_text("logged_in", encoding="utf-8")

print("\n" + "=" * 65)
print("🎉 LOGIN BERHASIL DISIMPAN!")
print("Sekarang Anda bisa menjalankan 'python 02_scrape_ulasan_penginapan.py'.")
print("=" * 65)
