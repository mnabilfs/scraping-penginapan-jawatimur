"""
Orkestrator Utama Pipeline Scraping & Dataset ABSA Penginapan Jawa Timur
Jalankan skrip ini untuk memilih tahap yang ingin dieksekusi:
  1. Jalankan Tahap 1 (Scraping Daftar Penginapan)
  2. Jalankan Tahap 2 (Scraping Ulasan Penginapan)
  3. Jalankan Tahap 3 (Preprocessing & Export ke Excel Lengkap)
  4. Jalankan Semua Tahap Sekaligus (1 -> 2 -> 3)
"""

import sys
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")
import subprocess

def main():
    print("""
=============================================================
   PIPELINE SCRAPING & DATASET ABSA PENGINAPAN JAWA TIMUR
=============================================================
1. Jalankan Tahap 1: Scraping Daftar Penginapan
2. Jalankan Tahap 2: Scraping Ulasan Penginapan
3. Jalankan Tahap 3: Pemrosesan ABSA & Export Excel Lengkap
4. Jalankan Semua Tahap (1 -> 2 -> 3)
5. Lengkapi Metadata Penginapan (Alamat, Telepon, Website, Harga)
6. Login Akun Google (1x Saja untuk Buka Semua Ulasan Tanpa Captcha)
0. Keluar
=============================================================
    """)
    choice = input("Pilih menu (0-6): ").strip()

    if choice == "1":
        subprocess.run([sys.executable, "01_scrape_daftar_penginapan.py"])
    elif choice == "2":
        subprocess.run([sys.executable, "02_scrape_ulasan_penginapan.py"])
    elif choice == "3":
        subprocess.run([sys.executable, "03_proses_dataset_absa.py"])
    elif choice == "4":
        print("\n--- Menjalankan Tahap 1 ---")
        subprocess.run([sys.executable, "01_scrape_daftar_penginapan.py"])
        print("\n--- Menjalankan Tahap 2 ---")
        subprocess.run([sys.executable, "02_scrape_ulasan_penginapan.py"])
        print("\n--- Menjalankan Tahap 3 ---")
        subprocess.run([sys.executable, "03_proses_dataset_absa.py"])
        print("\n🎉 Seluruh alur selesai!")
    elif choice == "5":
        subprocess.run([sys.executable, "enrich_detail_penginapan.py"])
    elif choice == "6":
        subprocess.run([sys.executable, "login_google.py"])
    elif choice == "0":
        print("Selesai.")
    else:
        print("Pilihan tidak valid.")

if __name__ == "__main__":
    main()
