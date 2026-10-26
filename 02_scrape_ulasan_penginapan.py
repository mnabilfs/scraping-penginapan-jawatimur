"""
TAHAP 2: Scraping Ulasan Penginapan Jawa Timur dari Google Maps (Versi Full Reviews & Tanggapan Pemilik)
Output: data/raw_ulasan_checkpoint.csv

Fitur Unggulan:
  1. Persistent Browser Profile: Sesi login Google Anda disimpan secara permanen di `data/browser_profile/`.
  2. Ekstraksi Balasan Pemilik: Mengambil teks tanggapan resmi dari pemilik hotel beserta tanggalnya.
  3. Full Infinite Scroll: Menggunakan selektor 'div.jftiEf' murni sehingga kuota ulasan puluhan hingga ratusan ulasan per hotel terambil penuh.
  4. Real-time Hotel Progress Bar: Menampilkan informasi hotel & progress live sesuai format:
     [Hotel] Hotel Majapahit Surabaya MGallery (2.500 ulasan)
     Scraping Review:  50% |█████     | 50/100 ulasan
  5. In-place Checkpoint Enrichment: Memperkaya ulasan lama yang belum ada teks balasannya, dan menambah ulasan baru.
  6. Parameter Fleksibel: Dukungan CLI (--max 100 / --max 500 / --hotel "Cicada").
"""

import sys
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import os
import re
import time
import hashlib
import csv
import argparse
import subprocess
from pathlib import Path
import pandas as pd
from tqdm import tqdm
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
INPUT_HOTEL_FILE = DATA_DIR / "daftar_penginapan_jatim_mentah.csv"
OUTPUT_REVIEW_FILE = DATA_DIR / "raw_ulasan_checkpoint.csv"
USER_DATA_DIR = DATA_DIR / "browser_profile"
LOGIN_FLAG_FILE = USER_DATA_DIR / "login_completed.flag"

# Kolom standar dataset ulasan (termasuk fasilitas, rating_penginapan, dan balasan_pemilik)
CSV_COLUMNS = [
    "review_id", "data_id", "nama_penginapan", "kabupaten", "kelompok_sampel",
    "kata_kunci", "fasilitas", "teks_ulasan", "teks_terjemahan_google", "rating_penginapan", "rating_bintang",
    "tanggal", "jumlah_like", "ada_balasan_pemilik", "balasan_pemilik", "sumber", "reviewer_total_ulasan"
]

def hash_review_id(text: str, author: str, date_str: str) -> str:
    seed = f"{author}_{date_str}_{text}".encode("utf-8")
    return hashlib.sha256(seed).hexdigest()[:24]

def clean_csv_cell(val) -> str:
    """
    Membersihkan karakter enter/newline (\\r, \\n), tab, dan spasi berlebih
    agar setiap teks ulasan, tanggal, dan balasan pemilik tersimpan tepat dalam 1 baris CSV.
    Hal ini mencegah baris bergeser atau rusak saat file CSV dibuka atau diproses di Excel.
    """
    if val is None or pd.isna(val):
        return ""
    txt = str(val)
    txt = re.sub(r"[\r\n]+", " ", txt)
    txt = txt.replace("\t", " ")
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt

# Script JS untuk ekstraksi fasilitas penginapan (bagian 'Detail hotel')
JS_EXTRACT_FACILITIES = r"""
() => {
    // 1. Scroll container sedikit ke bawah untuk memuat bagian 'Detail hotel'
    const container = document.querySelector('div.m6QErb.DxyBCb.kA9KIf.dS8AEf') || document.querySelector('.m6QErb');
    if (container) container.scrollTop = 1200;

    // 2. Klik tombol ekspansi fasilitas ("Tampilkan lebih banyak fasilitas") jika ada
    document.querySelectorAll('button').forEach(btn => {
        const txt = (btn.innerText || '').toLowerCase();
        const aria = (btn.getAttribute('aria-label') || '').toLowerCase();
        if (txt.includes('fasilitas') || aria.includes('fasilitas')) {
            try { btn.click(); } catch(e){}
        }
    });

    // 3. Ambil teks fasilitas dari span / div berkelas gSamH
    let items = Array.from(document.querySelectorAll('span.gSamH, div.gSamH'))
                     .map(s => s.innerText.trim())
                     .filter(Boolean);

    // Fallback jika class gSamH berbeda: cari di sekitar teks "Detail hotel"
    if (items.length === 0) {
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
        let node;
        let detailEl = null;
        while (node = walker.nextNode()) {
            if (/detail hotel/i.test(node.nodeValue.trim())) {
                detailEl = node.parentElement;
                break;
            }
        }
        if (detailEl) {
            const p = detailEl.parentElement;
            if (p) {
                const spans = Array.from(p.querySelectorAll('span.fontBodySmall, div.fontBodySmall'));
                items = spans.map(s => s.innerText.trim())
                             .filter(t => t && !/detail hotel|tampilkan|ulasan/i.test(t));
            }
        }
    }

    return Array.from(new Set(items));
}
"""

# Script JS untuk ekspansi semua teks terpotong ('Lainnya')
JS_EXPAND_ALL = r"""
() => {
    document.querySelectorAll('button').forEach(btn => {
        const txt = (btn.textContent || '').trim();
        const aria = (btn.getAttribute('aria-label') || '').trim();
        if (txt === 'Lainnya' || txt === 'More' || aria.includes('ulasan lengkap') || aria.includes('Lihat lainnya') || btn.classList.contains('w8nwRe')) {
            try { btn.click(); } catch(e){}
        }
    });
}
"""

# Script JS untuk ekstraksi ulasan lengkap + balasan pemilik
JS_EXTRACT_REVIEWS = r"""
() => {
    const cards = Array.from(document.querySelectorAll('div.jftiEf'));
    const results = [];

    cards.forEach(card => {
        try {
            const rawId = card.getAttribute('data-review-id') || '';
            
            // 1. Nama reviewer
            const authorEl = card.querySelector('div.d4r55');
            const author = authorEl ? authorEl.innerText.trim() : 'Pengguna Google';

            // 2. Info reviewer (misal: "Local Guide · 15 ulasan")
            const infoEl = card.querySelector('div.RfnDt');
            const infoText = infoEl ? infoEl.innerText.trim() : '';

            // 3. Bintang ulasan (aria-label atau format teks '5/5')
            let starText = '';
            const starEl = card.querySelector('span.kvMYJc');
            if (starEl) {
                starText = starEl.getAttribute('aria-label') || '';
            }
            if (!starText) {
                const matchSpan = Array.from(card.querySelectorAll('span, div')).find(s => /^[1-5]\/5$/.test((s.innerText || '').trim()));
                if (matchSpan) starText = matchSpan.innerText.trim();
            }

            // 4. Tanggal ulasan
            let date = '';
            const dateEl = card.querySelector('span.rsqaWe, span.DU9Pgb, span.xRkAfb');
            if (dateEl && dateEl.innerText.trim()) {
                date = dateEl.innerText.trim();
            } else {
                for (let s of card.querySelectorAll('span')) {
                    const txt = s.innerText.trim();
                    if (/(lalu|ago|diedit|edited|kemarin|yesterday|\d+\s+(hari|minggu|bulan|tahun))/i.test(txt) && txt.length < 35) {
                        date = txt;
                        break;
                    }
                }
            }

            // 5. Teks ulasan reviewer murni (bukan dari blok balasan pemilik)
            let reviewText = '';
            const allWi = Array.from(card.querySelectorAll('.wiI7pd'));
            const revWi = allWi.find(w => !w.closest('div.CDe7pd'));
            if (revWi) {
                reviewText = revWi.innerText.trim();
            }

            // 6. Likes (Tombol Suka / Upvote)
            let likesCount = 0;
            const likeBtn = card.querySelector('button[jsaction*="toggleThumbsUp"]') ||
                            card.querySelector('button[aria-label*="suka" i]') ||
                            card.querySelector('button[aria-label*="like" i]') ||
                            card.querySelector('div.GBkF3d') ||
                            card.querySelector('button.gllhef');
            if (likeBtn) {
                // Prioritas 1: span.pkWtMe spesifik (angka like resmi Google Maps)
                const pkSpan = likeBtn.querySelector('span.pkWtMe') || card.querySelector('span.pkWtMe');
                if (pkSpan && pkSpan.innerText.trim()) {
                    const m = pkSpan.innerText.trim().match(/(\d+)/);
                    if (m) likesCount = parseInt(m[1], 10);
                }
                // Prioritas 2: aria-label / title (contoh: "1 suka", "5 likes")
                if (likesCount === 0) {
                    const aria = ((likeBtn.getAttribute('aria-label') || '') + ' ' + (likeBtn.getAttribute('title') || '')).trim();
                    const m = aria.match(/(\d+)/);
                    if (m) likesCount = parseInt(m[1], 10);
                }
                // Prioritas 3: seluruh innerText pada likeBtn
                if (likesCount === 0) {
                    const txt = likeBtn.innerText || '';
                    const m = txt.match(/(\d+)/);
                    if (m) likesCount = parseInt(m[1], 10);
                }
            }
            const likes = likesCount;

            // 7. Balasan Pemilik (Owner Response)
            const replyBlock = card.querySelector('div.CDe7pd');
            let hasReply = false;
            let replyText = '';
            let replyDate = '';

            if (replyBlock) {
                hasReply = true;
                const replyBody = replyBlock.querySelector('.wiI7pd') || replyBlock;
                replyText = replyBody.innerText.trim();
                replyText = replyText.replace(/^Tanggapan dari pemilik[^\n]*\n?/i, '')
                                     .replace(/Lihat terjemahan.*$/i, '')
                                     .trim();
                const replyDateEl = replyBlock.querySelector('.DZSIDd, .DHIhCt, .rsqaWe');
                replyDate = replyDateEl ? replyDateEl.innerText.trim() : '';
            }

            results.push({
                raw_id: rawId,
                author: author,
                info: infoText,
                star: starText,
                date: date,
                text: reviewText,
                likes: likes,
                has_reply: hasReply,
                reply_text: replyText,
                reply_date: replyDate
            });
        } catch(err) {}
    });

    return results;
}
"""

def migrate_checkpoint_schema():
    """Pastikan checkpoint lama memiliki kolom 'fasilitas', 'balasan_pemilik', 'rating_penginapan' serta format rating desimal."""
    if not OUTPUT_REVIEW_FILE.exists():
        return
    try:
        df = pd.read_csv(OUTPUT_REVIEW_FILE)
        changed = False

        if "rating_penginapan" not in df.columns:
            df["rating_penginapan"] = ""
            changed = True

        # Isi rating_penginapan dari data hotel jika masih kosong
        if INPUT_HOTEL_FILE.exists():
            df_hotels = pd.read_csv(INPUT_HOTEL_FILE)
            hotel_map = df_hotels.drop_duplicates(subset=["data_id"]).set_index("data_id")["rating_rata2"].to_dict()
            df["rating_penginapan"] = df["data_id"].astype(str).map(hotel_map)
            changed = True

        # Format rating_bintang ke desimal (contoh: 5.0, 4.0)
        if "rating_bintang" in df.columns:
            def to_decimal_str(x):
                try:
                    return f"{float(x):.1f}"
                except Exception:
                    return ""
            df["rating_bintang"] = df["rating_bintang"].apply(to_decimal_str)
            changed = True

        if "fasilitas" not in df.columns:
            df["fasilitas"] = ""
            changed = True
        if "balasan_pemilik" not in df.columns:
            df["balasan_pemilik"] = ""
            changed = True
        for col in CSV_COLUMNS:
            if col not in df.columns:
                df[col] = ""
                changed = True
            elif df[col].dtype == object:
                df[col] = df[col].apply(clean_csv_cell)
                changed = True
        if changed:
            df = df[CSV_COLUMNS]
            df.to_csv(OUTPUT_REVIEW_FILE, index=False, encoding="utf-8")
            print("✅ Skema checkpoint berhasil diperbarui dengan kolom 'rating_penginapan' & format rating berdesimal.")
    except Exception as e:
        print(f"⚠️ Catatan migrasi schema: {e}")

def ensure_google_login():
    """Pastikan user sudah login via login_google.py agar tidak terdeteksi robot."""
    if not LOGIN_FLAG_FILE.exists():
        print("\n" + "=" * 65)
        print("🔑 STATUS: BELUM LOGIN GOOGLE")
        print("=" * 65)
        print("Untuk menghindari loop captcha 'bukan robot' dari Google,")
        print("login akan dibuka melalui Google Chrome asli Anda.")
        print("Membuka login_google.py sekarang...")
        print("=" * 65)
        subprocess.run([sys.executable, "login_google.py"])
    else:
        print("✅ Sesi login Google aktif ditemukan di folder profil.")

def scrape_ulasan_penginapan(
    headless: bool = False,
    max_ulasan_per_hotel: int = 100,
    hotel_filter: str = "",
    start_index: int = None,
    retry_incomplete: bool = False
):
    print("=" * 65)
    print("🚀 MEMULAI TAHAP 2: SCRAPING ULASAN PENGINAPAN (FULL REVIEWS)")
    if max_ulasan_per_hotel and max_ulasan_per_hotel > 0:
        print(f"🎯 Target kuota ulasan per hotel: hingga {max_ulasan_per_hotel} ulasan")
    else:
        print("🎯 Target kuota ulasan per hotel: SEMUA ulasan yang tersedia di Google Maps")
    if hotel_filter:
        print(f"🔍 Filter hotel aktif: '{hotel_filter}'")
    print(f"💾 Lokasi profil browser: {USER_DATA_DIR}")
    print("=" * 65)

    if not INPUT_HOTEL_FILE.exists():
        print(f"❌ File '{INPUT_HOTEL_FILE}' belum ada. Jalankan 01_scrape_daftar_penginapan.py terlebih dahulu!")
        return

    # Inisialisasi folder profil browser
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

    # Pastikan user sudah login
    ensure_google_login()

    # Migrasi schema CSV jika diperlukan
    migrate_checkpoint_schema()

    # Muat database ulasan yang sudah tersimpan
    all_reviews_dict = {}
    hotel_review_counts = {}
    hotels_with_facilities = set()
    last_hotel_data_id = None
    last_hotel_name = None

    if OUTPUT_REVIEW_FILE.exists():
        try:
            df_old = pd.read_csv(OUTPUT_REVIEW_FILE)
            for r in df_old.to_dict("records"):
                rid = str(r.get("review_id", "")).strip()
                if rid:
                    for k in r:
                        if isinstance(r[k], str):
                            r[k] = clean_csv_cell(r[k])
                    all_reviews_dict[rid] = r
            hotel_review_counts = df_old["data_id"].dropna().astype(str).value_counts().to_dict()
            if "fasilitas" in df_old.columns:
                has_fac_df = df_old[df_old["fasilitas"].notna() & (df_old["fasilitas"].astype(str).str.strip() != "")]
                hotels_with_facilities = set(has_fac_df["data_id"].dropna().astype(str).unique())

            valid_rows = df_old[df_old["data_id"].notna() & (df_old["data_id"].astype(str).str.strip() != "")]
            if len(valid_rows) > 0:
                last_hotel_data_id = str(valid_rows.iloc[-1]["data_id"]).strip()
                last_hotel_name = str(valid_rows.iloc[-1]["nama_penginapan"]).strip()

            print(f"📦 Checkpoint tersimpan: {len(all_reviews_dict)} ulasan dari {len(hotel_review_counts)} hotel ({len(hotels_with_facilities)} sudah berfasilitas).")
        except Exception as e:
            print(f"⚠️ Gagal membaca checkpoint ({e}). Memulai fresh database.")

    # Jika file belum ada, buat dengan header lengkap
    if not OUTPUT_REVIEW_FILE.exists():
        with open(OUTPUT_REVIEW_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
            writer.writeheader()

    # Muat daftar hotel master
    df_hotels = pd.read_csv(INPUT_HOTEL_FILE)
    total_master_hotels = len(df_hotels)
    print(f"📋 Total penginapan terdaftar: {total_master_hotels}")

    # Filter jika user ingin scrape hotel tertentu saja
    if hotel_filter:
        df_filtered = df_hotels[df_hotels["nama_penginapan"].str.contains(hotel_filter, case=False, na=False)].copy()
        print(f"🔍 Ditemukan {len(df_filtered)} hotel yang cocok dengan filter '{hotel_filter}'.")
        hotels_to_process = [(idx, row) for idx, row in df_filtered.iterrows()]
    elif retry_incomplete:
        print("🔄 Mode '--retry-incomplete' aktif: Memeriksa kembali hotel lama yang ulasannya kurang dari target kuota.")
        hotels_to_process = []
        for idx, row in df_hotels.iterrows():
            d_id = str(row.get("data_id", ""))
            current_rev_count = hotel_review_counts.get(d_id, 0)
            total_est = row.get("jumlah_ulasan_total", 0)
            if pd.notnull(total_est) and total_est == 0:
                continue
            if max_ulasan_per_hotel and max_ulasan_per_hotel > 0:
                target_quota = min(int(total_est), max_ulasan_per_hotel) if (pd.notnull(total_est) and total_est > 0) else max_ulasan_per_hotel
            else:
                target_quota = int(total_est) if (pd.notnull(total_est) and total_est > 0) else 100
            has_fac = d_id in hotels_with_facilities
            if current_rev_count >= target_quota and has_fac:
                continue
            hotels_to_process.append((idx, row))
    else:
        # Default: Lanjutkan secara sekuensial dari posisi hotel setelah checkpoint terakhir
        start_from_idx = 0
        if start_index is not None and start_index > 0:
            start_from_idx = max(0, min(start_index - 1, total_master_hotels - 1))
            target_name = df_hotels.iloc[start_from_idx]["nama_penginapan"]
            print(f"👉 Manual start diaktifkan: Memulai langsung dari [Hotel {start_from_idx + 1}/{total_master_hotels}] {target_name}")
        elif last_hotel_data_id:
            matches = df_hotels.index[df_hotels["data_id"].astype(str) == last_hotel_data_id].tolist()
            if matches:
                last_idx = matches[-1]
                start_from_idx = last_idx + 1
                print(f"📦 Checkpoint terakhir di CSV: [Hotel {last_idx + 1}/{total_master_hotels}] {last_hotel_name}")
                if start_from_idx < total_master_hotels:
                    next_hotel_name = df_hotels.iloc[start_from_idx]["nama_penginapan"]
                    print(f"⏩ Melanjutkan otomatis dari hotel berikutnya: [Hotel {start_from_idx + 1}/{total_master_hotels}] {next_hotel_name}")
                else:
                    print(f"🎉 Semua {total_master_hotels} penginapan sudah selesai diproses sampai nomor terakhir!")
                    return
            else:
                print("ℹ️ Hotel terakhir tidak ditemukan di daftar master. Memulai dari awal...")

        hotels_to_process = []
        for idx in range(start_from_idx, total_master_hotels):
            hotels_to_process.append((idx, df_hotels.iloc[idx]))

    print(f"🎯 Sisa penginapan yang akan diproses: {len(hotels_to_process)} hotel")

    if len(hotels_to_process) == 0:
        print("🎉 Semua penginapan sudah selesai diproses!")
        return

    with sync_playwright() as p:
        # Buka Chrome asli dengan persistent profile
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            channel="chrome",
            headless=headless,
            args=[
                "--lang=id-ID,id",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
                "--enable-webgl",
                "--use-gl=angle",
                "--use-angle=default",
                "--ignore-gpu-blocklist",
                "--no-sandbox"
            ],
            viewport={"width": 1400, "height": 900},
            locale="id-ID"
        )

        page = context.pages[0] if context.pages else context.new_page()

        # Blokir file gambar & media agar pergerakan super cepat dan hemat bandwidth
        def intercept_heavy(route):
            if route.request.resource_type in ["image", "media"]:
                route.abort()
            else:
                route.continue_()

        page.route("**/*", intercept_heavy)

        total_new_reviews = 0

        for h_counter, (global_idx, hotel) in enumerate(hotels_to_process, start=1):
            data_id = str(hotel.get("data_id", ""))
            nama = str(hotel.get("nama_penginapan", ""))
            kabupaten = str(hotel.get("kabupaten", ""))
            kelompok = str(hotel.get("kelompok_sampel", "umum"))
            kata_kunci = str(hotel.get("query_sumber", "_umum"))
            url = hotel.get("url", "")
            raw_total_est = hotel.get("jumlah_ulasan_total", 0)
            total_est = int(raw_total_est) if (pd.notnull(raw_total_est) and raw_total_est > 0) else 0
            hotel_rating_raw = hotel.get("rating_rata2")
            try:
                rating_penginapan = round(float(hotel_rating_raw), 1) if pd.notnull(hotel_rating_raw) else ""
            except Exception:
                rating_penginapan = ""

            # Hitung target spesifik hotel ini
            if max_ulasan_per_hotel and max_ulasan_per_hotel > 0:
                target_quota = min(total_est, max_ulasan_per_hotel) if total_est > 0 else max_ulasan_per_hotel
            else:
                target_quota = total_est if total_est > 0 else 100

            hotel_display_num = global_idx + 1
            print(f"\n🏨 [Hotel {hotel_display_num}/{total_master_hotels}] {nama} ({total_est} ulasan)")

            if not url or pd.isnull(url):
                print("   ⚠️ URL tidak valid. Melewati hotel ini.")
                continue

            if total_est == 0:
                print("   ℹ️ Hotel tercatat memiliki 0 ulasan di Google Maps. Melewati...")
                continue

            current_rev_count = hotel_review_counts.get(data_id, 0)
            has_fac = data_id in hotels_with_facilities

            try:
                page.goto(url, timeout=30000, wait_until="domcontentloaded")
                time.sleep(2.0)

                # Deteksi jika ada verifikasi captcha robot
                if "Akses dibatasi" in page.content() or "unusual traffic" in page.content():
                    print(f"\n⚠️ Verifikasi Captcha terdeteksi pada hotel: {nama}")
                    print("   Silakan selesaikan verifikasi di layar Chrome...")
                    while "Akses dibatasi" in page.content() or "unusual traffic" in page.content():
                        time.sleep(3)
                    print("   ✅ Akses terbuka kembali! Melanjutkan...")

                # 1. Ekstraksi Fasilitas Penginapan (sebelum beralih ke Tab Ulasan)
                fac_list = []
                fasilitas_str = ""
                try:
                    fac_list = page.evaluate(JS_EXTRACT_FACILITIES)
                    if fac_list:
                        fasilitas_str = ", ".join(fac_list)
                except Exception:
                    pass

                # Ekstraksi fallback rating_penginapan langsung dari Ringkasan Google Maps jika belum ada
                if not rating_penginapan:
                    try:
                        live_rat = page.evaluate("""() => {
                            const el = document.querySelector('div.jANrlb .fontDisplayLarge, span.ceNzKf, div.F7nice span[aria-hidden="true"], div.fontDisplayLarge');
                            if (el) {
                                const m = (el.innerText || '').match(/(\\d+[.,]\\d+)/);
                                if (m) return m[1].replace(',', '.');
                            }
                            return '';
                        }""")
                        if live_rat:
                            rating_penginapan = round(float(live_rat), 1)
                    except Exception:
                        pass

                # Sinkronkan fasilitas dan rating_penginapan untuk semua ulasan hotel ini di database jika ada
                if fasilitas_str or rating_penginapan:
                    for r_item in all_reviews_dict.values():
                        if str(r_item.get("data_id")) == data_id:
                            if fasilitas_str:
                                r_item["fasilitas"] = fasilitas_str
                            if rating_penginapan:
                                r_item["rating_penginapan"] = rating_penginapan

                # Jika kuota ulasan hotel ini sudah tercapai sebelumnya, cukup perbarui fasilitas lalu simpan
                if retry_incomplete and current_rev_count >= target_quota and has_fac:
                    fac_info = f" ({len(fac_list)} fasilitas terdata)" if fac_list else ""
                    print(f"   ✅ Fasilitas diperbarui{fac_info}. Kuota ulasan sudah lengkap ({current_rev_count} ulasan).")
                    continue

                # 2. Buka Tab Ulasan
                tab_clicked = page.evaluate("""() => {
                    const tabs = Array.from(document.querySelectorAll('[role="tab"]'));
                    const ulasanTab = tabs.find(t => {
                        const txt = (t.innerText || '').toLowerCase();
                        const aria = (t.getAttribute('aria-label') || '').toLowerCase();
                        return txt.includes('ulasan') || aria.includes('ulasan');
                    });
                    if (ulasanTab) {
                        ulasanTab.click();
                        return true;
                    }
                    return false;
                }""")

                if not tab_clicked:
                    for sel in ["button[role='tab']:has-text('Ulasan')", "button[aria-label*='Ulasan']", "button:has-text('Ulasan')"]:
                        try:
                            loc = page.locator(sel).first
                            if loc.count() > 0:
                                loc.click(timeout=1500)
                                tab_clicked = True
                                break
                        except Exception:
                            pass

                # Tunggu hingga kartu ulasan pertama muncul
                try:
                    page.wait_for_selector("div.jftiEf", timeout=10000)
                except Exception:
                    pass
                time.sleep(1.0)

                # Inisialisasi progress bar live ulasan untuk hotel ini
                curr_card_count = page.locator("div.jftiEf").count()
                pbar = tqdm(
                    total=target_quota,
                    desc="Scraping Review",
                    unit="ulasan",
                    bar_format="{desc}: {percentage:3.0f}% |{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]",
                    leave=True
                )
                if curr_card_count > 0:
                    pbar.update(min(curr_card_count, target_quota))

                max_scroll_attempts = max(20, (target_quota // 5) + 15)
                same_count_streak = 0
                last_count = curr_card_count

                for scroll_step in range(max_scroll_attempts):
                    if curr_card_count >= target_quota:
                        break

                    page.evaluate("""() => {
                        const container = document.querySelector('div.m6QErb.DxyBCb.kA9KIf.dS8AEf') || 
                                          document.querySelector('div[role="main"] .m6QErb') ||
                                          document.querySelector('.m6QErb');
                        if (container) {
                            container.scrollTop = container.scrollHeight;
                        }
                    }""")
                    time.sleep(0.8)

                    curr_card_count = page.locator("div.jftiEf").count()
                    diff = min(curr_card_count, target_quota) - pbar.n
                    if diff > 0:
                        pbar.update(diff)

                    if curr_card_count == 0:
                        # Coba klik ulang tab ulasan jika belum terbuka
                        page.evaluate("""() => {
                            const tabs = Array.from(document.querySelectorAll('[role="tab"]'));
                            const ulasanTab = tabs.find(t => (t.innerText || '').toLowerCase().includes('ulasan'));
                            if (ulasanTab) ulasanTab.click();
                        }""")
                        time.sleep(1.5)
                    elif curr_card_count == last_count:
                        same_count_streak += 1
                        if same_count_streak >= 5:
                            # Sudah mentok tidak ada ulasan lagi
                            break
                    else:
                        same_count_streak = 0
                        last_count = curr_card_count

                pbar.close()

                # 3. Ekspansi semua teks ulasan dan balasan pemilik ('Lainnya')
                page.evaluate(JS_EXPAND_ALL)
                time.sleep(0.8)

                # 4. Ekstraksi kartu ulasan + balasan pemilik
                raw_items = page.evaluate(JS_EXTRACT_REVIEWS)

                hotel_added = 0
                hotel_replies = 0
                hotel_new_records = []

                for item in raw_items[:target_quota]:
                    teks_ulasan = clean_csv_cell(item["text"] or nama)
                    tanggal_str = clean_csv_cell(item["date"])
                    author = clean_csv_cell(item["author"])
                    has_reply = bool(item["has_reply"])
                    reply_text = clean_csv_cell(item["reply_text"]) if has_reply else ""

                    if has_reply:
                        hotel_replies += 1

                    # Rating Bintang Reviewer (format desimal 1.0 - 5.0)
                    star_match = re.search(r"(\d+(?:[.,]\d+)?)", item["star"])
                    if star_match:
                        try:
                            val_f = float(star_match.group(1).replace(",", "."))
                            rating_bintang = f"{val_f:.1f}"
                        except Exception:
                            rating_bintang = "5.0"
                    else:
                        rating_bintang = "5.0"

                    # Total Ulasan Reviewer
                    rev_match = re.search(r"(\d+)", item["info"].replace(".", ""))
                    reviewer_total = int(rev_match.group(1)) if rev_match else 1

                    # Likes
                    like_raw = item.get("likes", 0)
                    if isinstance(like_raw, (int, float)):
                        likes_count = int(like_raw)
                    else:
                        like_m = re.search(r"(\d+)", str(like_raw).replace(".", "").replace(",", ""))
                        likes_count = int(like_m.group(1)) if like_m else 0

                    # ID Ulasan Unik
                    rev_id = item["raw_id"] if item["raw_id"] else hash_review_id(teks_ulasan, author, tanggal_str)

                    record = {
                        "review_id": rev_id,
                        "data_id": data_id,
                        "nama_penginapan": nama,
                        "kabupaten": kabupaten,
                        "kelompok_sampel": kelompok,
                        "kata_kunci": kata_kunci,
                        "fasilitas": fasilitas_str,
                        "teks_ulasan": teks_ulasan,
                        "teks_terjemahan_google": "",
                        "rating_penginapan": rating_penginapan,
                        "rating_bintang": rating_bintang,
                        "tanggal": tanggal_str,
                        "jumlah_like": likes_count,
                        "ada_balasan_pemilik": has_reply,
                        "balasan_pemilik": reply_text,
                        "sumber": "Google Maps",
                        "reviewer_total_ulasan": reviewer_total
                    }

                    # Deduplikasi cerdas & pengayaan ulasan lama
                    if rev_id in all_reviews_dict:
                        if rating_penginapan and not all_reviews_dict[rev_id].get("rating_penginapan"):
                            all_reviews_dict[rev_id]["rating_penginapan"] = rating_penginapan
                        if fasilitas_str and not all_reviews_dict[rev_id].get("fasilitas"):
                            all_reviews_dict[rev_id]["fasilitas"] = fasilitas_str
                        if has_reply and not all_reviews_dict[rev_id].get("balasan_pemilik"):
                            all_reviews_dict[rev_id]["balasan_pemilik"] = reply_text
                            all_reviews_dict[rev_id]["ada_balasan_pemilik"] = True
                        if len(teks_ulasan) > len(str(all_reviews_dict[rev_id].get("teks_ulasan", ""))):
                            all_reviews_dict[rev_id]["teks_ulasan"] = teks_ulasan
                        if likes_count > int(all_reviews_dict[rev_id].get("jumlah_like", 0) or 0):
                            all_reviews_dict[rev_id]["jumlah_like"] = likes_count
                    else:
                        all_reviews_dict[rev_id] = record
                        hotel_new_records.append(record)
                        hotel_added += 1

                # Sinkronkan fasilitas dan rating_penginapan untuk semua ulasan hotel ini di database
                for r_item in all_reviews_dict.values():
                    if str(r_item.get("data_id")) == data_id:
                        if fasilitas_str:
                            r_item["fasilitas"] = fasilitas_str
                        if rating_penginapan:
                            r_item["rating_penginapan"] = rating_penginapan

                # 5. Simpan Checkpoint ke File CSV (Fast Append per hotel)
                if hotel_new_records:
                    with open(OUTPUT_REVIEW_FILE, "a", newline="", encoding="utf-8") as f_out:
                        writer = csv.DictWriter(f_out, fieldnames=CSV_COLUMNS)
                        for rec in hotel_new_records:
                            row_dict = {c: clean_csv_cell(rec.get(c, "")) for c in CSV_COLUMNS}
                            writer.writerow(row_dict)
                elif retry_incomplete:
                    df_to_save = pd.DataFrame(list(all_reviews_dict.values()))
                    for c in CSV_COLUMNS:
                        if c not in df_to_save.columns:
                            df_to_save[c] = ""
                        elif df_to_save[c].dtype == object:
                            df_to_save[c] = df_to_save[c].apply(clean_csv_cell)
                    df_to_save = df_to_save[CSV_COLUMNS]
                    df_to_save.to_csv(OUTPUT_REVIEW_FILE, index=False, encoding="utf-8")

                total_new_reviews += hotel_added
                tot_for_hotel = sum(1 for r in all_reviews_dict.values() if r.get("data_id") == data_id)
                fac_info = f" | {len(fac_list)} fasilitas terdata" if fac_list else ""
                print(f"   ✅ Selesai: {hotel_added} ulasan baru tersimpan (Total hotel ini: {tot_for_hotel} ulasan | {hotel_replies} memiliki balasan pemilik{fac_info}).")

            except Exception as e_hotel:
                print(f"   ⚠️ Terjadi kesalahan saat scraping hotel ({e_hotel}). Melanjutkan...")
                continue

        context.close()

    print("\n" + "=" * 65)
    print(f"🎉 TAHAP 2 SELESAI! Berhasil menambah {total_new_reviews} ulasan baru.")
    print(f"📁 Database ulasan tersimpan: {len(all_reviews_dict)} ulasan di {OUTPUT_REVIEW_FILE}")
    print("=" * 65)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scraping Ulasan Penginapan Google Maps")
    parser.add_argument("--max", "-m", type=int, default=100, help="Target ulasan per hotel (default: 100, set 0 untuk semua)")
    parser.add_argument("--hotel", type=str, default="", help="Filter nama hotel tertentu (misal: Cicada)")
    parser.add_argument("--headless", action="store_true", help="Jalankan browser di background tanpa tampilan window")
    parser.add_argument("--start", "-s", type=int, default=None, help="Mulai langsung dari nomor urut penginapan tertentu (misal: 237)")
    parser.add_argument("--retry-incomplete", action="store_true", help="Ulangi hotel-hotel sebelumnya yang ulasannya belum mencapai target kuota")
    args = parser.parse_args()

    max_target = None if args.max <= 0 else args.max
    scrape_ulasan_penginapan(
        headless=args.headless,
        max_ulasan_per_hotel=max_target,
        hotel_filter=args.hotel,
        start_index=args.start,
        retry_incomplete=args.retry_incomplete
    )
