"""
Enricher Metadata Penginapan Google Maps:
Mengisi kolom `alamat`, `telepon`, `website`, dan `harga` yang masih kosong di `data/daftar_penginapan_jatim_mentah.csv`.
"""

import sys
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

import re
import time
from pathlib import Path
import pandas as pd
from tqdm import tqdm
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
HOTEL_FILE = DATA_DIR / "daftar_penginapan_jatim_mentah.csv"

def clean_icon_text(txt: str) -> str:
    """Membersihkan karakter ikon font Google Maps."""
    if not txt:
        return ""
    # Hapus karakter private use area Unicode (ikon Material)
    cleaned = re.sub(r"[\ue000-\uf8ff]", "", txt).strip()
    return cleaned

def to_usd_price(raw_price: str) -> str:
    """Mengonversi teks harga (SGD/IDR/USD) ke format standar USD (contoh: 'USD 15')."""
    if not raw_price:
        return ""
    
    txt = clean_icon_text(raw_price).replace('\xa0', ' ').replace('\u2009', ' ').strip()
    
    # 1. Jika dalam SGD (contoh: 'SGD 19', 'SGD 439')
    m_sgd = re.search(r'SGD\s*([\d.,]+)', txt, re.IGNORECASE)
    if m_sgd:
        val_str = m_sgd.group(1).replace(',', '')
        try:
            val_sgd = float(val_str)
            val_usd = round(val_sgd * 0.77)  # 1 SGD ~ 0.77 USD
            return f"USD {val_usd}"
        except Exception:
            return f"USD {val_str}"
            
    # 2. Jika dalam IDR / Rp (contoh: 'Rp 228.536', 'IDR 6,270,239')
    m_idr = re.search(r'(?:Rp|IDR)\s*([\d.,]+)', txt, re.IGNORECASE)
    if m_idr:
        num_str = re.sub(r'[.,]', '', m_idr.group(1))
        try:
            val_idr = float(num_str)
            val_usd = round(val_idr / 15600)  # 1 USD ~ 15,600 IDR
            return f"USD {val_usd}"
        except Exception:
            pass

    # 3. Jika sudah USD / $
    m_usd = re.search(r'(?:USD|\$)\s*([\d.,]+)', txt, re.IGNORECASE)
    if m_usd:
        return f"USD {m_usd.group(1).strip()}"
        
    return txt

def save_csv_safely(df_to_save: pd.DataFrame, file_path: Path):
    """Menyimpan dataframe ke CSV dengan aman (mencegah error jika file dibuka di Microsoft Excel)."""
    try:
        df_to_save.to_csv(file_path, index=False, encoding="utf-8")
    except PermissionError:
        print(f"\n⚠️ PERINGATAN: File '{file_path.name}' sedang terbuka di aplikasi lain (seperti Microsoft Excel)!")
        print("⚠️ Harap tutup file tersebut agar perubahan dapat disimpan permanen.")
        backup_path = file_path.with_name(f"{file_path.stem}_backup.csv")
        try:
            df_to_save.to_csv(backup_path, index=False, encoding="utf-8")
            print(f"💾 Progres otomatis diamankan sementara di: {backup_path.name}")
        except Exception:
            pass

USER_DATA_DIR = DATA_DIR / "browser_profile"

def enrich_hotels(headless: bool = True):
    if not HOTEL_FILE.exists():
        print(f"File {HOTEL_FILE} tidak ditemukan!")
        return

    df = pd.read_csv(HOTEL_FILE)
    print(f"📋 Total penginapan: {len(df)}")

    # Pastikan kolom ada dan bertipe object/string agar pandas tidak melempar TypeError float64
    for col in ["harga", "telepon", "website", "alamat"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype("object").fillna("")

    # Cari baris yang telepon atau websitenya masih kosong
    mask_empty = (df["telepon"] == "") | (df["website"] == "")
    hotels_to_enrich = df[mask_empty].copy()
    print(f"🎯 Penginapan yang perlu dilengkapi (telepon/web): {len(hotels_to_enrich)}")

    if len(hotels_to_enrich) == 0:
        print("🎉 Semua detail penginapan sudah lengkap 100%!")
        return

    with sync_playwright() as p:
        USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
        # Gunakan profil browser asli yang sudah login agar tidak terkena blokir Google
        launch_args = [
            "--lang=id-ID,id",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox"
        ]
        try:
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(USER_DATA_DIR),
                channel="chrome",
                headless=headless,
                args=launch_args,
                viewport={"width": 1400, "height": 900},
                locale="id-ID"
            )
        except Exception:
            # Fallback ke Playwright Chromium jika Google Chrome sistem sedang terbuka/terkunci
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(USER_DATA_DIR),
                headless=headless,
                args=launch_args,
                viewport={"width": 1400, "height": 900},
                locale="id-ID"
            )

        page = context.pages[0] if context.pages else context.new_page()

        # Blokir file gambar & media agar cepat
        def intercept_heavy(route):
            if route.request.resource_type in ["image", "media"]:
                route.abort()
            else:
                route.continue_()

        page.route("**/*", intercept_heavy)

        count = 0
        for idx, row in tqdm(hotels_to_enrich.iterrows(), total=len(hotels_to_enrich), desc="Melengkapi Detail Hotel"):
            url = row.get("url", "")
            if not url or pd.isnull(url):
                continue

            try:
                page.goto(url, timeout=25000, wait_until="domcontentloaded")
                try:
                    page.wait_for_selector('div.m6QErb', timeout=6000)
                except Exception:
                    pass
                time.sleep(1.5)

                details = page.evaluate(r"""() => {
                    // 1. Alamat
                    let alamat = '';
                    const addrEl = document.querySelector('button[data-item-id="address"], button[aria-label*="Alamat:"]');
                    if (addrEl) {
                        const aria = addrEl.getAttribute('aria-label') || '';
                        alamat = aria.replace(/^Alamat:\s*/i, '').trim() || addrEl.innerText.trim();
                    }

                    // 2. Telepon
                    let telepon = '';
                    const phoneEl = document.querySelector('button[data-item-id*="phone"], button[aria-label*="Telepon:"], a[href^="tel:"]');
                    if (phoneEl) {
                        const aria = phoneEl.getAttribute('aria-label') || '';
                        telepon = aria.replace(/^Telepon:\s*/i, '').trim() || phoneEl.innerText.trim();
                        telepon = telepon.replace(/^tel:\+?/i, '').trim();
                    }

                    // 3. Website
                    let website = '';
                    const webEl = document.querySelector('a[data-item-id="authority"], a[aria-label*="Situs Web" i], a[aria-label*="Website" i], a[data-tooltip*="situs web" i]');
                    if (webEl) {
                        let href = webEl.getAttribute('href') || '';
                        const m = href.match(/[?&]q=([^&]+)/);
                        if (m) {
                            try { href = decodeURIComponent(m[1]); } catch(e){}
                        }
                        website = href;
                    }

                    // 4. Harga (jika tersedia di Google Hotels / Partners)
                    let harga = '';
                    const priceEl = document.querySelector('button[aria-label*="per malam"], *[aria-label*="per malam"], div.drwW4b, div[aria-label*="Rp"], span[aria-label*="Rp"], .fontTitleLarge');
                    if (priceEl) {
                        const aria = priceEl.getAttribute('aria-label') || '';
                        const txt = priceEl.innerText || '';
                        const mDari = aria.match(/Dari\s+([^per]+?)\s+per\s+malam/i);
                        if (mDari) {
                            harga = mDari[1].trim();
                        } else {
                            const m = (aria + ' ' + txt).match(/(?:Rp|IDR|SGD|USD|AUD|EUR|[$€¥])\s*[\d.,]+/i);
                            if (m) harga = m[0];
                        }
                    }
                    if (!harga) {
                        for (let el of document.querySelectorAll('button, span, div')) {
                            const t = (el.innerText || '').trim();
                            if (/^(?:Rp|IDR|SGD|USD)\s*[\d.,]+$/i.test(t) && t.length < 25) {
                                harga = t;
                                break;
                            }
                        }
                    }

                    return { alamat, telepon, website, harga };
                }""")

                # Update jika ditemukan
                curr_alamat = str(df.at[idx, "alamat"]).strip()
                hotel_name = str(row.get("nama_penginapan", "")).strip().lower()
                if details["alamat"]:
                    # Update alamat jika masih kosong atau hanya berisi nama hotel
                    if not curr_alamat or curr_alamat.lower() == hotel_name:
                        df.at[idx, "alamat"] = clean_icon_text(details["alamat"])

                val_tel = str(df.at[idx, "telepon"]).strip()
                if details["telepon"] and (not val_tel or val_tel.lower() == "nan"):
                    df.at[idx, "telepon"] = clean_icon_text(details["telepon"])

                val_web = str(df.at[idx, "website"]).strip()
                if details["website"] and (not val_web or val_web.lower() == "nan"):
                    df.at[idx, "website"] = details["website"]

                val_hrg = str(df.at[idx, "harga"]).strip()
                if details["harga"] and (not val_hrg or val_hrg.lower() == "nan"):
                    df.at[idx, "harga"] = to_usd_price(details["harga"])

                count += 1
                # Simpan berkala setiap 5 hotel
                if count % 5 == 0:
                    save_csv_safely(df, HOTEL_FILE)

            except Exception:
                continue

        context.close()

    save_csv_safely(df, HOTEL_FILE)
    
    total_tel = (df["telepon"].astype(str).str.strip() != "").sum()
    total_web = (df["website"].astype(str).str.strip() != "").sum()
    total_hrg = (df["harga"].astype(str).str.strip() != "").sum()
    print("\n" + "=" * 60)
    print("✅ Berhasil melengkapi metadata penginapan!")
    print(f"📞 Total penginapan bertelepon: {total_tel}/{len(df)}")
    print(f"🌐 Total penginapan berwebsite: {total_web}/{len(df)}")
    print(f"💰 Total penginapan berharga:   {total_hrg}/{len(df)}")
    print(f"📁 Disimpan di: {HOTEL_FILE}")
    print("=" * 60)

if __name__ == "__main__":
    enrich_hotels(headless=True)

