"""
TAHAP 3: Pemrosesan Dataset ABSA & Export ke Excel Lengkap
Input:
  - data/daftar_penginapan_jatim_mentah.csv
  - data/raw_ulasan_checkpoint.csv
Output:
  - data/dataset_absa_penginapan_halal_jatim.csv
  - data/Dataset ABSA Penginapan Jawa Timur.xlsx (dengan sheet: ulasan, kalimat_aspek, daftar_penginapan, kamus_aspek, dan sheet statistik)
"""

import os
import sys
import re
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
from langdetect import detect

# Pastikan output konsol Windows mendukung UTF-8
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
HOTEL_FILE = DATA_DIR / "daftar_penginapan_jatim_mentah.csv"
REVIEW_FILE = DATA_DIR / "raw_ulasan_checkpoint.csv"
OUTPUT_CSV = DATA_DIR / "dataset_absa_penginapan_halal_jatim.csv"
OUTPUT_EXCEL = DATA_DIR / "Dataset ABSA Penginapan Jawa Timur.xlsx"

# Kamus Aspek Halal & Umum (sesuai referensi dataset asli)
KAMUS_ASPEK = {
    "makanan_halal": {
        "regex": r"halal|haram|non[- ]?halal|pork|babi|bacon|ham\b|sarapan|breakfast|makanan|makan|food|menu|restoran|restaurant|dapur|kitchen|catering|buffet|prasmanan|sertifikat|certified|certification|MUI",
        "inti_halal": True
    },
    "fasilitas_ibadah": {
        "regex": r"musala|mushola|musholla|masjid|mosque|sajadah|sarung|mukena|kiblat|qibla|qiblah|kibla|sholat|shalat|salat|prayer|pray|adzan|azan|adhan|quran|al[- ]?qur|wudhu|wudu|ablution",
        "inti_halal": True
    },
    "kesesuaian_syariah": {
        "regex": r"alkohol|alcohol|miras|bir\b|beer|wine|bar\b|nightclub|club|party|pesta|privasi|privacy|privat|hijab|berhijab|kolam renang|pool|pemisahan|terpisah|separate|islami|islamic|syariah|sharia|muslim|moslem|family friendly|ramah keluarga|mahram|non[- ]?muslim",
        "inti_halal": True
    },
    "kebersihan": {
        "regex": r"bersih|kotor|clean|dirty|bau|smell|odor|debu|dust|jorok|higienis|hygien|rapi|tidy|sprei|seprai|handuk|towel|kamar mandi|bathroom|toilet|kutu|bed ?bug|kecoa|cockroach|semut|ant\b",
        "inti_halal": False
    },
    "pelayanan": {
        "regex": r"layanan|pelayanan|service|staf|staff|karyawan|pegawai|resepsionis|receptionist|ramah|friendly|helpful|sopan|polite|responsif|responsive|lambat|slow|check[- ]?in|check[- ]?out|pelayan|waiter|housekeeping|manajemen|management|hospitality|keramahan",
        "inti_halal": False
    },
    "kamar_fasilitas": {
        "regex": r"kamar|room|kasur|bed|bantal|pillow|ac\b|air conditioner|pendingin|wifi|wi-fi|internet|tv\b|televisi|shower|air panas|hot water|water heater|kulkas|fridge|balkon|balcony|fasilitas|facilit|gym|spa|parkir|parking|lift|elevator|luas|sempit|spacious|small",
        "inti_halal": False
    },
    "lokasi_akses": {
        "regex": r"lokasi|location|strategis|strategic|dekat|near|close|jauh|far|akses|access|pantai|beach|bandara|airport|walking distance|jalan kaki|macet|traffic|sekitar|area|lingkungan|neighborhood|view|pemandangan|pusat|central",
        "inti_halal": False
    },
    "harga_value": {
        "regex": r"harga|price|priced|murah|cheap|mahal|expensive|worth|value|terjangkau|affordable|biaya|tarif|rate\b|promo|diskon|discount|sebanding|overpriced|budget",
        "inti_halal": False
    }
}

# Regex khusus untuk kolom flag aspek halal per ulasan
REGEX_FLAGS_HALAL = {
    "flag_makanan_halal": r"\b(halal|haram|non[- ]?halal|pork|babi|bacon|ham|sarapan halal|makanan halal)\b",
    "flag_fasilitas_ibadah": r"\b(musala|mushola|musholla|masjid|sajadah|sarung|mukena|kiblat|sholat|shalat|salat|wudhu|adzan)\b",
    "flag_privasi_kamar": r"\b(privasi|privacy|privat|tertutup|terpisah)\b",
    "flag_kolam_renang": r"\b(kolam renang|kolam renang privat|private pool|swimming pool)\b",
    "flag_citra_syariah": r"\b(syariah|sharia|islami|islamic|muslim|moslem|ramah keluarga|family friendly|mahram|alkohol|miras|bir|bar)\b"
}

def parse_relative_date(date_str: str) -> tuple[datetime, str]:
    """Mengubah tanggal relatif Google Maps (misal '2 bulan lalu') menjadi tanggal ISO."""
    now = datetime.now()
    if not isinstance(date_str, str) or not date_str.strip():
        iso_str = now.strftime("%Y-%m-%dT%H:%M:%SZ")
        return now, iso_str

    clean_str = re.sub(r"^(Diedit|Ditulis)\s+", "", date_str, flags=re.IGNORECASE).strip()
    match = re.search(r"(\d+)\s+(hari|minggu|bulan|tahun|jam|menit)", clean_str, re.IGNORECASE)
    
    if match:
        val = int(match.group(1))
        unit = match.group(2).lower()
        if "hari" in unit:
            dt = now - timedelta(days=val)
        elif "minggu" in unit:
            dt = now - timedelta(weeks=val)
        elif "bulan" in unit:
            dt = now - timedelta(days=val * 30)
        elif "tahun" in unit:
            dt = now - timedelta(days=val * 365)
        elif "jam" in unit:
            dt = now - timedelta(hours=val)
        elif "menit" in unit:
            dt = now - timedelta(minutes=val)
        else:
            dt = now
    elif "kemarin" in clean_str.lower():
        dt = now - timedelta(days=1)
    else:
        dt = now

    iso_str = dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    return dt, iso_str

def bersihkan_teks(teks: str) -> str:
    """Membersihkan teks ulasan dari noise, link URL, dan spasi ganda."""
    if not isinstance(teks, str):
        return ""
    teks = re.sub(r"http\S+|www\S+", "", teks)
    teks = re.sub(r"[\r\n]+", " ", teks)
    teks = re.sub(r"\s+", " ", teks).strip()
    return teks

def clean_csv_cell(val) -> str:
    """Membersihkan newline agar tidak memecah baris di CSV dan Excel."""
    if val is None or pd.isna(val):
        return ""
    txt = str(val)
    txt = re.sub(r"[\r\n]+", " ", txt)
    txt = txt.replace("\t", " ")
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt

def deteksi_bahasa(teks: str) -> str:
    """Mendeteksi bahasa ulasan menggunakan langdetect."""
    if not isinstance(teks, str) or len(teks.strip()) < 3:
        return "unk"
    try:
        return detect(teks)
    except Exception:
        return "unk"

def proses_pipeline():
    print("=" * 60)
    print("🚀 MEMULAI TAHAP 3: PREPROCESSING & EXPORT DATASET ABSA")
    print("=" * 60)

    if not HOTEL_FILE.exists() or not REVIEW_FILE.exists():
        print("❌ File hotel atau ulasan tidak ditemukan. Pastikan Tahap 1 & 2 sudah dijalankan!")
        return

    df_hotel = pd.read_csv(HOTEL_FILE)
    df_review = pd.read_csv(REVIEW_FILE)

    print(f"📋 Memproses {len(df_hotel)} penginapan dan {len(df_review)} ulasan...")

    # Standarisasi kolom join
    df_hotel["data_id"] = df_hotel["data_id"].astype(str)
    df_review["data_id"] = df_review["data_id"].astype(str)

    # 1. Join ulasan dengan profil penginapan
    hotel_cols = ["data_id", "nama_penginapan", "tipe", "alamat", "lat", "lon", "rating_rata2", "jumlah_ulasan_total", "harga", "telepon", "website", "place_id"]
    available_cols = [c for c in hotel_cols if c in df_hotel.columns]
    
    # Buat lookup penginapan
    df_hotel_unique = df_hotel.drop_duplicates(subset=["data_id"]).copy()
    hotel_dict = df_hotel_unique.set_index("data_id").to_dict(orient="index")

    # 2. Proses setiap baris ulasan
    processed_reviews = []
    kalimat_aspek_list = []

    for _, row in df_review.iterrows():
        did = str(row.get("data_id", ""))
        h_info = hotel_dict.get(did, {})

        nama_tempat = str(row.get("nama_penginapan") or h_info.get("nama_penginapan", ""))
        alamat = str(h_info.get("alamat", ""))
        kategori = str(h_info.get("tipe", "Hotel"))
        raw_hotel_rat = row.get("rating_penginapan") if (pd.notnull(row.get("rating_penginapan")) and str(row.get("rating_penginapan")).strip() != "") else h_info.get("rating_rata2")
        try:
            rating_tempat = f"{float(raw_hotel_rat):.1f}" if pd.notnull(raw_hotel_rat) and str(raw_hotel_rat).strip() != "" else ""
        except Exception:
            rating_tempat = str(raw_hotel_rat) if pd.notnull(raw_hotel_rat) else ""

        total_ulasan_tempat = h_info.get("jumlah_ulasan_total", 0)
        harga = h_info.get("harga", "")
        telepon = h_info.get("telepon", "")
        website = h_info.get("website", "")
        lat = h_info.get("lat")
        lon = h_info.get("lon")
        place_id = str(h_info.get("place_id", ""))

        review_id = str(row.get("review_id", ""))
        reviewer_total = row.get("reviewer_total_ulasan", 1)
        raw_star = row.get("rating_bintang", 5)
        try:
            rating_ulasan = f"{float(raw_star):.1f}" if pd.notnull(raw_star) and str(raw_star).strip() != "" else "5.0"
        except Exception:
            rating_ulasan = "5.0"
        tanggal_str = str(row.get("tanggal", ""))
        ulasan_asli = str(row.get("teks_ulasan", ""))
        likes = row.get("jumlah_like", 0)

        # Parse tanggal
        dt_obj, iso_str = parse_relative_date(tanggal_str)
        teks_bersih = bersihkan_teks(ulasan_asli)
        bahasa_code = deteksi_bahasa(teks_bersih)

        # Status halal & bukti
        is_halal_name = bool(re.search(r"\b(syariah|sharia|muslim|halal)\b", nama_tempat, re.IGNORECASE))
        
        status_halal = "tanpa_indikasi_non_halal"
        bukti_halal = ""

        if is_halal_name:
            status_halal = "halal_tanpa_bukti_sertifikat"
            bukti_halal = f"[tempat] {nama_tempat}"
        elif re.search(r"\b(sertifikat|certified|logo\s+halal|mui)\b", teks_bersih, re.IGNORECASE):
            status_halal = "bersertifikat_diklaim"
            bukti_halal = f"[ulasan] {teks_bersih[:100]}"
        elif re.search(r"\b(halal|syariah|muslim friendly)\b", teks_bersih, re.IGNORECASE):
            status_halal = "halal_tanpa_bukti_sertifikat"
            bukti_halal = f"[ulasan] {teks_bersih[:100]}"

        # Record untuk sheet ulasan (26 Kolom)
        rev_item = {
            "place_id": place_id,
            "data_id": did,
            "nama_tempat": nama_tempat,
            "alamat": alamat,
            "kategori": kategori,
            "rating_tempat": rating_tempat,
            "total_ulasan_tempat": total_ulasan_tempat,
            "harga": harga,
            "telepon": telepon,
            "website": website,
            "latitude": lat,
            "longitude": lon,
            "status_halal": status_halal,
            "bukti_halal": bukti_halal,
            "review_id": review_id,
            "reviewer_total_ulasan": reviewer_total,
            "rating_ulasan": rating_ulasan,
            "tanggal": tanggal_str,
            "iso_date": iso_str,
            "ulasan": ulasan_asli,
            "likes": likes,
            "label_halal_nama": is_halal_name,
            "teks_ulasan": ulasan_asli,
            "tanggal_iso": dt_obj,
            "teks_bersih": teks_bersih,
            "bahasa": bahasa_code,
            # Kolom tambahan untuk CSV akhir
            "kabupaten": str(row.get("kabupaten", "")),
            "kelompok_sampel": str(row.get("kelompok_sampel", "umum")),
            "kata_kunci": str(row.get("kata_kunci", "_umum")),
            "teks_terjemahan_google": str(row.get("teks_terjemahan_google", "")),
            "rating_penginapan": row.get("rating_penginapan", rating_tempat),
            "ada_balasan_pemilik": row.get("ada_balasan_pemilik", False),
            "balasan_pemilik": str(row.get("balasan_pemilik", "")),
            "fasilitas": str(row.get("fasilitas", "")),
            "sumber": "Google Maps"
        }
        processed_reviews.append(rev_item)

        # 3. Ekstraksi pasangan Kalimat-Aspek (untuk sheet kalimat_aspek)
        kalimat_split = [k.strip() for k in re.split(r"[.!?\n]+", teks_bersih) if len(k.strip()) > 3]
        for k_idx, kal in enumerate(kalimat_split):
            for aspek_nama, aspek_cfg in KAMUS_ASPEK.items():
                if re.search(aspek_cfg["regex"], kal, re.IGNORECASE):
                    kalimat_aspek_list.append({
                        "review_id": review_id,
                        "data_id": did,
                        "place_id": place_id,
                        "nama_tempat": nama_tempat,
                        "status_halal": status_halal,
                        "kalimat_ke": k_idx,
                        "kalimat": kal,
                        "aspek": aspek_nama,
                        "aspek_inti_halal": aspek_cfg["inti_halal"],
                        "rating_ulasan": rating_ulasan,
                        "bahasa": bahasa_code,
                        "tanggal_iso": dt_obj
                    })

    df_full = pd.DataFrame(processed_reviews)
    df_kalimat = pd.DataFrame(kalimat_aspek_list)

    # 4. Tambahkan Flag Aspek Halal & Kolom Anotasi untuk dataset_absa_penginapan_halal_jatim.csv
    df_absa_csv = df_full.copy()
    df_absa_csv["tahun"] = df_absa_csv["tanggal_iso"].apply(lambda d: d.year if pd.notnull(d) else "")
    df_absa_csv["jumlah_kata"] = df_absa_csv["teks_bersih"].apply(lambda t: len(str(t).split()))

    for flag_col, pattern in REGEX_FLAGS_HALAL.items():
        df_absa_csv[flag_col] = df_absa_csv["teks_bersih"].apply(lambda t: bool(re.search(pattern, str(t), re.IGNORECASE)))

    df_absa_csv["flag_ada_aspek_halal"] = (
        df_absa_csv["flag_makanan_halal"] |
        df_absa_csv["flag_fasilitas_ibadah"] |
        df_absa_csv["flag_privasi_kamar"] |
        df_absa_csv["flag_kolam_renang"] |
        df_absa_csv["flag_citra_syariah"]
    )

    # Kolom label anotasi kosong siap isi
    label_cols = [
        "label_kebersihan", "label_layanan_staf", "label_kamar_fasilitas",
        "label_lokasi", "label_harga", "label_makanan_halal",
        "label_fasilitas_ibadah", "label_privasi_kamar", "label_kolam_renang",
        "anotator_1", "anotator_2"
    ]
    for lc in label_cols:
        df_absa_csv[lc] = ""

    # Simpan dataset CSV Tahap 3 (Pastikan setiap ulasan tepat 1 baris)
    for col in df_absa_csv.columns:
        if df_absa_csv[col].dtype == object:
            df_absa_csv[col] = df_absa_csv[col].apply(clean_csv_cell)
    df_absa_csv.to_csv(OUTPUT_CSV, index=False, encoding="utf-8")
    print(f"✅ File CSV akhir tersimpan: {OUTPUT_CSV}")

    # 5. Siapkan Sheet-Sheet untuk File Excel Lengkap
    # Sheet 1: ulasan (Semua Kolom Profil Penginapan + Detail Ulasan + Fasilitas)
    kolom_ulasan_lengkap = [
        "place_id", "data_id", "nama_tempat", "kategori", "alamat", "kabupaten",
        "kelompok_sampel", "area_query", "kab_query", "query_sumber", "kata_kunci", "url",
        "rating_tempat", "total_ulasan_tempat", "harga", "telepon", "website",
        "latitude", "longitude", "fasilitas", "status_halal", "bukti_halal", "label_halal_nama",
        "review_id", "reviewer_total_ulasan", "rating_ulasan", "tanggal", "iso_date",
        "ulasan", "teks_bersih", "teks_terjemahan_google", "likes", "ada_balasan_pemilik",
        "balasan_pemilik", "bahasa", "sumber"
    ]
    for c in kolom_ulasan_lengkap:
        if c not in df_full.columns:
            # Map kolom yang mungkin ada di df_hotel
            if c in df_hotel.columns:
                df_full[c] = df_full["data_id"].map(df_hotel.set_index("data_id")[c].to_dict()).fillna("")
            else:
                df_full[c] = ""
    df_ulasan_sheet = df_full[kolom_ulasan_lengkap].copy()

    # Sheet 2: daftar_penginapan (Profil Lengkap 464 Penginapan + Fasilitas)
    kolom_hotel_lengkap = [
        "data_id", "place_id", "nama_penginapan", "tipe", "alamat", "kabupaten",
        "kelompok_sampel", "area_query", "kab_query", "query_sumber", "url",
        "lat", "lon", "rating_rata2", "jumlah_ulasan_total", "harga", "telepon", "website",
        "fasilitas", "status_halal", "bukti_halal", "label_halal_nama"
    ]
    status_map = df_full.drop_duplicates("data_id").set_index("data_id")["status_halal"].to_dict()
    bukti_map = df_full.drop_duplicates("data_id").set_index("data_id")["bukti_halal"].to_dict()
    halal_nama_map = df_full.drop_duplicates("data_id").set_index("data_id")["label_halal_nama"].to_dict()

    df_hotel["status_halal"] = df_hotel["data_id"].map(status_map).fillna("tanpa_indikasi_non_halal")
    df_hotel["bukti_halal"] = df_hotel["data_id"].map(bukti_map).fillna("")
    df_hotel["label_halal_nama"] = df_hotel["data_id"].map(halal_nama_map).fillna(False)

    for c in kolom_hotel_lengkap:
        if c not in df_hotel.columns:
            df_hotel[c] = ""
    df_hotel_sheet = df_hotel[kolom_hotel_lengkap].copy()

    # Sheet 3: olah_fasilitas (Matrix One-Hot Indicator per Penginapan)
    TOP_FACILITIES = [
        "Wi-Fi gratis", "AC", "Restoran", "Sarapan", "Sarapan gratis",
        "Kolam renang luar", "Kolam renang", "Parkir gratis", "Pusat kebugaran",
        "Spa", "Layanan kamar", "Jemputan bandara", "Layanan binatu",
        "Cocok untuk anak-anak", "Akses ke pantai", "Bar", "Boleh bawa hewan",
        "Bebas asap rokok", "Pusat bisnis", "Dapat diakses"
    ]

    rows_olah = []
    for _, r in df_hotel_sheet.iterrows():
        f_str = str(r.get("fasilitas", "")).strip()
        f_list = [x.strip() for x in f_str.split(",") if x.strip()]
        item = {
            "data_id": r.get("data_id"),
            "nama_penginapan": r.get("nama_penginapan"),
            "tipe": r.get("tipe"),
            "kabupaten": r.get("kabupaten"),
            "kelompok_sampel": r.get("kelompok_sampel"),
            "status_halal": r.get("status_halal"),
            "rating_rata2": r.get("rating_rata2"),
            "jumlah_ulasan_total": r.get("jumlah_ulasan_total"),
            "harga": r.get("harga"),
            "telepon": r.get("telepon"),
            "website": r.get("website"),
            "total_fasilitas": len(f_list),
            "daftar_fasilitas": f_str
        }
        for tf in TOP_FACILITIES:
            col_name = tf.lower().replace(" ", "_").replace("-", "_")
            item[col_name] = "Ya" if tf in f_list else "Tidak"
        rows_olah.append(item)

    df_olah_fasilitas = pd.DataFrame(rows_olah)

    # Sheet 4: stat_fasilitas (Rekap ranking fasilitas di Jawa Timur)
    all_facs = set()
    for f in df_hotel_sheet["fasilitas"]:
        for item in str(f).split(","):
            if item.strip():
                all_facs.add(item.strip())

    stat_rows = []
    total_h = len(df_hotel_sheet)
    df_hotel_sheet["rating_num"] = pd.to_numeric(df_hotel_sheet["rating_rata2"], errors="coerce")

    for fac in sorted(all_facs):
        mask = df_hotel_sheet["fasilitas"].str.contains(fac, regex=False)
        sub = df_hotel_sheet[mask]
        n = len(sub)
        n_halal = (sub["kelompok_sampel"] == "halal").sum()
        n_umum = (sub["kelompok_sampel"] != "halal").sum()
        avg_r = sub["rating_num"].mean()
        stat_rows.append({
            "fasilitas": fac,
            "jumlah_penginapan": n,
            "persentase_penginapan": f"{n/total_h*100:.1f}%",
            "penginapan_halal": n_halal,
            "penginapan_umum": n_umum,
            "rata_rata_rating": round(avg_r, 2) if pd.notnull(avg_r) else 0.0
        })

    df_stat_fasilitas = pd.DataFrame(stat_rows).sort_values(by="jumlah_penginapan", ascending=False).reset_index(drop=True)

    # Sheet 5: kalimat_aspek (df_kalimat)

    # Sheet 6: kamus_aspek
    kamus_rows = [{"aspek": k, "regex": v["regex"], "inti_halal": v["inti_halal"]} for k, v in KAMUS_ASPEK.items()]
    df_kamus_sheet = pd.DataFrame(kamus_rows)

    # Sheet Statistik
    df_stat_ringkas = pd.DataFrame([
        {"Keterangan": "Jumlah penginapan dalam dataset", "Nilai": len(df_hotel_sheet)},
        {"Keterangan": "Jumlah ulasan (setelah pembersihan)", "Nilai": len(df_ulasan_sheet)},
        {"Keterangan": "Jumlah pasangan kalimat-aspek", "Nilai": len(df_kalimat)},
        {"Keterangan": "Jumlah jenis fasilitas terdata", "Nilai": len(df_stat_fasilitas)},
        {"Keterangan": "Rentang tanggal ulasan (awal)", "Nilai": str(df_full['tanggal_iso'].min())},
        {"Keterangan": "Rentang tanggal ulasan (akhir)", "Nilai": str(df_full['tanggal_iso'].max())}
    ])

    df_stat_status_halal = df_hotel_sheet["status_halal"].value_counts().reset_index()
    df_stat_status_halal.columns = ["status_halal", "jumlah_penginapan"]

    if not df_kalimat.empty:
        df_kalimat_stat = df_kalimat.copy()
        df_kalimat_stat["rating_ulasan"] = pd.to_numeric(df_kalimat_stat["rating_ulasan"], errors="coerce")
        df_stat_aspek = df_kalimat_stat.groupby("aspek").agg(
            jumlah_kalimat=("kalimat", "count"),
            jumlah_ulasan=("review_id", "nunique"),
            rata_rata_rating=("rating_ulasan", "mean")
        ).round(2).reset_index()
    else:
        df_stat_aspek = pd.DataFrame(columns=["aspek", "jumlah_kalimat", "jumlah_ulasan", "rata_rata_rating"])

    df_stat_bahasa = df_ulasan_sheet["bahasa"].value_counts().reset_index()
    df_stat_bahasa.columns = ["bahasa", "jumlah_ulasan"]

    df_stat_rating = df_ulasan_sheet["rating_ulasan"].value_counts().reset_index()
    df_stat_rating.columns = ["rating", "jumlah_ulasan"]

    df_stat_tahun = df_absa_csv["tahun"].value_counts().reset_index()
    df_stat_tahun.columns = ["tahun", "jumlah_ulasan"]

    df_ulasan_num = df_ulasan_sheet.copy()
    df_ulasan_num["rating_ulasan"] = pd.to_numeric(df_ulasan_num["rating_ulasan"], errors="coerce")
    df_stat_penginapan = df_ulasan_num.groupby(["data_id", "nama_tempat", "status_halal"]).agg(
        jumlah_ulasan=("review_id", "count"),
        rata_rata_rating=("rating_ulasan", "mean")
    ).round(2).reset_index()

    def sanitize_for_excel(df_in: pd.DataFrame) -> pd.DataFrame:
        """Mencegah karakter '=' atau '@' di awal teks ulasan memicu error formula di Excel."""
        df_out = df_in.copy()
        for col in df_out.select_dtypes(include=["object"]).columns:
            df_out[col] = df_out[col].apply(
                lambda x: re.sub(r"^[=@]+\s*", "", str(x)) if isinstance(x, str) and len(x) > 0 and x[0] in ["=", "@"] else x
            )
        return df_out

    # Tulis ke file Excel dengan seluruh 13 sheet sekaligus
    print(f"⏳ Mengekspor ke Excel: '{OUTPUT_EXCEL}'...")
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        sanitize_for_excel(df_ulasan_sheet).to_excel(writer, sheet_name="ulasan", index=False)
        sanitize_for_excel(df_kalimat).to_excel(writer, sheet_name="kalimat_aspek", index=False)
        sanitize_for_excel(df_hotel_sheet).to_excel(writer, sheet_name="daftar_penginapan", index=False)
        sanitize_for_excel(df_olah_fasilitas).to_excel(writer, sheet_name="olah_fasilitas", index=False)
        sanitize_for_excel(df_stat_fasilitas).to_excel(writer, sheet_name="stat_fasilitas", index=False)
        df_stat_ringkas.to_excel(writer, sheet_name="stat_ringkas", index=False)
        df_stat_status_halal.to_excel(writer, sheet_name="stat_status_halal", index=False)
        df_stat_aspek.to_excel(writer, sheet_name="stat_aspek", index=False)
        df_stat_bahasa.to_excel(writer, sheet_name="stat_bahasa", index=False)
        df_stat_rating.to_excel(writer, sheet_name="stat_rating", index=False)
        df_stat_tahun.to_excel(writer, sheet_name="stat_tahun", index=False)
        df_stat_penginapan.to_excel(writer, sheet_name="stat_penginapan", index=False)
        df_kamus_sheet.to_excel(writer, sheet_name="kamus_aspek", index=False)

    print("\n" + "=" * 60)
    print("🎉 SEMUA PROSES BERHASIL SELESAI 100%!")
    print(f"📁 Excel Lengkap: {OUTPUT_EXCEL}")
    print("=" * 60)

if __name__ == "__main__":
    proses_pipeline()
