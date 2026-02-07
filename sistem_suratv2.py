import json
import os
import datetime
from docxtpl import DocxTemplate

# ================= KONFIGURASI ORGANISASI =================
TINGKATAN = "PAC"          # Ganti dengan PR, PAC, atau PC
PERIODE_IPNU = "XXI"       # Periode Romawi IPNU
PERIODE_IPPNU = "XX"       # Periode Romawi IPPNU
KODE_THN_IPNU = "7354"
KODE_THN_IPPNU = "7455"
# ==========================================================

FILE_DATABASE = "database_nomor.json"

# DAFTAR KODE SURAT
# Kamu bisa menambah kode lain di sini sesuai PD/PRT
INDEX_UMUM = ["A", "B", "C", "D", "E", "F"]
INDEX_KHUSUS = ["PH", "SK", "SP", "ST", "SMA", "SPI"] 
# PH=Peringatan Hari Besar, SK=Surat Keputusan, SP=Surat Pengantar, ST=Surat Tugas, dll.

def get_romawi(bulan):
    map_romawi = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
    return map_romawi[int(bulan)]

def load_database():
    if not os.path.exists(FILE_DATABASE):
        return {}
    with open(FILE_DATABASE, 'r') as f:
        return json.load(f)

def save_database(data):
    with open(FILE_DATABASE, 'w') as f:
        json.dump(data, f, indent=4)

def generate_nomor_surat(jenis_kop, kode_index):
    db = load_database()
    
    # Buat kunci unik, misal: "IPNU_A", "IPPNU_SK"
    kunci_db = f"{jenis_kop}_{kode_index}"
    
    # Ambil nomor terakhir, tambah 1
    nomor_urut = db.get(kunci_db, 0) + 1
    
    # Simpan nomor baru ke database
    db[kunci_db] = nomor_urut
    save_database(db)

    # Variabel Waktu
    sekarang = datetime.datetime.now()
    bulan_romawi = get_romawi(sekarang.month)
    tahun_full = sekarang.year
    tahun_2digit = str(tahun_full)[-2:]

    hasil_nomor = ""

    # === LOGIKA FORMAT NOMOR ===
    if jenis_kop == "IPNU":
        # Format: 001/PAC/A/XXVI/7354/XI/25
        nomor_str = f"{nomor_urut:03d}"
        hasil_nomor = f"{nomor_str}/{TINGKATAN}/{kode_index}/{PERIODE_IPNU}/{KODE_THN_IPNU}/{bulan_romawi}/{tahun_2digit}"
        
    elif jenis_kop == "IPPNU":
        # Format: 001/PAC/A/7455/XXVI/XI/2025
        nomor_str = f"{nomor_urut:03d}"
        hasil_nomor = f"{nomor_str}/{TINGKATAN}/{kode_index}/{KODE_THN_IPPNU}/{PERIODE_IPPNU}/{bulan_romawi}/{tahun_full}"
        
    elif jenis_kop == "BERSAMA":
        # Format: 001/PAC/A/7354-7455/XXVI/XI/2025
        nomor_str = f"{nomor_urut:03d}"
        hasil_nomor = f"{nomor_str}/{TINGKATAN}/{kode_index}/{KODE_THN_IPNU}-{KODE_THN_IPPNU}/{PERIODE_IPNU}/{bulan_romawi}/{tahun_full}"
        
    elif jenis_kop == "PANITIA":
        # Format: 01/PAC/Pan.Bersama/IPNU-IPPNU/XXVI/XI/2025
        # Panitia biasanya 2 digit (01), bukan 001
        nomor_str = f"{nomor_urut:02d}"
        # Jika kode index panitia kosong, default ke Pan.Bersama
        if not kode_index: kode_index = "Pan.Bersama"
        hasil_nomor = f"{nomor_str}/{TINGKATAN}/{kode_index}/IPNU-IPPNU/{PERIODE_IPNU}/{bulan_romawi}/{tahun_full}"

    return hasil_nomor

def main():
    print("\n=== SISTEM GENERATOR SURAT IPNU IPPNU ===")
    
    # --- LEVEL 1: PILIH KOP ---
    print("\n[ LANGKAH 1 ] Pilih Jenis Organisasi (Kop):")
    print("1. IPNU")
    print("2. IPPNU")
    print("3. BERSAMA (IPNU-IPPNU)")
    print("4. KEPANITIAAN")
    pil_kop = input(">> Masukkan angka (1-4): ")

    jenis_kop = ""
    template_file = ""

    if pil_kop == "1":
        jenis_kop = "IPNU"
        template_file = "TEMPLATE_IPNU.docx"
    elif pil_kop == "2":
        jenis_kop = "IPPNU"
        template_file = "TEMPLATE_IPPNU.docx"
    elif pil_kop == "3":
        jenis_kop = "BERSAMA"
        template_file = "TEMPLATE_BERSAMA.docx"
    elif pil_kop == "4":
        jenis_kop = "PANITIA"
        template_file = "TEMPLATE_PANITIA.docx"
    else:
        print("Pilihan tidak valid!")
        return

    # --- LEVEL 2: PILIH KATEGORI SURAT ---
    kode_index_final = ""
    
    if jenis_kop == "PANITIA":
        # Panitia biasanya langsung kode khusus
        print("\n[ LANGKAH 2 ] Masukkan Kode Kepanitiaan")
        print("Contoh: Pan.Bersama, Pan.Konferancab, Pan.Makesta")
        kode_index_final = input(">> Ketik Kode (Tekan Enter untuk default 'Pan.Bersama'): ")
        if not kode_index_final: kode_index_final = "Pan.Bersama"
    else:
        print(f"\n[ LANGKAH 2 ] Pilih Kategori Surat untuk {jenis_kop}:")
        print("1. Surat UMUM (A, B, C, dll)")
        print("2. Surat KHUSUS (SK, SP, ST, dll)")
        pil_kategori = input(">> Masukkan angka (1-2): ")

        # --- LEVEL 3: PILIH INDEX ---
        if pil_kategori == "1":
            print("\n--- Daftar Kode Umum ---")
            for i, kode in enumerate(INDEX_UMUM):
                print(f"{i+1}. Index {kode}")
            
            pilihan = input(">> Pilih nomor index (atau ketik hurufnya): ")
            if pilihan.isdigit() and 1 <= int(pilihan) <= len(INDEX_UMUM):
                kode_index_final = INDEX_UMUM[int(pilihan)-1]
            else:
                kode_index_final = pilihan.upper() # Kalau user ngetik 'A' manual

        elif pil_kategori == "2":
            print("\n--- Daftar Kode Khusus ---")
            for i, kode in enumerate(INDEX_KHUSUS):
                print(f"{i+1}. {kode}")
            
            pilihan = input(">> Pilih nomor/ketik kode (misal SK): ")
            if pilihan.isdigit() and 1 <= int(pilihan) <= len(INDEX_KHUSUS):
                kode_index_final = INDEX_KHUSUS[int(pilihan)-1]
            else:
                kode_index_final = pilihan.upper()
        else:
            print("Pilihan kategori salah.")
            return

    # GENERATE NOMOR
    print("\nMemproses Nomor Surat...")
    nomor_jadi = generate_nomor_surat(jenis_kop, kode_index_final)
    print(f"✅ Nomor Terbuat: {nomor_jadi}")

    # --- INPUT KONTEN SURAT ---
    print("\n[ LANGKAH 3 ] Isi Data Surat")
    context = {'nomor_surat': nomor_jadi}
    
    # Tips: Data ini bisa disesuaikan dengan template kamu
    context['nama_penerima'] = input("Nama Penerima (Yth.): ")
    context['alamat_penerima'] = input("Alamat Penerima: ")
    context['nama_acara'] = input("Nama Acara/Keperluan: ")
    context['hari_tanggal'] = input("Hari, Tanggal Acara: ")
    context['waktu'] = input("Waktu Acara: ")
    context['tempat'] = input("Tempat Acara: ")
    
    print("\n--- Tanggal Surat ---")
    context['tanggal_surath'] = input("Tanggal Hijriyah (misal 16 Rajab 1447 H): ")
    context['tanggal_suratm'] = input("Tanggal Masehi (misal 06 Februari 2026 M): ")

    # PROSES RENDER TEMPLATE
    try:
        doc = DocxTemplate(template_file)
        doc.render(context)
        
        nama_file = f"SURAT_{jenis_kop}_{kode_index_final}_{context['nama_penerima'].strip()}.docx"
        doc.save(nama_file)
        
        print("\n" + "="*40)
        print(f"🎉 SELESAI! File tersimpan: {nama_file}")
        print("="*40)
        
    except Exception as e:
        print(f"\n❌ GAGAL: {e}")
        print(f"Pastikan file template '{template_file}' sudah ada di folder ini.")

if __name__ == "__main__":
    main()