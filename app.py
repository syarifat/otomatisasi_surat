import streamlit as st
import mysql.connector
from docxtpl import DocxTemplate
import datetime
import io
import os

# ================= KONFIGURASI UMUM =================
TINGKATAN = "PAC"
PERIODE_IPNU = "IX"
PERIODE_IPPNU = "IX"
KODE_THN_IPNU = "7354"
KODE_THN_IPPNU = "7455"

# --- UPDATE KONFIGURASI INDEX ---
# Semua KOP (IPNU, IPPNU, BERSAMA) hanya menggunakan A, B, C
INDEX_FULL = ["A", "B", "C"] 

# Update Index Khusus (Surat Mandat, Surat Pengantar, Surat Tugas)
INDEX_KHUSUS = ["SM", "SPt", "ST"] 

# ================= KONEKSI DATABASE (TiDB) =================
def init_connection():
    return mysql.connector.connect(
        host=st.secrets["tidb"]["host"],
        port=st.secrets["tidb"]["port"],
        user=st.secrets["tidb"]["user"],
        password=st.secrets["tidb"]["password"],
        database=st.secrets["tidb"]["database"]
    )

def init_db():
    """Membuat tabel jika belum ada"""
    try:
        conn = init_connection()
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS counter_surat (
                kode_surat VARCHAR(50) PRIMARY KEY,
                nomor_terakhir INT DEFAULT 0,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Error Database Init: {e}")

def get_nomor_terakhir(kode_unik):
    conn = init_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT nomor_terakhir FROM counter_surat WHERE kode_surat = %s", (kode_unik,))
    result = cursor.fetchone()
    conn.close()
    if result:
        return result[0]
    return 0

def increment_nomor(kode_unik):
    conn = init_connection()
    cursor = conn.cursor()
    sql = """
        INSERT INTO counter_surat (kode_surat, nomor_terakhir) 
        VALUES (%s, 1) 
        ON DUPLICATE KEY UPDATE nomor_terakhir = nomor_terakhir + 1
    """
    cursor.execute(sql, (kode_unik,))
    conn.commit()
    conn.close()

# ================= FUNGSI FORMAT SURAT =================
def get_romawi(bulan):
    map_romawi = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
    return map_romawi[int(bulan)]

def format_tanggal_indo(tgl):
    bulan_indo = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", 
                  "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    return f"{tgl.day:02d} {bulan_indo[tgl.month]} {tgl.year}"

def get_hari_indo(tgl):
    days = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    return days[tgl.weekday()]

def generate_string_nomor(jenis_kop, kode_index, nomor_int, tgl_obj):
    nomor_str = f"{nomor_int:03d}"
    if jenis_kop == "PANITIA": nomor_str = f"{nomor_int:02d}"

    bulan_romawi = get_romawi(tgl_obj.month)
    tahun_full = tgl_obj.year
    tahun_2digit = str(tahun_full)[-2:]

    if jenis_kop == "IPNU":
        return f"{nomor_str}/{TINGKATAN}/{kode_index}/{PERIODE_IPNU}/{KODE_THN_IPNU}/{bulan_romawi}/{tahun_2digit}"
    elif jenis_kop == "IPPNU":
        return f"{nomor_str}/{TINGKATAN}/{kode_index}/{KODE_THN_IPPNU}/{PERIODE_IPPNU}/{bulan_romawi}/{tahun_full}"
    elif jenis_kop == "BERSAMA":
        return f"{nomor_str}/{TINGKATAN}/{kode_index}/{KODE_THN_IPNU}-{KODE_THN_IPPNU}/{PERIODE_IPNU}/{bulan_romawi}/{tahun_full}"
    elif jenis_kop == "PANITIA":
        return f"{nomor_str}/{TINGKATAN}/{kode_index}/IPNU-IPPNU/{PERIODE_IPNU}/{bulan_romawi}/{tahun_full}"

# ================= APLIKASI UTAMA =================
st.set_page_config(page_title="Sistem Surat IPNU IPPNU", page_icon="📝", layout="wide")

# Init DB
init_db()

st.title("📝 Generator Surat (TiDB Cloud)")
st.markdown("Sistem Administrasi Otomatis - PAC Kauman")

# --- SIDEBAR: LOGIC PEMILIHAN SURAT ---
with st.sidebar:
    st.header("⚙️ Konfigurasi Surat")
    
    # 1. Pilih Jenis KOP
    jenis_kop = st.selectbox("Pilih Kop Organisasi", ["BERSAMA", "PANITIA", "IPNU", "IPPNU"])
    
    # Mapping Template (Pastikan nama file SAMA PERSIS dengan yg diupload)
    # Sesuaikan nama file PANITIA dengan file yang kamu upload (Surat Kepanitiaan.docx atau TEMPLATE_PANITIA.docx)
    template_map = {
        "BERSAMA": "Surat Bersama (ABC).docx", 
        "IPNU": "TEMPLATE_IPNU.docx",
        "IPPNU": "TEMPLATE_IPPNU.docx",
        "PANITIA": "Surat Kepanitiaan.docx" 
    }
    file_template = template_map[jenis_kop]

    # 2. Logic Index Berdasarkan Kop
    kode_index_final = ""
    
    if jenis_kop == "PANITIA":
        # Untuk Panitia biasanya kodenya custom (Misal: Pan.Bersama, Pan.Konferancab)
        # Tapi jika mau A/B/C juga bisa diketik manual
        kode_index_final = st.text_input("Kode Kepanitiaan", value="Pan.Bersama")
        kategori_surat = "Panitia"
    else:
        # Pilihan Kategori (Umum vs Khusus)
        kategori_surat = st.radio("Kategori Surat", ["Surat Umum (A-C)", "Surat Khusus"])
        
        if kategori_surat == "Surat Khusus":
            # List Khusus Baru: SM, SPt, ST
            kode_index_final = st.selectbox("Pilih Kode Khusus", INDEX_KHUSUS)
        else:
            # Semua KOP (IPNU, IPPNU, BERSAMA) sekarang pakai A, B, C
            kode_index_final = st.selectbox("Pilih Index", INDEX_FULL)

    # 3. Preview Nomor
    kode_unik_db = f"{jenis_kop}_{kode_index_final}"
    nomor_terakhir_db = get_nomor_terakhir(kode_unik_db)
    nomor_calon = nomor_terakhir_db + 1
    tgl_sekarang = datetime.date.today()
    preview_str = generate_string_nomor(jenis_kop, kode_index_final, nomor_calon, tgl_sekarang)
    
    st.divider()
    st.info(f"🔢 Nomor Berikutnya:\n**{preview_str}**")

# --- HALAMAN UTAMA: FORM INPUT ---

# Tanggal Pembuatan Surat (Header Atas)
st.subheader("1. Tanggal Pembuatan Surat")
col_tgl1, col_tgl2 = st.columns(2)
with col_tgl1:
    tgl_buat_masehi = st.date_input("Tanggal Masehi", datetime.date.today())
    str_tgl_buat_masehi = format_tanggal_indo(tgl_buat_masehi) + " M"
with col_tgl2:
    str_tgl_buat_hijriyah = st.text_input("Tanggal Hijriyah", "16 Sya'ban 1447 H")

st.divider()

# --- FORM DINAMIS (BERUBAH SESUAI JENIS KOP) ---

if jenis_kop == "BERSAMA":
    st.subheader("2. Detail Surat Bersama")
    
    with st.form("form_bersama"):
        c1, c2 = st.columns(2)
        with c1:
            lampiran = st.text_input("Jumlah Lampiran", value="1 (Satu) Berkas")
            penerima = st.text_input("Penerima (Yth)", placeholder="Misal: Bapak Kepala Sekolah...")
        with c2:
            perihal = st.text_input("Perihal Surat", value="Permohonan Peminjaman Tempat")
            alamat_penerima = st.text_input("Alamat Penerima", placeholder="Misal: di- Tempat")

        st.markdown("---")
        st.caption("Bagian di bawah ini sudah terisi otomatis, tapi BISA DIEDIT jika perlu.")

        # BAGIAN PEMBUKA
        default_pembuka = "Dalam rangka menindaklanjuti hasil musyawarah Pimpinan Anak Cabang IPNU IPPNU Kecamatan Kauman, kami bermaksud mengadakan kegiatan RAPAT KERJA II PAC IPNU IPPNU Kauman pada :"
        pembuka = st.text_area("Kalimat Pembuka", value=default_pembuka, height=100)

        # BAGIAN DETAIL ACARA
        st.markdown("##### Detail Pelaksanaan:")
        ca, cb, cc = st.columns(3)
        with ca:
            tgl_acara = st.date_input("Tanggal Acara", datetime.date.today())
            hari_acara = get_hari_indo(tgl_acara)
            str_tgl_acara = format_tanggal_indo(tgl_acara)
        with cb:
            waktu_acara = st.text_input("Waktu", value="13.00 WIB - Selesai")
        with cc:
            tempat_acara = st.text_input("Tempat", value="SD Negeri 1 Kauman")

        # BAGIAN ISI
        default_isi = "Sehubungan dengan kegiatan tersebut, kami mengharapkan kehadiran segenap Pengurus PAC IPNU IPPNU Kauman masa khidmat 2025-2027 sebagai peserta aktif guna merumuskan serta menetapkan rencana program kerja organisasi selama satu tahun ke depan. Mengingat pentingnya agenda ini demi kesinambungan roda organisasi, kami memohon kehadiran Rekan dan Rekanita tepat pada waktunya."
        isi = st.text_area("Isi Paragraf Utama", value=default_isi, height=150)

        # BAGIAN PENUTUP
        default_penutup = "Demikian surat undangan ini kami buat, atas perhatian dan kehadirannya kami sampaikan terimakasih."
        penutup = st.text_area("Kalimat Penutup", value=default_penutup, height=80)

        submitted = st.form_submit_button("🚀 Generate Surat Bersama")

        if submitted:
            context = {
                'nomor_surat': generate_string_nomor(jenis_kop, kode_index_final, nomor_calon, tgl_buat_masehi),
                'jumlah_lampiran': lampiran,
                'perihal': perihal,
                'penerima': penerima,
                'alamat_penerima': alamat_penerima,
                'pembuka': pembuka,
                'hari_pelaksanaan': hari_acara,
                'tanggal_pelaksanaan': str_tgl_acara,
                'waktu_pelaksanaan': waktu_acara,
                'tempat_pelaksanaan': tempat_acara,
                'isi': isi,
                'penutup': penutup,
                'tanggal_pembuatan_surat_hijriyah': str_tgl_buat_hijriyah,
                'tanggal_pembuatan_surat_masehi': str_tgl_buat_masehi
            }

elif jenis_kop == "PANITIA":
    st.subheader("2. Detail Surat Kepanitiaan")
    
    with st.form("form_panitia"):
        c1, c2 = st.columns(2)
        with c1:
            nama_acara = st.text_input("Nama Acara", value="Rapat Kerja II PAC IPNU IPPNU Kauman")
            lampiran = st.text_input("Jumlah Lampiran", value="1 (Satu) Berkas")
            penerima = st.text_input("Penerima (Yth)", placeholder="Misal: Bapak Kepala Sekolah...")
            
        with c2:
            perihal = st.text_input("Perihal Surat", value="Permohonan Peminjaman Tempat")
            alamat_penerima = st.text_input("Alamat Penerima", placeholder="Misal: di- Tempat")
        
        st.divider()
        st.markdown("##### Penanggung Jawab (Tanda Tangan):")
        cp1, cp2 = st.columns(2)
        with cp1:
            nama_ketupel = st.text_input("Nama Ketua Pelaksana", placeholder="Nama Ketua Panitia")
        with cp2:
            nama_sekpel = st.text_input("Nama Sekretaris Pelaksana", placeholder="Nama Sekretaris Panitia")

        st.markdown("---")
        st.caption("Bagian di bawah ini sudah terisi otomatis, tapi BISA DIEDIT jika perlu.")

        # BAGIAN PEMBUKA
        default_pembuka = "Dalam rangka menindaklanjuti hasil musyawarah Pimpinan Anak Cabang IPNU IPPNU Kecamatan Kauman, kami bermaksud mengadakan kegiatan Rapat Kerja II PAC IPNU IPPNU Kauman pada :"
        pembuka = st.text_area("Kalimat Pembuka", value=default_pembuka, height=100)

        # BAGIAN DETAIL ACARA
        st.markdown("##### Detail Pelaksanaan:")
        ca, cb, cc = st.columns(3)
        with ca:
            tgl_acara = st.date_input("Tanggal Acara", datetime.date.today())
            hari_acara = get_hari_indo(tgl_acara)
            str_tgl_acara = format_tanggal_indo(tgl_acara)
        with cb:
            waktu_acara = st.text_input("Waktu", value="13.00 WIB - Selesai")
        with cc:
            tempat_acara = st.text_input("Tempat", value="SD Negeri 1 Kauman")

        # BAGIAN ISI
        default_isi = "Sehubungan dengan hal tersebut, kami memohon izin dan kesediaan Bapak untuk meminjamkan sarana dan prasarana sekolah guna mendukung kelancaran kegiatan tersebut. Adapun daftar unit yang kami perlukan telah kami rincikan pada lampiran surat ini."
        isi = st.text_area("Isi Paragraf Utama", value=default_isi, height=150)

        # BAGIAN PENUTUP
        default_penutup = "Demikian surat undangan ini kami buat, atas perhatian dan kehadirannya kami sampaikan terimakasih."
        penutup = st.text_area("Kalimat Penutup", value=default_penutup, height=80)

        submitted = st.form_submit_button("🚀 Generate Surat Panitia")

        if submitted:
            context = {
                'nomor_surat': generate_string_nomor(jenis_kop, kode_index_final, nomor_calon, tgl_buat_masehi),
                'jumlah_lampiran': lampiran,
                'perihal': perihal,
                'penerima': penerima,
                'alamat_penerima': alamat_penerima,
                'pembuka': pembuka,
                'hari_pelaksanaan': hari_acara,
                'tanggal_pelaksanaan': str_tgl_acara,
                'waktu_pelaksanaan': waktu_acara,
                'tempat_pelaksanaan': tempat_acara,
                'isi': isi,
                'penutup': penutup,
                'tanggal_pembuatan_surat_hijriyah': str_tgl_buat_hijriyah,
                'tanggal_pembuatan_surat_masehi': str_tgl_buat_masehi,
                'nama_acara': nama_acara,
                'nama_ketupel': nama_ketupel,
                'nama_sekpel': nama_sekpel
            }

else:
    # --- FORM UNTUK IPNU / IPPNU (TEMPLATE LAMA/LAIN) ---
    st.info("Anda memilih Kop IPNU/IPPNU. Form disesuaikan dengan template standar.")
    with st.form("form_standar"):
        nama_penerima = st.text_input("Nama Penerima")
        alamat_penerima = st.text_input("Alamat Penerima")
        nama_acara = st.text_input("Nama Acara")
        hari_tgl_acara = st.text_input("Hari, Tanggal Acara")
        waktu_acara = st.text_input("Waktu")
        tempat_acara = st.text_input("Tempat")
        
        submitted = st.form_submit_button("🚀 Generate Surat")

        if submitted:
            # Context disesuaikan dengan variabel di template lama
            context = {
                'nomor_surat': generate_string_nomor(jenis_kop, kode_index_final, nomor_calon, tgl_buat_masehi),
                'nama_penerima': nama_penerima,
                'alamat_penerima': alamat_penerima,
                'nama_acara': nama_acara,
                'hari_tanggal': hari_tgl_acara,
                'waktu': waktu_acara,
                'tempat': tempat_acara,
                'tanggal_surath': str_tgl_buat_hijriyah,
                'tanggal_suratm': str_tgl_buat_masehi
            }

# --- PROSES GENERATE & DOWNLOAD ---
if submitted:
    if not os.path.exists(file_template):
        st.error(f"❌ File template '{file_template}' tidak ditemukan! Harap upload file .docx yang sesuai.")
    else:
        try:
            doc = DocxTemplate(file_template)
            doc.render(context)
            
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("✅ Surat Berhasil Dibuat!")
            
            # Nama file output
            nama_penerima_file = context.get('penerima') or context.get('nama_penerima') or "TanpaNama"
            nama_file_clean = nama_penerima_file.replace("/", "_").replace("\\", "_")
            final_filename = f"Surat_{jenis_kop}_{kode_index_final}_{nama_file_clean}.docx"

            # Tombol Download + Trigger DB Increment
            st.download_button(
                label="⬇️ Download File Word",
                data=bio.getvalue(),
                file_name=final_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                on_click=increment_nomor,
                args=(kode_unik_db,)
            )
            
        except Exception as e:
            st.error(f"Terjadi kesalahan saat render template: {e}")