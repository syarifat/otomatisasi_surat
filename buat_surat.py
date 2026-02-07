from docxtpl import DocxTemplate
import os

def buat_surat():
    # Nama file template harus SAMA PERSIS dengan file yang kamu punya
    file_template = "template_surat.docx"

    # Cek apakah file ada
    if not os.path.exists(file_template):
        print(f"Error: File '{file_template}' tidak ditemukan di folder ini.")
        return

    # 1. Load Template
    doc = DocxTemplate(file_template)

    print("=== PROGRAM ISI SURAT OTOMATIS ===")
    print("Silakan masukkan data berikut:\n")

    # 2. Input Data User (Sesuai tanda {{...}} di Word)
    # Tips: Tekan Enter setelah mengetik
    
    nomor = input("1. Nomor Surat (misal: 01/PAC...): ")
    nama = input("2. Nama Penerima (misal: Bapak Kepala Sekolah...): ")
    alamat = input("3. Alamat Penerima (misal: Tempat): ")
    acara = input("4. Nama Acara (misal: RAPAT KERJA II...): ")
    hari = input("5. Hari/Tanggal Acara (misal: Jumat, 06 Februari 2026): ")
    jam = input("6. Waktu Acara (misal: 19.00 WIB): ")
    lokasi = input("7. Tempat Acara (misal: Aula Sekolah): ")
    tgl_surath = input("8. Tanggal Surat (misal: 26 Sya'ban 1447H): ")
    tgl_suratm = input("8. Tanggal Surat (misal: 06 Februari 2026): ")

    # 3. Menghubungkan inputan ke Placeholder Word
    context = {
        'nomor_surat': nomor,       # Sesuai [cite: 1]
        'nama_penerima': nama,      # Sesuai [cite: 5]
        'alamat_penerima': alamat,  # Sesuai [cite: 7]
        'nama_acara': acara,        # Sesuai [cite: 11]
        'hari_tanggal': hari,       # Sesuai [cite: 12]
        'waktu': jam,               # Sesuai [cite: 13]
        'tempat': lokasi,           # Sesuai [cite: 14]
        'tanggal_surath': tgl_surath,
        'tanggal_suratm': tgl_suratm
    }

    # 4. Render (Mengisi template)
    doc.render(context)

    # 5. Simpan Hasil
    # Nama file output dibuat unik berdasarkan nama penerima
    nama_file_jadi = f"Surat_Peminjaman_untuk_{nama.replace(' ', '_')}.docx"
    doc.save(nama_file_jadi)

    print("\n" + "="*30)
    print(f"BERHASIL! File surat sudah jadi: {nama_file_jadi}")
    print("="*30)

if __name__ == "__main__":
    buat_surat()