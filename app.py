import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

# =====================================================================
# FUNGSI UTAMA: MENYUSUN DATA MENJADI TABEL WORD YANG RAPI
# =====================================================================
def buat_dokumen_rpm(data):
    doc = Document()
    
    # Pengaturan Margin Halaman Standard (1 Inci / 2.54 cm)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Pengaturan Font Default Dokumen
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)

    # 1. Judul Utama Dokumen
    title = doc.add_paragraph()
    title_run = title.add_run("RENCANA PEMBELAJARAN MENDALAM (RPM)")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title.alignment = 1  # Rata Tengah (Center)
    doc.add_paragraph()  # Jarak baris kosong

    # 2. Bagian I: Identitas & Validasi (Format Tabel)
    doc.add_heading("I. IDENTITAS DAN VALIDASI", level=2)
    table_identitas = doc.add_table(rows=7, cols=2)
    table_identitas.style = 'Table Grid'
    
    identitas_labels = [
        ("Nama Sekolah", data['sekolah']),
        ("Nama Guru", data['guru']),
        ("Mata Pelajaran", data['mapel']),
        ("Kelas / Semester", data['kelas_semester']),
        ("Alokasi Waktu", data['alokasi_waktu']),
        ("Topik Utama", data['topik']),
        ("Capaian Pembelajaran (CP)", data['cp'])
    ]
    
    for i, (label, value) in enumerate(identitas_labels):
        row = table_identitas.rows[i]
        row.cells[0].text = label
        row.cells[1].text = value
        # Menebalkan teks label di kolom pertama
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        
    doc.add_paragraph()

    # 3. Bagian II: 8 Struktur Utama Rencana Pembelajaran Mendalam
    doc.add_heading("II. KOMPONEN INTI RPM MENDALAM", level=2)
    table_inti = doc.add_table(rows=9, cols=2)
    table_inti.style = 'Table Grid'
    
    # Membuat Header Tabel Inti
    hdr_cells = table_inti.rows[0].cells
    hdr_cells[0].text = 'Komponen RPM'
    hdr_cells[1].text = 'Deskripsi / Detail Rencana Kerja'
    hdr_cells[0].paragraphs[0].runs[0].font.bold = True
    hdr_cells[1].paragraphs[0].runs[0].font.bold = True
    
    # Mewarnai Background Header Tabel (Abu-abu Terang)
    shading_1 = parse_xml(r'<w:shd {} w:fill="E6E6E6"/>'.format(nsdecls('w')))
    shading_2 = parse_xml(r'<w:shd {} w:fill="E6E6E6"/>'.format(nsdecls('w')))
    hdr_cells[0]._tc.get_or_add_tcPr().append(shading_1)
    hdr_cells[1]._tc.get_or_add_tcPr().append(shading_2)

    # Menggabungkan data input ke dalam format baris tabel
    komponen_data = [
        ("1. Dimensi Profil Lulusan", data['dimensi_profil']),
        ("2. Tujuan Pembelajaran", data['tujuan_pembelajaran']),
        ("3. Praktik Pedagogis", data['praktik_pedagogis']),
        ("4. Lingkungan Pembelajaran", data['lingkungan_belajar']),
        ("5. Kemitraan Pembelajaran", data['kemitraan_belajar']),
        ("6. Pemanfaatan Digital", data['pemanfaatan_digital']),
        ("7. Langkah Pembelajaran", data['langkah_pembelajaran']),
        ("8. Asesmen & Lembar Kerja", f"Metode Evaluasi (Formatif & Sumatif):\n{data['asesmen_metode']}\n\nKriteria / Rubrik Penilaian:\n{data['asesmen_rubrik']}\n\nLembar Kerja Murid (LKM):\n{data['asesmen_lkm']}")
    ]

    for i, (komponen, isi) in enumerate(komponen_data):
        row = table_inti.rows[i+1]
        row.cells[0].text = komponen
        row.cells[1].text = isi
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph()
    doc.add_paragraph()

    # 4. Bagian III: Pengesahan Oleh Kepala Sekolah & Guru
    doc.add_heading("III. PENGESAHAN", level=2)
    table_ttd = doc.add_table(rows=1, cols=2)
    
    # Menghapus garis tepi tabel tanda tangan agar terlihat bersih seperti dokumen resmi
    for row in table_ttd.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders {}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w')))
            tcPr.append(tcBorders)

    cell_kiri = table_ttd.rows[0].cells[0].paragraphs[0]
    cell_kiri.add_run(f"Mengetahui,\nKepala Sekolah {data['sekolah']}\n\n\n\n\n( _______________________ )")
    
    cell_kanan = table_ttd.rows[0].cells[1].paragraphs[0]
    cell_kanan.add_run(f"Guru Mata Pelajaran,\n\n\n\n\n\n( {data['guru']} )")

    # Menyimpan dokumen ke dalam memori sementara agar bisa langsung diunduh
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# =====================================================================
# ANTARMUKA WEB APLIKASI (STREAMLIT INTERFACE)
# =====================================================================
st.set_page_config(page_title="Aplikasi Pembuat RPM", layout="wide")

st.title("📝 Aplikasi Pembuat Rencana Pembelajaran Mendalam (RPM)")
st.write("Isi formulir di bawah ini untuk menyusun dokumen pembelajaran terstruktur, lalu ekspor langsung ke file Word (.docx) dengan format tabel yang rapi.")

# Membagi layout input menjadi dua kolom utama agar seimbang di layar
col1, col2 = st.columns(2)

with col1:
    st.subheader("I. Identitas RPM")
    sekolah = st.text_input("Nama Sekolah", "SMA Negeri 1 Pembelajaran")
    guru = st.text_input("Nama Guru", "Nama Guru, S.Pd.")
    mapel = st.text_input("Mata Pelajaran", "Bahasa Indonesia / Biologi / Informatika")
    kelas_semester = st.text_input("Kelas / Semester", "XI / Ganjil")
    alokasi_waktu = st.text_input("Alokasi Waktu", "2 x 45 Menit")
    topik = st.text_input("Topik Pembelajaran", "Eksplorasi Gagasan Kreatif")
    cp = st.text_area("Capaian Pembelajaran (CP)", "Peserta didik mampu menganalisis, mengevaluasi, dan menciptakan solusi kreatif berdasarkan materi yang diajarkan...")

with col2:
    st.subheader("II. Komponen Inti (8 Pilar RPM)")
    dimensi_profil = st.text_area("1. Dimensi Profil Lulusan", "Bernalar Kritis, Kreatif, dan Mandiri (Disesuaikan dengan keterampilan yang dikembangkan).")
    tujuan_pembelajaran = st.text_area("2. Tujuan Pembelajaran (Berkesadaran, Bermakna, Menggembirakan)", "Melalui diskusi pemecahan masalah, siswa secara sadar memahami kebermaknaan materi dalam kehidupan nyata dengan atmosfer kelas yang menggembirakan.")
    praktik_pedagogis = st.text_area("3. Praktik Pedagogis (Pendekatan Berbasis Masalah)", "Menggunakan Problem-Based Learning (PBL). Guru memberikan pemantik masalah nyata, kemudian siswa bekerja kelompok untuk merumuskan solusinya.")
    lingkungan_belajar = st.text_area("4. Lingkungan Pembelajaran", "Fisik: Susunan meja fleksibel/berkelompok. Budaya Belajar: Saling menghargai argumen, aman untuk melakukan kesalahan, dan refleksi terbuka.")

st.markdown("---")
col3, col4 = st.columns(2)

with col3:
    kemitraan_belajar = st.text_area("5. Kemitraan Pembelajaran", "Kolaborasi aktif antar peserta didik di dalam kelompok, guru bertindak sebagai fasilitator mitra belajar, dan terintegrasi dengan teknologi.")
    pemanfaatan_digital = st.text_area("6. Pemanfaatan Digital", "Menggunakan platform interaktif (misalnya Canva, Google Docs, Jamboard) untuk kolaborasi real-time dan pencarian referensi digital.")

with col4:
    langkah_pembelajaran = st.text_area("7. Langkah Pembelajaran (Tahapan Proses)", "1. Pendahuluan (15 Menit): Orientasi, apersepsi bermakna, penyampaian tujuan.\n2. Kegiatan Inti (60 Menit): Orientasi masalah, investigasi kelompok, presentasi karya.\n3. Penutup (15 Menit): Refleksi bersama, kesimpulan, dan rencana tindak lanjut.")

st.subheader("III. Komponen Asesmen Pembelajaran")
col5, col6, col7 = st.columns(3)

with col5:
    asesmen_metode = st.text_area("Metode Evaluasi (Formatif & Sumatif)", "Formatif: Penilaian jurnal refleksi, observasi partisipasi diskusi kelompok.\nSumatif: Penilaian produk akhir solusi masalah.")
with col6:
    asesmen_rubrik = st.text_area("Kriteria Penilaian / Rubrik", "Skor 4: Solusi sangat tajam dan aplikatif.\nSkor 3: Solusi logis tetapi standar.\nSkor 2: Solusi kurang menjawab masalah.\nSkor 1: Tidak menyertakan solusi.")
with col7:
    asesmen_lkm = st.text_area("Lembar Kerja Murid (LKM)", "1. Analisis apa akar masalah dari studi kasus tadi?\n2. Rumuskan 2 solusi terbaik kelompok Anda beserta alasannya!")

# --- PROSES PACKAGING DATA & UNDUH ---
st.markdown("---")
rpm_data = {
    'sekolah': sekolah, 'guru': guru, 'mapel': mapel, 'kelas_semester': kelas_semester,
    'alokasi_waktu': alokasi_waktu, 'topik': topik, 'cp': cp, 'dimensi_profil': dimensi_profil,
    'tujuan_pembelajaran': tujuan_pembelajaran, 'praktik_pedagogis': praktik_pedagogis,
    'lingkungan_belajar': lingkungan_belajar, 'kemitraan_belajar': kemitraan_belajar,
    'pemanfaatan_digital': pemanfaatan_digital, 'langkah_pembelajaran': langkah_pembelajaran,
    'asesmen_metode': asesmen_metode, 'asesmen_rubrik': asesmen_rubrik, 'asesmen_lkm': asesmen_lkm
}

# Membuat file dalam bentuk byte stream
file_word_ready = buat_dokumen_rpm(rpm_data)

st.write("Silakan klik tombol di bawah untuk mengunduh dokumen Word yang rapi:")
st.download_button(
    label="📥 Unduh Dokumen RPM (.docx)",
    data=file_word_ready,
    file_name=f"RPM_{topik.replace(' ', '_')}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)