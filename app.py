import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from google import genai
import io

# =====================================================================
# FUNGSI INTEGRASI AI GOOGLE (GEMINI)
# =====================================================================
def panggil_ai_guru(topik, cp, komponen_rpp, instruksi_khusus):
    # Mengambil API Key yang disimpan aman di Streamlit Secrets
    api_key_ai = st.secrets.get("GEMINI_API_KEY", None)
    
    if not api_key_ai:
        return "⚠️ Eror: Kunci API AI belum dikonfigurasi di pengaturan server Streamlit Anda."
    
    try:
        client = genai.Client(api_key=api_key_ai)
        
        prompt = f"""
        Anda adalah pakar kurikulum pendidikan modern dan perancang Rencana Pembelajaran Mendalam (RPM).
        Tugas Anda adalah mengembangkan bagian '{komponen_rpp}' secara sangat rinci, mendalam, aplikatif, 
        dan siap pakai sebagai referensi utama guru di kelas.
        
        Informasi Kelas:
        - Topik Pembelajaran: {topik}
        - Capaian Pembelajaran (CP): {cp}
        
        Instruksi Khusus untuk Komponen Ini:
        {instruksi_khusus}
        
        Berikan jawaban yang padat, aplikatif, tuntas, tanpa basa-basi pembuka yang tidak penting.
        """
        
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        return response.text
    except Exception as e:
        return f"⚠️ Terjadi kesalahan saat menghubungi AI: {str(e)}"

# =====================================================================
# FUNGSI UTAMA: MENYUSUN DATA MENJADI TABEL WORD YANG RAPI
# =====================================================================
def buat_dokumen_rpm(data):
    doc = Document()
    
    # Pengaturan Margin Halaman Standard (1 Inci)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)

    # Judul Utama Dokumen
    title = doc.add_paragraph()
    title_run = title.add_run("RENCANA PEMBELAJARAN MENDALAM (RPM)")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title.alignment = 1
    doc.add_paragraph()

    # Bagian I: Identitas & Validasi
    doc.add_heading("I. IDENTITAS DAN VALIDASI", level=2)
    table_identitas = doc.add_table(rows=7, cols=2)
    table_identitas.style = 'Table Grid'
    
    identitas_labels = [
        ("Nama Sekolah", data['sekolah']), ("Nama Guru", data['guru']),
        ("Mata Pelajaran", data['mapel']), ("Kelas / Semester", data['kelas_semester']),
        ("Alokasi Waktu", data['alokasi_waktu']), ("Topik Utama", data['topik']),
        ("Capaian Pembelajaran (CP)", data['cp'])
    ]
    
    for i, (label, value) in enumerate(identitas_labels):
        row = table_identitas.rows[i]
        row.cells[0].text = label
        row.cells[1].text = value
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        
    doc.add_paragraph()

    # Bagian II: 8 Struktur Utama Rencana Pembelajaran Mendalam
    doc.add_heading("II. KOMKOMPONEN INTI RPM MENDALAM", level=2)
    table_inti = doc.add_table(rows=9, cols=2)
    table_inti.style = 'Table Grid'
    
    hdr_cells = table_inti.rows[0].cells
    hdr_cells[0].text = 'Komponen RPM'
    hdr_cells[1].text = 'Deskripsi / Detail Rencana Kerja (Hasil AI & Guru)'
    hdr_cells[0].paragraphs[0].runs[0].font.bold = True
    hdr_cells[1].paragraphs[0].runs[0].font.bold = True
    
    shading_1 = parse_xml(r'<w:shd {} w:fill="E6E6E6"/>'.format(nsdecls('w')))
    shading_2 = parse_xml(r'<w:shd {} w:fill="E6E6E6"/>'.format(nsdecls('w')))
    hdr_cells[0]._tc.get_or_add_tcPr().append(shading_1)
    hdr_cells[1]._tc.get_or_add_tcPr().append(shading_2)

    komponen_data = [
        ("1. Dimensi Profil Lulusan", data['dimensi_profil']),
        ("2. Tujuan Pembelajaran", data['tujuan_pembelajaran']),
        ("3. Praktik Pedagogis", data['praktik_pedagogis']),
        ("4. Lingkungan Pembelajaran", data['lingkungan_belajar']),
        ("5. Kemitraan Pembelajaran", data['kemitraan_belajar']),
        ("6. Pemanfaatan Digital", data['pemanfaatan_digital']),
        ("7. Langkah Pembelajaran", data['langkah_pembelajaran']),
        ("8. Asesmen & Lembar Kerja", data['asesmen_total'])
    ]

    for i, (komponen, isi) in enumerate(komponen_data):
        row = table_inti.rows[i+1]
        row.cells[0].text = komponen
        row.cells[1].text = isi
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph()
    doc.add_paragraph()

    # Bagian III: Pengesahan
    doc.add_heading("III. PENGESAHAN", level=2)
    table_ttd = doc.add_table(rows=1, cols=2)
    for row in table_ttd.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders {}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w')))
            tcPr.append(tcBorders)

    cell_kiri = table_ttd.rows[0].cells[0].paragraphs[0]
    cell_kiri.add_run(f"Mengetahui,\nKepala Sekolah {data['sekolah']}\n\n\n\n\n( _______________________ )")
    
    cell_kanan = table_ttd.rows[0].cells[1].paragraphs[0]
    cell_kanan.add_run(f"Guru Mata Pelajaran,\n\n\n\n\n\n( {data['guru']} )")

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# =====================================================================
# ANTARMUKA WEB APLIKASI (STREAMLIT INTERFACE)
# =====================================================================
st.set_page_config(page_title="Aplikasi Pembuat RPM Cerdas", layout="wide")

st.title("🤖 Aplikasi Pembuat Rencana Pembelajaran Mendalam (RPM) Berbasis AI")
st.write("Isi identitas, lalu gunakan bantuan AI untuk mengembangkan langkah kegiatan dan asesmen yang rinci dan mendalam.")

# State untuk menyimpan hasil AI agar tidak hilang saat halaman reload
if "langkah_ai" not in st.session_state: st.session_state.langkah_ai = ""
if "asesmen_ai" not in st.session_state: st.session_state.asesmen_ai = ""
if "tujuan_ai" not in st.session_state: st.session_state.tujuan_ai = ""

col1, col2 = st.columns(2)

with col1:
    st.subheader("I. Identitas Dasar")
    sekolah = st.text_input("Nama Sekolah", "SMA Negeri 1 Pembelajaran")
    guru = st.text_input("Nama Guru", "Nama Guru, S.Pd.")
    mapel = st.text_input("Mata Pelajaran", "Informatika / Biologi")
    kelas_semester = st.text_input("Kelas / Semester", "XI / Ganjil")
    alokasi_waktu = st.text_input("Alokasi Waktu", "2 x 45 Menit")
    topik = st.text_input("Topik Pembelajaran", "Mitigasi Perubahan Iklim")
    cp = st.text_area("Capaian Pembelajaran (CP)", "Peserta didik mampu menganalisis fenomena perubahan iklim global, mengevaluasi dampaknya terhadap ekosistem lokal, dan menciptakan solusi praktis berbasis komunitas.")

with col2:
    st.subheader("II. Tombol Generator Cerdas AI")
    st.info("💡 Ketik Topik & CP di sebelah kiri terlebih dahulu, lalu klik tombol AI di bawah ini untuk menghasilkan rencana kerja yang mendalam.")
    
    if st.button("✨ Detailing Tujuan Pembelajaran & Profil (AI)"):
        with st.spinner("AI sedang merumuskan tujuan yang bermakna..."):
            instruksi = "Buat 2 bagian: 1) Dimensi Profil Lulusan Abad 21 yang dikembangkan. 2) Tujuan Pembelajaran yang berkesadaran, bermakna bagi kehidupan, dan dirancang dengan atmosfer kelas yang menggembirakan."
            st.session_state.tujuan_ai = panggil_ai_guru(topik, cp, "Tujuan Pembelajaran & Profil", instruksi)
            
    if st.button("🔥 Kembangkan Kegiatan Pembelajaran Mendalam (AI)"):
        with st.spinner("AI sedang merancang langkah pedagogis berbasis masalah secara detail..."):
            instruksi = "Buat langkah pembelajaran berbasis masalah (Problem-Based Learning) yang sangat detail mencakup menit dan aksi nyata guru serta murid: Pembukaan (Apersepsi mendalam), Kegiatan Inti (Penyelidikan masalah, diskusi kolaboratif kelompok, pemanfaatan alat digital), dan Penutup (Refleksi emosional dan kognitif)."
            st.session_state.langkah_ai = panggil_ai_guru(topik, cp, "Langkah Pembelajaran Rinci", instruksi)
            
    if st.button("📊 Kembangkan Kriteria Asesmen & LKM Lengkap (AI)"):
        with st.spinner("AI sedang menyusun lembar kerja murid dan rubrik..."):
            instruksi = "Buat instrumen asesmen lengkap yang terdiri dari: 1) Metode Evaluasi Formatif & Sumatif. 2) Lembar Kerja Murid (LKM) berisi studi kasus riil dan pertanyaan pemantik logika tinggi. 3) Kriteria / Rubrik penilaian performa kelompok secara terperinci (Skor 1-4 beserta indikatornya)."
            st.session_state.asesmen_ai = panggil_ai_guru(topik, cp, "Asesmen & LKM", instruksi)

st.markdown("---")
st.subheader("III. Peninjauan & Penyempurnaan Teks (Dapat Diedit Manual)")

dimensi_profil = st.text_area("1 & 2. Dimensi Profil & Tujuan Pembelajaran", st.session_state.tujuan_ai if st.session_state.tujuan_ai else "Klik tombol AI di atas untuk mengisi otomatis...")
praktik_pedagogis = st.text_area("3. Praktik Pedagogis", "Menggunakan pendekatan Problem-Based Learning (PBL) berbasis penyelidikan kasus nyata.")
lingkungan_belajar = st.text_area("4. Lingkungan Pembelajaran", "Fisik: Meja berkelompok fleksibel. Budaya: Kolaboratif, ramah kesalahan, aman berpendapat.")
kemitraan_belajar = st.text_area("5. Kemitraan Pembelajaran", "Kolaborasi aktif antar peserta didik, guru sebagai fasilitator, dan pemanfaatan gawai cerdas.")
pemanfaatan_digital = st.text_area("6. Pemanfaatan Digital", "Platform kolaborasi online (Google Workspace/Canva/Padlet) untuk pengerjaan kelompok.")