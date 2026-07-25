import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

# =====================================================================
# FUNGSI INTEGRASI AI GOOGLE (GEMINI) - VERSI AMAN TANPA CRASH
# =====================================================================
def panggil_ai_guru(topik, cp, komponen_rpp, instruksi_khusus):
    api_key_ai = st.secrets.get("GEMINI_API_KEY", None)
    
    if not api_key_ai:
        return "⚠️ Eror: Kunci API AI belum dikonfigurasi di pengaturan Secrets Streamlit Anda."
    
    try:
        import requests
        import json
        
        url = f"https://googleapis.com{api_key_ai}"
        headers = {'Content-Type': 'application/json'}
        
        prompt = f"Topik: {topik}\nCP: {cp}\nKomponen: {komponen_rpp}\nInstruksi: {instruksi_khusus}"
        
        payload = {
            "contents": [{
                "parts": [{"text": prompt}]
            }]
        }
        
        response = requests.post(url, headers=headers, data=json.dumps(payload))
        res_json = response.json()
        
        return res_json['candidates']['content']['parts']['text']
    except Exception as e:
        return f"⚠️ Gagal memuat AI secara otomatis. Anda dapat mengisi kolom ini secara manual. (Detail Eror: {str(e)})"

# =====================================================================
# FUNGSI UTAMA: MENYUSUN DATA MENJADI TABEL WORD YANG RAPI
# =====================================================================
def buat_dokumen_rpm(data):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)

    title = doc.add_paragraph()
    title_run = title.add_run("RENCANA PEMBELAJARAN MENDALAM (RPM)")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title.alignment = 1
    doc.add_paragraph()

    doc.add_heading("I. IDENTITAS DAN VALIDASI", level=2)
    table_identitas = doc.add_table(rows=7, cols=2)
    table_identitas.style = 'Table Grid'
    
    identitas_labels = [
        ("Nama Sekolah", data.get('sekolah', '')), 
        ("Nama Guru", data.get('guru', '')),
        ("Mata Pelajaran", data.get('mapel', '')), 
        ("Kelas / Semester", data.get('kelas_semester', '')),
        ("Alokasi Waktu", data.get('alokasi_waktu', '')), 
        ("Topik Utama", data.get('topik', '')),
        ("Capaian Pembelajaran (CP)", data.get('cp', ''))
    ]
    
    for i, (label, value) in enumerate(identitas_labels):
        row = table_identitas.rows[i]
        row.cells.text = str(label)
        row.cells.text = str(value)
        row.cells.paragraphs.runs.font.bold = True
        
    doc.add_paragraph()

    doc.add_heading("II. KOMPONEN INTI RPM MENDALAM", level=2)
    table_inti = doc.add_table(rows=9, cols=2)
    table_inti.style = 'Table Grid'
    
    hdr_cells = table_inti.rows.cells
    hdr_cells.text = 'Komponen RPM'
    hdr_cells.text = 'Deskripsi / Detail Rencana Kerja (Hasil AI & Guru)'
    hdr_cells.paragraphs.runs.font.bold = True
    hdr_cells.paragraphs.runs.font.bold = True
    
    shading_1 = parse_xml(r'<w:shd {} w:fill="E6E6E6"/>'.format(nsdecls('w')))
    shading_2 = parse_xml(r'<w:shd {} w:fill="E6E6E6"/>'.format(nsdecls('w')))
    hdr_cells._tc.get_or_add_tcPr().append(shading_1)
    hdr_cells._tc.get_or_add_tcPr().append(shading_2)

    komponen_data = [
        ("1. Dimensi Profil Lulusan & Tujuan", data.get('dimensi_profil', '')),
        ("2. Tujuan Pembelajaran", "Terintegrasi pada kolom nomor 1 di atas"),
        ("3. Praktik Pedagogis", data.get('praktik_pedagogis', '')),
        ("4. Lingkungan Pembelajaran", data.get('lingkungan_belajar', '')),
        ("5. Kemitraan Pembelajaran", data.get('kemitraan_belajar', '')),
        ("6. Pemanfaatan Digital", data.get('pemanfaatan_digital', '')),
        ("7. Langkah Pembelajaran Rinci", data.get('langkah_pembelajaran', '')),
        ("8. Asesmen & Lembar Kerja", data.get('asesmen_total', ''))
    ]

    for i, (komponen, isi) in enumerate(komponen_data):
        row = table_inti.rows[i+1]
        row.cells.text = str(komponen)
        row.cells.text = str(isi)
        row.cells.paragraphs.runs.font.bold = True

    doc.add_paragraph()
    doc.add_paragraph()

    doc.add_heading("III. PENGESAHAN", level=2)
    table_ttd = doc.add_table(rows=1, cols=2)
    for row in table_ttd.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'<w:tcBorders {}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w')))
            tcPr.append(tcBorders)

    cell_kiri = table_ttd.rows.cells.paragraphs
    cell_kiri.add_run(f"Mengetahui,\nKepala Sekolah {data.get('sekolah', '')}\n\n\n\n\n( _______________________ )")
    
    cell_kanan = table_ttd.rows.cells.paragraphs
    cell_kanan.add_run(f"Guru Mata Pelajaran,\n\n\n\n\n\n( {data.get('guru', '')} )")

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

if "tujuan_ai" not in st.session_state: st.session_state.tujuan_ai = ""
if "langkah_ai" not in st.session_state: st.session_state.langkah_ai = ""
if "asesmen_ai" not in st.session_state: st.session_state.asesmen_ai = ""

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
    st.info("💡 Ketik Topik & CP di sebelah kiri terlebih dahulu, lalu klik tombol AI di bawah ini untuk memicu pencarian ide mendalam.")
    
    if st.button("✨ Detailing Tujuan Pembelajaran & Profil (AI)"):
        with st.spinner("AI sedang merumuskan tujuan mendalam..."):
            instruksi = "Buat dimensi Profil Lulusan Abad 21 dan Tujuan Pembelajaran yang berkesadaran, bermakna, dan menggembirakan."
            st.session_state.tujuan_ai = panggil_ai_guru(topik, cp, "Tujuan & Profil", instruksi)
            st.rerun()
            
    if st.button("🔥 Kembangkan Kegiatan Pembelajaran Mendalam (AI)"):
        with st.spinner("AI sedang merancang langkah kerja rinci..."):
            instruksi = "Buat langkah pembelajaran berbasis masalah (PBL) sangat detail mencakup menit dan aksi nyata guru murid: Pembukaan, Inti (Penyelidikan & pemanfaatan teknologi digital), dan Penutup (Refleksi)."
            st.session_state.langkah_ai = panggil_ai_guru(topik, cp, "Langkah Kerja", instruksi)
            st.rerun()
            
    if st.button("📊 Kembangkan Kriteria Asesmen & LKM Lengkap (AI)"):
        with st.spinner("AI sedang menyusun lembar kerja murid..."):
            instruksi = "Buat instrumen asesmen lengkap: 1) Metode Formatif Sumatif. 2) Lembar Kerja Murid (LKM) studi kasus riil. 3) Rubrik penilaian skor 1-4."
            st.session_state.asesmen_ai = panggil_ai_guru(topik, cp, "Asesmen & LKM", instruksi)
            st.rerun()

st.markdown("---")
st.subheader("III. Peninjauan & Penyempurnaan Teks (Dapat Diedit Manual)")

val_tujuan = st.session_state.tujuan_ai if st.session_state.tujuan_ai else "Klik tombol AI di atas untuk mengisi otomatis..."
val_langkah = st.session_state.langkah_ai if st.session_state.langkah_ai else "Klik tombol AI di atas untuk menyusun kegiatan mendalam..."
val_asesmen = st.session_state.asesmen_ai if st.session_state.asesmen_ai else "Klik tombol AI di atas untuk membuat instrumen evaluasi..."

dimensi_profil = st.text_area("1 & 2. Dimensi Profil & Tujuan Pembelajaran (Hasil AI)", val_tujuan, height=150)
praktik_pedagogis = st.text_area("3. Praktik Pedagogis", "Menggunakan pendekatan Problem-Based Learning (PBL) berbasis penyelidikan kasus nyata.")
lingkungan_belajar = st.text_area("4. Lingkungan Pembelajaran", "Fisik: Meja berkelompok fleksibel. Budaya: Kolaboratif, ramah kesalahan, aman berpendapat.")
kemitraan_belajar = st.text_area("5. Kemitraan Pembelajaran", "Kolaborasi aktif antar peserta didik, guru sebagai fasilitator, dan pemanfaatan gawai cerdas.")
pemanfaatan_digital = st.text_area("6. Pemanfaatan Digital", "Platform kolaborasi online (Google Workspace/Canva/Padlet) untuk pengerjaan kelompok.")

langkah_pembelajaran = st.text_area("7. Langkah Pembelajaran Rinci (Hasil Pengembangan AI)", val_langkah, height=200)
asesmen_total = st.text_area("8. Asesmen, Rubrik & Lembar Kerja Murid (Hasil Pengembangan AI)", val_asesmen, height=200)

# Pemaketan data yang sudah dibersihkan
rpm_data = {
    'sekolah': sekolah,
    'guru': guru,
    'mapel': mapel,
    'kelas_semester': kelas_semester,
