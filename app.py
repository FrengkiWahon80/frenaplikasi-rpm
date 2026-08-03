# =====================================================
# RPM CERDAS AI OFFLINE
# Generator Rencana Pembelajaran Mendalam
# Kurikulum Merdeka SMP
# =====================================================


import io
import streamlit as st

from docx import Document
from docx.shared import Inches, Pt



# =====================================================
# KONFIGURASI APLIKASI
# =====================================================


APP_NAME = "RPM CERDAS AI"



st.set_page_config(

    page_title=APP_NAME,

    page_icon="📘",

    layout="wide"

)



# =====================================================
# HEADER APLIKASI
# =====================================================


st.markdown(

    """
    <h1 style="text-align:center;">
    📘 RPM CERDAS AI
    </h1>

    <h4 style="text-align:center;">
    Generator Rencana Pembelajaran Mendalam
    Kurikulum Merdeka SMP
    </h4>

    """,

    unsafe_allow_html=True

)



st.divider()



# =====================================================
# IDENTITAS GURU
# =====================================================


st.subheader(

    "🏫 Identitas Pembelajaran"

)



kolom1, kolom2 = st.columns(2)



with kolom1:


    sekolah = st.text_input(

        "Nama Sekolah",

        value="SMP Negeri"

    )


    guru = st.text_input(

        "Nama Guru"

    )


    mapel = st.selectbox(

        "Mata Pelajaran",

        [

            "Pendidikan Agama",

            "PPKn",

            "Bahasa Indonesia",

            "Matematika",

            "IPA",

            "IPS",

            "Bahasa Inggris",

            "Seni Budaya",

            "PJOK",

            "Informatika",

            "Prakarya"

        ]

    )


    kelas = st.selectbox(

        "Kelas",

        [

            "VII",

            "VIII",

            "IX"

        ]

    )




with kolom2:


    semester = st.selectbox(

        "Semester",

        [

            "Ganjil",

            "Genap"

        ]

    )


    tahun_pelajaran = st.text_input(

        "Tahun Pelajaran",

        value="2026/2027"

    )


    alokasi_waktu = st.text_input(

        "Alokasi Waktu",

        value="2 x 40 Menit"

    )


    model_pembelajaran = st.selectbox(

        "Model Pembelajaran",

        [

            "Problem Based Learning",

            "Project Based Learning",

            "Discovery Learning",

            "Inquiry Learning",

            "Cooperative Learning"

        ]

    )



st.divider()



# =====================================================
# DATA MATERI
# =====================================================


st.subheader(

    "📚 Materi Pembelajaran"

)



topik = st.text_input(

    "Topik Pembelajaran"

)



sub_topik = st.text_input(

    "Sub Topik (Opsional)"

)



cp = st.text_area(

    "Capaian Pembelajaran (CP)",

    height=140,

    placeholder=

    "Masukkan CP Kurikulum Merdeka..."

)



karakteristik = st.text_area(

    "Karakteristik Peserta Didik (Opsional)",

    height=100

)



tujuan_manual = st.text_area(

    "Tujuan Pembelajaran (Opsional)",

    height=120

)



st.info(

    """
    💡 Versi offline menggunakan template cerdas.
    Tidak membutuhkan API Key atau koneksi AI.
    """

)



st.divider()
# =====================================================
# MESIN GENERATOR RPM OFFLINE
# =====================================================


def buat_rpm_offline(

    mapel,

    kelas,

    topik,

    cp,

    model,

    tujuan_manual

):


    # ================================================
    # TUJUAN PEMBELAJARAN
    # ================================================


    if tujuan_manual.strip():


        tujuan = tujuan_manual


    else:


        tujuan = f"""
Peserta didik mampu memahami konsep {topik}
pada mata pelajaran {mapel}, mengembangkan
kemampuan berpikir kritis, kreatif, kolaboratif,
serta mampu menerapkan pengetahuan dalam
kehidupan sehari-hari.
"""




    # ================================================
    # DIMENSI PROFIL LULUSAN
    # ================================================


    dimensi = f"""
Peserta didik mengembangkan:

1. Bernalar kritis melalui analisis materi {topik}.

2. Kreatif dalam menemukan solusi dan menghasilkan karya.

3. Mandiri dalam proses belajar.

4. Bergotong royong melalui kerja kelompok.

5. Komunikatif dalam menyampaikan ide dan hasil belajar.

Penguatan karakter disesuaikan dengan pembelajaran
{mapel} kelas {kelas}.
"""




    # ================================================
    # PRAKTIK PEDAGOGIS
    # ================================================


    pedagogis = f"""
Model pembelajaran yang digunakan:

{model}


Guru menerapkan pembelajaran berbasis masalah
dengan langkah:

1. Guru memberikan permasalahan kontekstual
   berkaitan dengan {topik}.

2. Peserta didik mengidentifikasi masalah.

3. Peserta didik mencari informasi dan berdiskusi.

4. Peserta didik menyampaikan hasil pemecahan masalah.

5. Guru memberikan penguatan dan refleksi.
"""




    # ================================================
    # LINGKUNGAN PEMBELAJARAN
    # ================================================


    lingkungan = f"""
Lingkungan pembelajaran dirancang agar:

- kelas aman, nyaman, dan menyenangkan.
- peserta didik aktif bertanya dan berdiskusi.
- terdapat budaya saling menghargai.
- tersedia sumber belajar yang mendukung materi {topik}.
- guru menciptakan suasana pembelajaran mendalam.
"""




    # ================================================
    # KEMITRAAN PEMBELAJARAN
    # ================================================


    kemitraan = f"""
Kemitraan pembelajaran dilakukan melalui:

- kolaborasi guru dan peserta didik.
- kerja kelompok antar peserta didik.
- pemanfaatan sumber belajar dari lingkungan.
- dukungan orang tua dalam penguatan belajar.
- penggunaan teknologi sebagai pendukung pembelajaran.
"""




    # ================================================
    # PEMANFAATAN DIGITAL
    # ================================================


    digital = f"""
Teknologi digital dimanfaatkan untuk:

- mencari sumber belajar terpercaya.
- menampilkan video atau media interaktif.
- membuat presentasi digital.
- melakukan evaluasi menggunakan media digital.

Contoh media:
- Google Classroom
- Quizizz
- Canva Pendidikan
- Video pembelajaran
"""




    # ================================================
    # LANGKAH PEMBELAJARAN
    # ================================================


    langkah = f"""
A. Pendahuluan

- Guru membuka pembelajaran.
- Apersepsi tentang {topik}.
- Menyampaikan tujuan pembelajaran.


B. Kegiatan Inti

- Orientasi masalah.
- Eksplorasi konsep.
- Diskusi kelompok.
- Presentasi hasil.
- Refleksi pembelajaran.


C. Penutup

- Guru bersama peserta didik menyimpulkan materi.
- Melakukan refleksi.
- Memberikan tindak lanjut.
"""




    # ================================================
    # ASESMEN
    # ================================================


    asesmen = f"""
1. Asesmen Diagnostik

Mengidentifikasi kemampuan awal peserta didik
tentang {topik}.


2. Asesmen Formatif

Teknik:
- observasi
- diskusi
- LKPD
- pertanyaan pemantik


3. Asesmen Sumatif

Teknik:
- tes tertulis
- proyek
- presentasi


Indikator keberhasilan:
Peserta didik mampu memahami konsep,
menerapkan pengetahuan, dan menunjukkan
keterampilan sesuai tujuan pembelajaran.
"""




    return {


        "dimensi_profil":

            dimensi,


        "tujuan_pembelajaran":

            tujuan,


        "praktik_pedagogis":

            pedagogis,


        "lingkungan_belajar":

            lingkungan,


        "kemitraan_belajar":

            kemitraan,


        "pemanfaatan_digital":

            digital,


        "langkah_pembelajaran":

            langkah,


        "asesmen_total":

            asesmen

    }





# =====================================================
# TOMBOL GENERATE RPM
# =====================================================


if st.button(

    "🚀 BUAT RPM CERDAS AI OFFLINE",

    type="primary",

    use_container_width=True

):


    if not topik.strip():


        st.warning(

            "Silakan isi Topik Pembelajaran terlebih dahulu."

        )


        st.stop()



    hasil_rpm = buat_rpm_offline(

        mapel,

        kelas,

        topik,

        cp,

        model_pembelajaran,

        tujuan_manual

    )



    st.session_state.hasil_rpm = hasil_rpm



    st.success(

        "✅ RPM berhasil dibuat tanpa API Key."

    )
    # =====================================================
# MENAMPILKAN HASIL RPM
# =====================================================


if "hasil_rpm" in st.session_state:


    st.divider()


    st.subheader(

        "📄 HASIL RPM CERDAS AI OFFLINE"

    )



    hasil = st.session_state.hasil_rpm




    daftar_komponen = [


        (

            "1. Dimensi Profil Lulusan",

            "dimensi_profil"

        ),


        (

            "2. Tujuan Pembelajaran",

            "tujuan_pembelajaran"

        ),


        (

            "3. Praktik Pedagogis",

            "praktik_pedagogis"

        ),


        (

            "4. Lingkungan Pembelajaran",

            "lingkungan_belajar"

        ),


        (

            "5. Kemitraan Pembelajaran",

            "kemitraan_belajar"

        ),


        (

            "6. Pemanfaatan Digital",

            "pemanfaatan_digital"

        ),


        (

            "7. Langkah Pembelajaran",

            "langkah_pembelajaran"

        ),


        (

            "8. Asesmen Pembelajaran",

            "asesmen_total"

        )

    ]




    for judul, kode in daftar_komponen:


        with st.expander(judul, expanded=False):


            st.write(

                hasil.get(

                    kode,

                    ""

                )

            )




    st.divider()



    st.subheader(

        "📌 Ringkasan Identitas RPM"

    )



    tabel_identitas = {


        "Nama Sekolah":

            sekolah,


        "Nama Guru":

            guru,


        "Mata Pelajaran":

            mapel,


        "Kelas":

            kelas,


        "Semester":

            semester,


        "Tahun Pelajaran":

            tahun_pelajaran,


        "Alokasi Waktu":

            alokasi_waktu,


        "Topik":

            topik

    }




    for k, v in tabel_identitas.items():


        st.write(

            f"**{k}:** {v}"

        )
        # =====================================================
# FUNGSI MEMBUAT DOKUMEN WORD
# =====================================================


def buat_dokumen_rpm(data):


    doc = Document()



    # ================================================
    # PENGATURAN HALAMAN
    # ================================================


    for section in doc.sections:

        section.top_margin = Inches(1)

        section.bottom_margin = Inches(1)

        section.left_margin = Inches(1)

        section.right_margin = Inches(1)



    style = doc.styles["Normal"]

    style.font.name = "Arial"

    style.font.size = Pt(11)




    # ================================================
    # JUDUL
    # ================================================


    judul = doc.add_heading(

        level=0

    )


    judul.alignment = 1


    run = judul.add_run(

        "RENCANA PEMBELAJARAN MENDALAM (RPM)"

    )


    run.bold = True

    run.font.size = Pt(16)



    doc.add_paragraph()




    # ================================================
    # IDENTITAS
    # ================================================


    doc.add_heading(

        "I. IDENTITAS PEMBELAJARAN",

        level=2

    )



    identitas = [

        ("Nama Sekolah", data["sekolah"]),

        ("Nama Guru", data["guru"]),

        ("Mata Pelajaran", data["mapel"]),

        ("Kelas / Semester",

         f'{data["kelas"]} / {data["semester"]}'),

        ("Tahun Pelajaran",

         data["tahun_pelajaran"]),

        ("Alokasi Waktu",

         data["alokasi_waktu"]),

        ("Topik",

         data["topik"])

    ]



    tabel = doc.add_table(

        rows=len(identitas),

        cols=2

    )


    tabel.style = "Table Grid"



    for i, (a,b) in enumerate(identitas):


        tabel.rows[i].cells[0].text = a

        tabel.rows[i].cells[1].text = str(b)




    doc.add_paragraph()




    # ================================================
    # KOMPONEN RPM
    # ================================================


    doc.add_heading(

        "II. KOMPONEN RPM",

        level=2

    )



    komponen = [


        ("1. Dimensi Profil Lulusan",
         data["dimensi_profil"]),


        ("2. Tujuan Pembelajaran",
         data["tujuan_pembelajaran"]),


        ("3. Praktik Pedagogis",
         data["praktik_pedagogis"]),


        ("4. Lingkungan Pembelajaran",
         data["lingkungan_belajar"]),


        ("5. Kemitraan Pembelajaran",
         data["kemitraan_belajar"]),


        ("6. Pemanfaatan Digital",
         data["pemanfaatan_digital"]),


        ("7. Langkah Pembelajaran",
         data["langkah_pembelajaran"]),


        ("8. Asesmen Pembelajaran",
         data["asesmen_total"])

    ]




    for judul, isi in komponen:


        doc.add_heading(

            judul,

            level=3

        )


        doc.add_paragraph(

            str(isi)

        )





    # ================================================
    # PENGESAHAN
    # ================================================


    doc.add_heading(

        "III. PENGESAHAN",

        level=2

    )


    tanda_tangan = doc.add_table(

        rows=1,

        cols=2

    )


    tanda_tangan.style = "Table Grid"



    tanda_tangan.cell(0,0).text = (

        "Mengetahui,\n\n"

        "Kepala Sekolah\n\n\n\n"

        "(............................)"

    )



    tanda_tangan.cell(0,1).text = (

        "Guru Mata Pelajaran\n\n\n\n"

        f"({data['guru']})"

    )




    file = io.BytesIO()


    doc.save(file)


    file.seek(0)


    return file






# =====================================================
# DOWNLOAD WORD
# =====================================================


if "hasil_rpm" in st.session_state:



    hasil = st.session_state.hasil_rpm



    data_word = {


        "sekolah":

            sekolah,


        "guru":

            guru,


        "mapel":

            mapel,


        "kelas":

            kelas,


        "semester":

            semester,


        "tahun_pelajaran":

            tahun_pelajaran,


        "alokasi_waktu":

            alokasi_waktu,


        "topik":

            topik,



        **hasil

    }




    file_word = buat_dokumen_rpm(

        data_word

    )



    st.download_button(

        label="⬇️ DOWNLOAD RPM WORD (.DOCX)",


        data=file_word,


        file_name="RPM_CERDAS_AI.docx",


        mime=(

            "application/vnd.openxmlformats-officedocument."

            "wordprocessingml.document"

        ),


        use_container_width=True

    )





# =====================================================
# TOMBOL RESET
# =====================================================


st.divider()



if st.button(

    "🔄 HAPUS HASIL RPM",

    use_container_width=True

):


    if "hasil_rpm" in st.session_state:

        del st.session_state["hasil_rpm"]



    st.success(

        "Hasil RPM berhasil dihapus."

    )


    st.rerun()




# =====================================================
# FOOTER
# =====================================================


st.divider()



st.markdown(

    """
    <center>

    📘 <b>RPM CERDAS AI OFFLINE</b>

    <br>

    Generator Rencana Pembelajaran Mendalam

    <br>

    Kurikulum Merdeka SMP

    </center>
    """,

    unsafe_allow_html=True

)
