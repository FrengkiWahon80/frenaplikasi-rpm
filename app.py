# =====================================================
# RPM CERDAS AI
# Generator Rencana Pembelajaran Mendalam
# Kurikulum Merdeka SMP
# =====================================================


import io
import requests
import streamlit as st


from docx import Document
from docx.shared import Inches, Pt



# =====================================================
# KONFIGURASI APLIKASI
# =====================================================


APP_NAME = "RPM CERDAS AI"


SUPPORTED_MODELS = [

    "gemini-2.5-flash"

]


REQUEST_TIMEOUT = 120


AI_STUDIO_URL = (
    "https://aistudio.google.com/app/apikey"
)



# =====================================================
# GOOGLE GEMINI SDK
# =====================================================


try:

    from google import genai

    HAS_GENAI = True


except ImportError:

    HAS_GENAI = False




# =====================================================
# SETTING HALAMAN STREAMLIT
# =====================================================


st.set_page_config(

    page_title=APP_NAME,

    page_icon="📘",

    layout="wide"

)




# =====================================================
# FUNGSI GEMINI AI
# =====================================================


def panggil_ai_guru(

    topik,

    cp,

    komponen_rpm,

    instruksi,

    api_key


):


    if not api_key:


        return (

            "⚠️ API Key Gemini belum dimasukkan.\n\n"

            "Silakan masukkan API Key dari:\n"

            f"{AI_STUDIO_URL}"

        )



    clean_key = str(api_key).strip()




    prompt = f"""

Anda adalah asisten ahli Kurikulum Merdeka.

Buatkan komponen Rencana Pembelajaran Mendalam (RPM).

Topik:
{topik}


Capaian Pembelajaran:
{cp}


Komponen:
{komponen_rpm}


Instruksi:
{instruksi}


Gunakan bahasa Indonesia yang jelas,
praktis, dan mudah diterapkan guru SMP.

"""




    # =================================================
    # GOOGLE GENAI SDK
    # =================================================


    if HAS_GENAI:


        try:


            client = genai.Client(

                api_key=clean_key

            )



            for model_name in SUPPORTED_MODELS:



                try:



                    response = client.models.generate_content(


                        model=model_name,


                        contents=prompt


                    )



                    if (

                        response

                        and hasattr(response, "text")

                        and response.text

                    ):


                        return response.text.strip()



                except Exception as e:


                    pesan = str(e)



                    if "429" in pesan or "quota" in pesan.lower():


                        return (

                            "⚠️ Kuota Gemini habis.\n\n"

                            "Silakan:\n"

                            "1. Gunakan API Key lain\n"

                            "2. Tunggu kuota diperbarui\n"

                            "3. Periksa paket Gemini API"

                        )



                    print(

                        "Gemini SDK error:",

                        pesan

                    )



        except Exception as e:


            print(

                "SDK Error:",

                e

            )






    # =================================================
    # REST API CADANGAN
    # =================================================


    try:


        headers = {


            "Content-Type":

            "application/json",



            "x-goog-api-key":

            clean_key

        }




        payload = {


            "contents":

            [

                {

                    "parts":

                    [

                        {

                            "text": prompt

                        }

                    ]

                }

            ]

        }




        for model_name in SUPPORTED_MODELS:



            url = (

                "https://generativelanguage.googleapis.com"

                f"/v1beta/models/{model_name}:generateContent"

            )



            response = requests.post(

                url,

                headers=headers,

                json=payload,

                timeout=REQUEST_TIMEOUT

            )



            hasil = response.json()




            if response.status_code == 200:



                kandidat = hasil.get(

                    "candidates",

                    []

                )



                if kandidat:


                    return (

                        kandidat[0]

                        ["content"]

                        ["parts"]

                        [0]

                        ["text"]

                    )





            if "error" in hasil:



                pesan_error = hasil["error"].get(

                    "message",

                    ""

                )



                if "quota" in pesan_error.lower():



                    return (

                        "⚠️ Kuota Gemini habis.\n\n"

                        "Silakan gunakan API Key lain."

                    )



                if "401" in str(response.status_code):


                    return (

                        "⚠️ API Key Gemini tidak valid."

                    )



                return (

                    "⚠️ Gemini Error:\n\n"

                    + pesan_error

                )



    except Exception as e:



        return (

            "⚠️ Tidak dapat terhubung ke Gemini.\n\n"

            + str(e)

        )




    return (

        "⚠️ Gemini belum memberikan jawaban."

    )
    # =====================================================
# HEADER APLIKASI
# =====================================================


st.markdown(
    """
    <h1 style="text-align:center;">
    📘 RPM CERDAS AI
    </h1>

    <p style="text-align:center;">
    Generator Rencana Pembelajaran Mendalam
    Kurikulum Merdeka SMP Berbasis AI
    </p>

    """,

    unsafe_allow_html=True

)



st.divider()



# =====================================================
# INPUT API KEY GEMINI
# =====================================================


st.subheader(
    "🔑 Koneksi Gemini AI"
)



api_key_input = st.text_input(

    "Masukkan Gemini API Key",

    type="password",

    placeholder="Tempel API Key Gemini di sini"

)



st.caption(

    "API Key dapat dibuat melalui Google AI Studio."

)



st.divider()




# =====================================================
# IDENTITAS PEMBELAJARAN
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
# INFORMASI MATERI
# =====================================================


st.subheader(

    "📚 Informasi Materi Pembelajaran"

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

    placeholder=(

        "Masukkan Capaian Pembelajaran "
        "sesuai Kurikulum Merdeka..."

    )

)



karakteristik = st.text_area(

    "Karakteristik Peserta Didik (Opsional)",

    height=100

)



tujuan_manual = st.text_area(

    "Tujuan Pembelajaran Manual (Opsional)",

    height=120

)



st.info(

    """
    💡 Tips:
    Isi CP dan topik secara lengkap agar
    hasil RPM AI lebih sesuai dengan kebutuhan kelas.
    """

)



st.divider()




# =====================================================
# GENERATE PROFIL DAN TUJUAN AWAL
# =====================================================


if st.button(

    "✨ BUAT PROFIL LULUSAN & TUJUAN AI",

    type="primary",

    use_container_width=True

):


    if not api_key_input.strip():


        st.error(

            "Masukkan Gemini API Key terlebih dahulu."

        )


        st.stop()



    with st.spinner(

        "AI sedang membuat Profil Lulusan dan Tujuan..."

    ):



        profil_ai = panggil_ai_guru(

            topik,

            cp,

            "Dimensi Profil Lulusan",

            """
            Buatkan dimensi profil lulusan.
            Sesuaikan dengan keterampilan abad 21,
            karakter peserta didik,
            dan pembelajaran mendalam.
            """,

            api_key_input

        )



        tujuan_ai = panggil_ai_guru(

            topik,

            cp,

            "Tujuan Pembelajaran",

            """
            Buat tujuan pembelajaran yang:
            - berkesadaran
            - bermakna
            - menggembirakan
            - dapat diukur
            """,

            api_key_input

        )



        st.session_state.profil_ai = profil_ai

        st.session_state.tujuan_ai = tujuan_ai



    st.success(

        "Profil dan tujuan berhasil dibuat."

    )
    # =====================================================
# FUNGSI GENERATOR RPM LENGKAP
# =====================================================


def buat_semua_komponen_rpm(

    topik,

    cp,

    mapel,

    kelas,

    model,

    api_key

):


    hasil = {}



    hasil["dimensi_profil"] = panggil_ai_guru(

        topik,

        cp,

        "Dimensi Profil Lulusan",

        """
        Buatkan dimensi profil lulusan.

        Jelaskan:
        - karakter yang dikembangkan
        - keterampilan abad 21
        - kemampuan berpikir kritis
        - kreativitas
        - kolaborasi
        - komunikasi
        """,

        api_key

    )



    hasil["tujuan_pembelajaran"] = panggil_ai_guru(

        topik,

        cp,

        "Tujuan Pembelajaran",

        """
        Buat tujuan pembelajaran berdasarkan
        prinsip pembelajaran mendalam:

        - berkesadaran
        - bermakna
        - menggembirakan

        Gunakan kalimat yang dapat diukur.
        """,

        api_key

    )



    hasil["praktik_pedagogis"] = panggil_ai_guru(

        topik,

        cp,

        "Praktik Pedagogis",

        f"""
        Gunakan model pembelajaran:

        {model}

        Jelaskan pendekatan pembelajaran
        berbasis masalah.

        Sertakan:
        - aktivitas guru
        - aktivitas peserta didik
        - strategi pembelajaran aktif
        """,

        api_key

    )



    hasil["lingkungan_belajar"] = panggil_ai_guru(

        topik,

        cp,

        "Lingkungan Pembelajaran",

        """
        Buat lingkungan pembelajaran yang mendukung:

        - kelas aman dan nyaman
        - budaya belajar positif
        - ruang kolaborasi
        - suasana menyenangkan
        """,

        api_key

    )



    hasil["kemitraan_belajar"] = panggil_ai_guru(

        topik,

        cp,

        "Kemitraan Pembelajaran",

        """
        Jelaskan bentuk kemitraan:

        - guru
        - peserta didik
        - orang tua
        - komunitas
        - teknologi digital

        untuk mendukung pembelajaran.
        """,

        api_key

    )



    hasil["pemanfaatan_digital"] = panggil_ai_guru(

        topik,

        cp,

        "Pemanfaatan Digital",

        """
        Jelaskan pemanfaatan teknologi digital
        dalam pembelajaran.

        Sertakan contoh:
        - aplikasi
        - media digital
        - sumber belajar online
        - AI pembelajaran
        """,

        api_key

    )



    hasil["langkah_pembelajaran"] = panggil_ai_guru(

        topik,

        cp,

        "Langkah Pembelajaran",

        """
        Susun langkah pembelajaran lengkap:

        A. Pendahuluan

        B. Kegiatan Inti

        C. Penutup


        Gunakan pembelajaran mendalam.

        Sertakan:
        - aktivitas guru
        - aktivitas peserta didik
        - refleksi
        """,

        api_key

    )



    hasil["asesmen_total"] = panggil_ai_guru(

        topik,

        cp,

        "Asesmen Pembelajaran",

        """
        Buat asesmen lengkap:

        1. Diagnostik

        2. Formatif

        3. Sumatif


        Sertakan:
        - teknik penilaian
        - instrumen
        - indikator keberhasilan
        """,

        api_key

    )



    return hasil





# =====================================================
# TOMBOL GENERATE RPM LENGKAP
# =====================================================


st.divider()



if st.button(

    "🚀 GENERATE RPM LENGKAP AI",

    type="primary",

    use_container_width=True

):


    if not api_key_input.strip():


        st.error(

            "Masukkan Gemini API Key terlebih dahulu."

        )


        st.stop()



    if not topik.strip():


        st.warning(

            "Silakan isi Topik Pembelajaran terlebih dahulu."

        )


        st.stop()




    with st.spinner(

        "AI sedang menyusun RPM lengkap... Mohon tunggu."

    ):



        hasil_rpm = buat_semua_komponen_rpm(

            topik,

            cp,

            mapel,

            kelas,

            model_pembelajaran,

            api_key_input

        )



        st.session_state.hasil_rpm = hasil_rpm



    st.success(

        "✅ RPM lengkap berhasil dibuat."

    )
    # =====================================================
# FUNGSI MEMBUAT FILE WORD RPM
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



    tulisan = judul.add_run(

        "RENCANA PEMBELAJARAN MENDALAM (RPM)"

    )


    tulisan.bold = True


    tulisan.font.size = Pt(16)




    doc.add_paragraph()




    # ================================================
    # IDENTITAS
    # ================================================


    doc.add_heading(

        "I. IDENTITAS PEMBELAJARAN",

        level=2

    )



    tabel = doc.add_table(

        rows=7,

        cols=2

    )


    tabel.style = "Table Grid"



    identitas = [


        ("Nama Sekolah", data.get("sekolah","")),


        ("Nama Guru", data.get("guru","")),


        ("Mata Pelajaran", data.get("mapel","")),


        ("Kelas / Semester", data.get("kelas_semester","")),


        ("Alokasi Waktu", data.get("alokasi_waktu","")),


        ("Topik", data.get("topik","")),


        ("Capaian Pembelajaran", data.get("cp",""))

    ]



    for i, (judul, isi) in enumerate(identitas):


        tabel.rows[i].cells[0].text = judul


        tabel.rows[i].cells[1].text = str(isi)



        tabel.rows[i].cells[0].paragraphs[0].runs[0].bold = True





    # ================================================
    # KOMPONEN RPM
    # ================================================


    doc.add_heading(

        "II. KOMPONEN RPM",

        level=2

    )



    komponen = [


        (
            "1. Dimensi Profil Lulusan",

            data.get("dimensi_profil","")

        ),


        (
            "2. Tujuan Pembelajaran",

            data.get("tujuan_pembelajaran","")

        ),


        (
            "3. Praktik Pedagogis",

            data.get("praktik_pedagogis","")

        ),


        (
            "4. Lingkungan Pembelajaran",

            data.get("lingkungan_belajar","")

        ),


        (
            "5. Kemitraan Pembelajaran",

            data.get("kemitraan_belajar","")

        ),


        (
            "6. Pemanfaatan Digital",

            data.get("pemanfaatan_digital","")

        ),


        (
            "7. Langkah Pembelajaran",

            data.get("langkah_pembelajaran","")

        ),


        (
            "8. Asesmen Pembelajaran",

            data.get("asesmen_total","")

        )

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



    tabel_ttd = doc.add_table(

        rows=1,

        cols=2

    )



    tabel_ttd.style = "Table Grid"



    tabel_ttd.cell(0,0).text = (

        "Mengetahui,\n\n"

        "Kepala Sekolah\n\n\n\n"

        "(................................)"

    )



    tabel_ttd.cell(0,1).text = (

        "Guru Mata Pelajaran\n\n\n\n"

        f"({data.get('guru','')})"

    )





    file = io.BytesIO()


    doc.save(file)


    file.seek(0)



    return file






# =====================================================
# MENAMPILKAN HASIL RPM
# =====================================================


if "hasil_rpm" in st.session_state:



    st.divider()



    st.subheader(

        "📄 HASIL RPM CERDAS AI"

    )



    hasil = st.session_state.hasil_rpm




    daftar = [


        (
            "Dimensi Profil Lulusan",

            "dimensi_profil"

        ),


        (
            "Tujuan Pembelajaran",

            "tujuan_pembelajaran"

        ),


        (
            "Praktik Pedagogis",

            "praktik_pedagogis"

        ),


        (
            "Lingkungan Pembelajaran",

            "lingkungan_belajar"

        ),


        (
            "Kemitraan Pembelajaran",

            "kemitraan_belajar"

        ),


        (
            "Pemanfaatan Digital",

            "pemanfaatan_digital"

        ),


        (
            "Langkah Pembelajaran",

            "langkah_pembelajaran"

        ),


        (
            "Asesmen Pembelajaran",

            "asesmen_total"

        )

    ]




    for judul, key in daftar:


        with st.expander(judul):


            st.write(

                hasil.get(

                    key,

                    ""

                )

            )





    # ================================================
    # DOWNLOAD WORD
    # ================================================


    data_word = {


        "sekolah": sekolah,


        "guru": guru,


        "mapel": mapel,


        "kelas_semester":

        f"{kelas} / {semester}",



        "alokasi_waktu":

        alokasi_waktu,



        "topik":

        topik,



        "cp":

        cp,



        **hasil

    }





    file_docx = buat_dokumen_rpm(

        data_word

    )




    st.download_button(

        label="⬇️ DOWNLOAD RPM WORD (.DOCX)",


        data=file_docx,


        file_name="RPM_CERDAS_AI.docx",


        mime=(

            "application/vnd.openxmlformats-officedocument."

            "wordprocessingml.document"

        ),


        use_container_width=True

    )
    # =====================================================
# CSS TAMPILAN PROFESIONAL
# =====================================================


st.markdown(

    """
    <style>

    .stButton button {

        border-radius: 10px;

        font-weight: bold;

        height: 45px;

    }


    .stTextInput input,
    .stTextArea textarea {

        border-radius: 8px;

    }


    [data-testid="stExpander"] {

        border-radius: 10px;

    }


    </style>

    """,

    unsafe_allow_html=True

)





# =====================================================
# SIDEBAR INFORMASI
# =====================================================


with st.sidebar:



    st.title(

        "📘 RPM CERDAS AI"

    )



    st.write(

        """
        Generator Rencana Pembelajaran Mendalam
        untuk guru SMP Kurikulum Merdeka.

        Membantu guru membuat RPM lebih cepat,
        praktis, dan terstruktur dengan bantuan AI.
        """

    )



    st.divider()



    st.subheader(

        "📌 Langkah Penggunaan"

    )



    st.markdown(

        """
        1. Masukkan Gemini API Key

        2. Isi identitas pembelajaran

        3. Masukkan Topik dan CP

        4. Klik Generate RPM AI

        5. Download Word

        """

    )



    st.divider()



    st.caption(

        "RPM CERDAS AI © 2026"

    )






# =====================================================
# RESET HASIL
# =====================================================



st.divider()



kolom_a, kolom_b = st.columns(2)




with kolom_a:



    if st.button(

        "🔄 HAPUS HASIL RPM",

        use_container_width=True

    ):



        if "hasil_rpm" in st.session_state:

            del st.session_state["hasil_rpm"]



        if "profil_ai" in st.session_state:

            del st.session_state["profil_ai"]



        if "tujuan_ai" in st.session_state:

            del st.session_state["tujuan_ai"]



        st.success(

            "Hasil RPM berhasil dihapus."

        )



        st.rerun()





with kolom_b:



    st.info(

        """
        💡 Gunakan API Key masing-masing guru
        agar kuota Gemini aman.
        """

    )







# =====================================================
# FOOTER
# =====================================================



st.divider()



st.markdown(

    """
    <center>

    <b>
    📘 RPM CERDAS AI
    </b>

    <br>

    Generator Rencana Pembelajaran Mendalam

    <br>

    Kurikulum Merdeka SMP

    <br><br>

    Dibuat untuk membantu guru Indonesia

    </center>

    """,

    unsafe_allow_html=True

)
