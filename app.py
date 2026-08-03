import io
import requests
import streamlit as st

from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# =====================================================
# GOOGLE GEMINI SDK
# =====================================================

try:
    from google import genai

    HAS_GENAI = True

except ImportError:

    HAS_GENAI = False

APP_NAME = "RPM CERDAS AI"

SUPPORTED_MODELS = [

    "gemini-2.5-flash",

]

REQUEST_TIMEOUT = 120

AI_STUDIO_URL = "https://aistudio.google.com/app/apikey"
def panggil_ai_guru(topik, cp, komponen_rpp, instruksi_khusus, api_key_ai):

    if not api_key_ai:
        return (
            "⚠️ Kunci API kosong.\n\n"
            f"Silakan buat API Key di:\n{AI_STUDIO_URL}"
        )

    clean_key = str(api_key_ai).strip()

    prompt = (
        f"Topik: {topik}\n"
        f"CP: {cp}\n"
        f"Komponen RPM: {komponen_rpp}\n"
        f"Instruksi: {instruksi_khusus}"
    )

    # =====================================================
    # GOOGLE GENAI SDK
    # =====================================================

    if HAS_GENAI:

        try:

            client = genai.Client(
                api_key=clean_key
            )

            error_terakhir = ""

            for model_name in SUPPORTED_MODELS:

                try:

                    response = client.models.generate_content(

                        model=model_name,

                        contents=prompt,

                    )

                    if (
                        response
                        and hasattr(response, "text")
                        and response.text
                    ):

                        return response.text.strip()

                except Exception as e:

                    error_terakhir = str(e)

                    print(
                        f"[SDK] {model_name} gagal : {e}"
                    )

                    continue

            if error_terakhir:

                if "404" in error_terakhir:

                    return (
                        "⚠️ Model Gemini tidak tersedia.\n\n"
                        "Silakan gunakan model terbaru."
                    )

                elif "429" in error_terakhir:

                    return (
                        "⚠️ Kuota Gemini telah habis.\n"
                        "Silakan coba beberapa saat lagi."
                    )

                elif "401" in error_terakhir:

                    return (
                        "⚠️ API Key tidak valid."
                    )

                return error_terakhir

        except Exception as e:

            print(e)

    # =====================================================
    # REST API (CADANGAN)
    # =====================================================

    try:

        headers = {

            "Content-Type": "application/json",

            "x-goog-api-key": clean_key,

        }

        payload = {

            "contents": [

                {

                    "parts": [

                        {

                            "text": prompt

                        }

                    ]

                }

            ]

        }

        error_api = ""

        for model_name in SUPPORTED_MODELS:

            url = (
                "https://generativelanguage.googleapis.com"
                f"/v1beta/models/{model_name}:generateContent"
            )

            response = requests.post(

                url,

                headers=headers,

                json=payload,

                timeout=REQUEST_TIMEOUT,

            )

            hasil = response.json()

            if response.status_code == 200:

                if (
                    "candidates" in hasil
                    and len(hasil["candidates"]) > 0
                ):

                    return hasil["candidates"][0]["content"]["parts"][0]["text"]

            if "error" in hasil:

                error_api = hasil["error"].get(
                    "message",
                    "Unknown Error",
                )

        return (
            "⚠️ Gemini tidak dapat memberikan jawaban.\n\n"
            + error_api
        )

    except Exception as e:

        return (
            "⚠️ Tidak dapat terhubung ke Gemini.\n\n"
            + str(e)
        )
        def buat_dokumen_rpm(data):

    doc = Document()

    # =====================================================
    # PENGATURAN HALAMAN
    # =====================================================

    for section in doc.sections:

        section.top_margin = Inches(1)

        section.bottom_margin = Inches(1)

        section.left_margin = Inches(1)

        section.right_margin = Inches(1)

    style = doc.styles["Normal"]

    style.font.name = "Arial"

    style.font.size = Pt(11)

    # =====================================================
    # JUDUL
    # =====================================================

    judul = doc.add_heading("", level=0)

    judul.alignment = 1

    run = judul.add_run(
        "RENCANA PEMBELAJARAN MENDALAM (RPM)"
    )

    run.bold = True

    run.font.size = Pt(16)

    doc.add_paragraph()

    # =====================================================
    # IDENTITAS
    # =====================================================

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

        ("Nama Sekolah", data.get("sekolah", "")),

        ("Nama Guru", data.get("guru", "")),

        ("Mata Pelajaran", data.get("mapel", "")),

        ("Kelas / Semester", data.get("kelas_semester", "")),

        ("Alokasi Waktu", data.get("alokasi_waktu", "")),

        ("Topik", data.get("topik", "")),

        ("Capaian Pembelajaran", data.get("cp", "")),

    ]

    for i, (judul, isi) in enumerate(identitas):

        tabel.rows[i].cells[0].text = judul

        tabel.rows[i].cells[1].text = str(isi)

        tabel.rows[i].cells[0].paragraphs[0].runs[0].bold = True

    doc.add_paragraph()

    # =====================================================
    # KOMPONEN RPM
    # =====================================================

    doc.add_heading(
        "II. KOMPONEN RPM",
        level=2
    )

    komponen = [

        (
            "1. Dimensi Profil Lulusan",
            data.get("dimensi_profil", "")
        ),

        (
            "2. Tujuan Pembelajaran",
            data.get("tujuan_pembelajaran", "")
        ),

        (
            "3. Praktik Pedagogis",
            data.get("praktik_pedagogis", "")
        ),

        (
            "4. Lingkungan Pembelajaran",
            data.get("lingkungan_belajar", "")
        ),

        (
            "5. Kemitraan Pembelajaran",
            data.get("kemitraan_belajar", "")
        ),

        (
            "6. Pemanfaatan Digital",
            data.get("pemanfaatan_digital", "")
        ),

        (
            "7. Langkah Pembelajaran",
            data.get("langkah_pembelajaran", "")
        ),

        (
            "8. Asesmen Pembelajaran",
            data.get("asesmen_total", "")
        ),

    ]

    tabel2 = doc.add_table(
        rows=len(komponen) + 1,
        cols=2
    )

    tabel2.style = "Table Grid"

    tabel2.rows[0].cells[0].text = "Komponen"

    tabel2.rows[0].cells[1].text = "Isi RPM"

    for c in tabel2.rows[0].cells:

        c.paragraphs[0].runs[0].bold = True

        c._tc.get_or_add_tcPr().append(

            parse_xml(

                r'<w:shd {} w:fill="D9EAD3"/>'.format(

                    nsdecls("w")

                )

            )

        )

    for i, (judul, isi) in enumerate(komponen):

        tabel2.rows[i + 1].cells[0].text = judul

        tabel2.rows[i + 1].cells[1].text = str(isi)

        tabel2.rows[i + 1].cells[0].paragraphs[0].runs[0].bold = True

    doc.add_paragraph()

    # =====================================================
    # PENGESAHAN
    # =====================================================

    doc.add_heading(
        "III. PENGESAHAN",
        level=2
    )

    tanda = doc.add_table(
        rows=1,
        cols=2
    )

    kiri = tanda.rows[0].cells[0]

    kanan = tanda.rows[0].cells[1]

    kiri.text = (
        "Mengetahui,\n\n"
        "Kepala Sekolah\n\n\n\n\n"
        "(....................................)"
    )

    kanan.text = (
        f"Guru Mata Pelajaran\n\n\n\n\n({data.get('guru','')})"
    )

    stream = io.BytesIO()

    doc.save(stream)

    stream.seek(0)

    return stream
    if st.button("✨ 1 & 2. Rumuskan Profil Lulusan & Tujuan (AI)"):

    if not api_key_input.strip():

        st.error("Silakan masukkan Gemini API Key terlebih dahulu.")

        st.stop()

    with st.spinner("AI sedang menyusun Profil Lulusan..."):

        st.session_state.profil_ai = panggil_ai_guru(
            topik,
            cp,
            "Dimensi Profil Lulusan",
            "Rincikan keterampilan abad ke-21.",
            api_key_input,
        )

        st.session_state.tujuan_ai = panggil_ai_guru(
            topik,
            cp,
            "Tujuan Pembelajaran",
            "Buat tujuan pembelajaran yang berkesadaran, bermakna, dan menggembirakan.",
            api_key_input,
        )

    st.rerun()
