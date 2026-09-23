import os
import streamlit as st
from docx import Document

# Konfigurasi Halaman
st.set_page_config(
    page_title="Bot Surat Rasmi SK Kundang Ulu",
    page_icon="🏫",
    layout="centered"
)

# Tajuk Sistem
st.title("🏫 Bot Surat Rasmi SK Kundang Ulu")
st.markdown("Sistem penjanaan surat rasmi sekolah dengan pantas dan mengekalkan format asal templat.")
st.markdown("---")

# Direktori templat
TEMPLATE_PATH = "template_surat.docx"

# Fungsi untuk mengisi templat Word
def generate_docx(data):
    if not os.path.exists(TEMPLATE_PATH):
        return None
    
    doc = Document(TEMPLATE_PATH)
    
    # Padanan kata kunci (Placeholder) di dalam dokumen Word asal
    # Contoh tag dalam Word: {{rujukan_tuan}}, {{rujukan_kami}}, {{tarikh}}, dll.
    replacements = {
        "{{rujukan_tuan}}": data.get("rujukan_tuan", ""),
        "{{rujukan_kami}}": data.get("rujukan_kami", ""),
        "{{tarikh}}": data.get("tarikh", ""),
        "{{nama_penerima}}": data.get("nama_penerima", ""),
        "{{jawatan_penerima}}": data.get("jawatan_penerima", ""),
        "{{organisasi}}": data.get("organisasi", ""),
        "{{alamat_penerima}}": data.get("alamat_penerima", ""),
        "{{perkara}}": data.get("perkara", ""),
        "{{isi_kandungan}}": data.get("isi_kandungan", ""),
        "{{tarikh_program}}": data.get("tarikh_program", ""),
        "{{masa_program}}": data.get("masa_program", ""),
        "{{tempat_program}}": data.get("tempat_program", ""),
    }

    # Semak dan gantikan teks dalam perenggan dokumen
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, val)

    # Semak dan gantikan teks dalam jadual (jika ada bahagian surat dalam bentuk jadual)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in replacements.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, val)

    output_path = "Surat_Rasmi_Janaan.docx"
    doc.save(output_path)
    return output_path

# Borang Maklumat Surat
with st.form("surat_form"):
    st.subheader("📝 Borang Maklumat Surat Rasmi")
    
    col1, col2 = st.columns(2)
    with col1:
        rujukan_tuan = st.text_input("Rujukan Tuan (jika ada)")
        rujukan_kami = st.text_input("Rujukan Kami", value="SKKU/600-4/1/ (  )")
        tarikh = st.text_input("Tarikh Surat", value="11 Februari 2026")
    with col2:
        perkara = st.text_input("Perkara / Tajuk Surat")
        nama_penerima = st.text_input("Nama Penerima")
        jawatan_penerima = st.text_input("Jawatan Penerima")

    organisasi = st.text_input("Nama Organisasi / Sekolah Penerima")
    alamat_penerima = st.text_area("Alamat Penerima")

    st.markdown("### Butiran Kandungan & Program")
    isi_kandungan = st.text_area("Isi Perenggan Utama Surat")
    
    col3, col4, col5 = st.columns(3)
    with col3:
        tarikh_program = st.text_input("Tarikh Program (Jika berkaitan)")
    with col4:
        masa_program = st.text_input("Masa Program (Jika berkaitan)")
    with col5:
        tempat_program = st.text_input("Tempat Program (Jika berkaitan)")

    submitted = st.form_submit_button("Jana Surat Rasmi")

if submitted:
    if not perkara or not nama_penerima:
        st.error("Sila lengkapkan ruangan Perkara dan Nama Penerima sekurang-kurangnya.")
    else:
        form_data = {
            "rujukan_tuan": rujukan_tuan,
            "rujukan_kami": rujukan_kami,
            "tarikh": tarikh,
            "nama_penerima": nama_penerima,
            "jawatan_penerima": jawatan_penerima,
            "organisasi": organisasi,
            "alamat_penerima": alamat_penerima,
            "perkara": perkara,
            "isi_kandungan": isi_kandungan,
            "tarikh_program": tarikh_program,
            "masa_program": masa_program,
            "tempat_program": tempat_program,
        }
        
        output_file = generate_docx(form_data)
        if output_file:
            st.success("Surat berjaya dijana dengan mengekalkan format templat!")
            with open(output_file, "rb") as f:
                st.download_button(
                    label="📥 Muat Turun Surat (.docx)",
                    data=f,
                    file_name="Surat_Rasmi_SKKU.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.error("Fail templat 'template_surat.docx' tidak dijumpai di dalam sistem. Sila muat naik fail templat terlebih dahulu.")