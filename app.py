import os
import tempfile
import streamlit as st
from xlsx2html import xlsx2html
import pdfkit

st.set_page_config(page_title="🔄 แปลง Excel เป็น PDF", page_icon="📄")
st.title("📄 ตัวแปลง Excel → PDF (ใช้งานผ่านเว็บ)")

uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel", type=["xlsx"])

if uploaded_file and st.button("🚀 เริ่มแปลง"):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getbuffer())
        temp_excel_path = tmp.name

    with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as html_file:
        xlsx2html(temp_excel_path, html_file.name)
        html_path = html_file.name

    pdf_path = html_path.replace(".html", ".pdf")
    try:
        pdfkit.from_file(html_path, pdf_path)
        with open(pdf_path, "rb") as f:
            st.success("✅ แปลงสำเร็จ!")
            st.download_button("📥 ดาวน์โหลด PDF", f, file_name="converted.pdf", mime="application/pdf")
    except Exception as e:
        st.error(f"❌ แปลงไม่สำเร็จ: {e}")
