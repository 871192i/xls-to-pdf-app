import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ลงทะเบียนฟอนต์ Sarabun
pdfmetrics.registerFont(TTFont("THSarabun", "THSarabunNew001.ttf"))
pdfmetrics.registerFont(TTFont("THSarabun-Bold", "THSarabunNew Bold001.ttf"))

# ตั้งค่าหน้าจอแอป
st.set_page_config(page_title="ตัวแปลง Excel → PDF (ฟอนต์ภาษาไทย)", layout="centered")

st.markdown("## 🧾 ตัวแปลง Excel → PDF (ใช้ฟอนต์ภาษาไทย)")
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="xlrd" if uploaded_file.name.endswith(".xls") else None, na_filter=False)
        st.success("📄 โหลดข้อมูลสำเร็จ")
        st.dataframe(df)
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")
    else:
        if st.button("📤 สร้าง PDF"):
            pdf_path = "output.pdf"
            doc = SimpleDocTemplate(pdf_path, pagesize=A4)
            elements = []

            styles = getSampleStyleSheet()
            styles.add(ParagraphStyle(name='ThaiStyle', fontName="THSarabun", fontSize=14))

            elements.append(Paragraph("รายงานจาก Excel (ภาษาไทย)", styles['ThaiStyle']))
            elements.append(Spacer(1, 12))

            data = [df.columns.tolist()] + df.values.tolist()

            table = Table(data, hAlign='LEFT')
            table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, 0), 'THSarabun-Bold'),
                ('FONTNAME', (0, 1), (-1, -1), 'THSarabun'),
                ('FONTSIZE', (0, 0), (-1, -1), 12),
                ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER')
            ]))

            elements.append(table)
            doc.build(elements)

            with open(pdf_path, "rb") as f:
                st.download_button("📥 ดาวน์โหลด PDF", f, file_name="excel_report.pdf")
