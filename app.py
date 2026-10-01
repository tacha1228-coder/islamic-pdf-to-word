import streamlit as st
import io
import tempfile
import os
import re

import pdfplumber
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

st.set_page_config(page_title="محول PDF إلى Word", page_icon="📖", layout="wide")

st.title("📖 محول PDF إلى Word")
st.markdown("**مخصص لتفريغ دروس العلماء والمحتوى الديني** بأمانة علمية كاملة")
st.caption("✅ يحافظ على النص كما هو | ✅ يدعم التشكيل | ✅ يوثّق المصدر")

with st.sidebar:
    st.header("⚙️ الإعدادات")
    font_choice = st.selectbox("الخط العربي", ["Traditional Arabic", "Amiri", "Arial", "Scheherazade New"])
    font_size = st.slider("حجم الخط", 10, 20, 14)
    add_page_numbers = st.checkbox("إضافة أرقام الصفحات", value=True)
    add_source = st.checkbox("إضافة اسم الملف الأصلي", value=True)

uploaded = st.file_uploader("اختاري ملف PDF", type=["pdf"])

def set_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    pPr.append(bidi)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def clean_text(text):
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

if uploaded:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.write(uploaded.read())
    tmp.close()
    st.success(f"✅ تم رفع الملف: {uploaded.name}")
    
    if st.button("🚀 ابدأ التحويل", type="primary"):
        with st.spinner("جارٍ استخراج النص من الصفحات..."):
            try:
                pages_data = []
                pdf_metadata = {}
                with pdfplumber.open(tmp.name) as pdf:
                    pdf_metadata = pdf.metadata or {}
                    total = len(pdf.pages)
                    progress = st.progress(0)
                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text(x_tolerance=2, y_tolerance=2) or ""
                        pages_data.append(text)
                        progress.progress((i + 1) / total)
                all_text = "\n\n".join([p for p in pages_data if p.strip()])
                if not all_text.strip():
                    st.warning("⚠️ لا يوجد نص قابل للاستخراج. قد يكون الملف صورًا.")
                else:
                    st.session_state['pages'] = pages_data
                    st.session_state['filename'] = uploaded.name
                    st.session_state['metadata'] = pdf_metadata
                    st.success(f"✅ تم استخراج {len(pages_data)} صفحة!")
            except Exception as e:
                st.error(f"حدث خطأ: {e}")
            finally:
                try:
                    os.unlink(tmp.name)
                except:
                    pass

if 'pages' in st.session_state and st.session_state['pages']:
    pages = st.session_state['pages']
    filename = st.session_state['filename']
    meta = st.session_state.get('metadata', {})
    
    st.subheader("📝 معاينة النص")
    preview_text = "\n\n".join([p for p in pages[:3] if p.strip()])
    st.text_area("أول 3 صفحات", preview_text, height=250)
    
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = font_choice
    style.font.size = Pt(font_size)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), font_choice)
    style.element.rPr.rFonts.set(qn('w:cs'), font_choice)
    
    if add_source:
        title = doc.add_heading(f"المصدر: {filename}", level=1)
        set_rtl(title)
        if meta.get('Title'):
            p = doc.add_paragraph(f"العنوان الأصلي: {meta['Title']}")
            set_rtl(p)
        if meta.get('Author'):
            p = doc.add_paragraph(f"المؤلف: {meta['Author']}")
            set_rtl(p)
        doc.add_paragraph("—" * 30)
    
    for i, page_text in enumerate(pages, start=1):
        if not page_text.strip():
            continue
        if add_page_numbers:
            page_marker = doc.add_paragraph()
            set_rtl(page_marker)
            run = page_marker.add_run(f"【 صفحة {i} 】")
            run.bold = True
            run.font.color.rgb = RGBColor(0x88, 0x44, 0x00)
        cleaned = clean_text(page_text)
        for line in cleaned.split('\n'):
            if line.strip():
                p = doc.add_paragraph(line)
                set_rtl(p)
        doc.add_paragraph()
    
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    
    st.divider()
    st.subheader("📥 تحميل الملف")
    st.download_button(
        "📄 تحميل ملف Word",
        buf,
        f"converted_{filename.replace('.pdf', '')}.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.info("💡 الملف نصي بالكامل — انسخي مع ذكر المصدر للأمانة العلمية")
