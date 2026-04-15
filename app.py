import streamlit as st
from docxtpl import DocxTemplate
from docx import Document
import io
import zipfile
import pandas as pd
from datetime import date
import re
import subprocess
import tempfile
import os

# ─── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Word Template Filler",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Helper: ตรวจชื่อตัวแปรแล้วเลือก input widget ที่เหมาะสม ─────────────────
def get_input_widget(var: str, label: str, key: str):
    v = var.lower()
    if re.search(r"date|วันที่|วันเกิด|start_date|end_date|เริ่ม|สิ้นสุด", v):
        val = st.date_input(label, key=key)
        return val.strftime("%d/%m/%Y")
    elif re.search(r"note|detail|description|remark|address|ที่อยู่|หมายเหตุ|รายละเอียด", v):
        return st.text_area(label, key=key, placeholder=f"กรอก {label}", height=80)
    elif re.search(r"email|อีเมล", v):
        return st.text_input(label, key=key, placeholder="example@email.com")
    elif re.search(r"phone|tel|mobile|โทร|เบอร์", v):
        return st.text_input(label, key=key, placeholder="0XX-XXX-XXXX")
    elif re.search(r"amount|price|cost|salary|ราคา|เงิน|ค่า|จำนวนเงิน|budget", v):
        val = st.number_input(label, key=key, min_value=0.0, step=1.0, format="%.2f")
        return f"{val:,.2f}"
    else:
        return st.text_input(label, key=key, placeholder=f"กรอก {label}")

# ─── Helper: สร้าง sample template (หนังสือรับรองการทำงาน) ───────────────────
@st.cache_data
def create_sample_template() -> bytes:
    doc = Document()
    doc.add_heading("หนังสือรับรองการทำงาน", 0)
    doc.add_paragraph("")
    doc.add_paragraph("วันที่  {{ date }}")
    doc.add_paragraph("")
    p = doc.add_paragraph()
    p.add_run(
        "หนังสือฉบับนี้ขอรับรองว่า  {{ full_name }}  อายุ  {{ age }}  ปี "
        "ดำรงตำแหน่ง  {{ position }}  ได้ปฏิบัติงานกับ  {{ company_name }} "
        "ตั้งแต่วันที่  {{ start_date }}  จนถึงปัจจุบัน "
        "โดยได้รับเงินเดือน  {{ salary }}  บาทต่อเดือน"
    )
    doc.add_paragraph("")
    doc.add_paragraph("ที่อยู่ปัจจุบัน:  {{ address }}")
    doc.add_paragraph("อีเมล:  {{ email }}")
    doc.add_paragraph("เบอร์โทรศัพท์:  {{ phone }}")
    doc.add_paragraph("")
    doc.add_paragraph("จึงออกหนังสือรับรองฉบับนี้ให้ไว้เพื่อเป็นหลักฐาน")
    doc.add_paragraph("")
    doc.add_paragraph("ลงชื่อ  ____________________________")
    doc.add_paragraph("(  {{ approver_name }}  )")
    doc.add_paragraph("{{ approver_position }}")
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

# ─── Helper: เรนเดอร์เทมเพลตและคืน bytes ─────────────────────────────────────
def render_doc(template_bytes: bytes, context: dict) -> bytes:
    doc = DocxTemplate(io.BytesIO(template_bytes))
    doc.render(context)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

# ─── Helper: แปลง docx bytes → PDF bytes ─────────────────────────────────────
def convert_to_pdf(docx_bytes: bytes) -> bytes:
    """ลอง docx2pdf (ต้องมี Word) ก่อน ถ้าไม่มีก็ใช้ LibreOffice"""
    tmp_dir = tempfile.mkdtemp()
    tmp_docx = os.path.join(tmp_dir, "input.docx")
    tmp_pdf  = os.path.join(tmp_dir, "input.pdf")

    try:
        with open(tmp_docx, "wb") as f:
            f.write(docx_bytes)

        # ── ลอง Microsoft Word ก่อน (docx2pdf) ───────────────────────────────
        try:
            from docx2pdf import convert
            convert(tmp_docx, tmp_pdf)
        except Exception:
            # ── Fallback: LibreOffice headless ────────────────────────────────
            soffice_candidates = [
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                "soffice",
            ]
            converted = False
            for soffice in soffice_candidates:
                result = subprocess.run(
                    [soffice, "--headless", "--convert-to", "pdf",
                     "--outdir", tmp_dir, tmp_docx],
                    capture_output=True, timeout=60,
                )
                if result.returncode == 0:
                    converted = True
                    break
            if not converted:
                raise RuntimeError(
                    "ไม่พบ Microsoft Word หรือ LibreOffice\n"
                    "กรุณาติดตั้ง LibreOffice: https://www.libreoffice.org/download/libreoffice-fresh/"
                )

        with open(tmp_pdf, "rb") as f:
            return f.read()

    finally:
        for p in [tmp_docx, tmp_pdf]:
            if os.path.exists(p):
                os.unlink(p)
        os.rmdir(tmp_dir)

# ─── Helper: แสดง Step Indicator ──────────────────────────────────────────────
def show_steps(current: int):
    labels = ["📁 อัปโหลด Template", "✏️ กรอกข้อมูล", "⬇️ ดาวน์โหลด"]
    cols = st.columns(3)
    for i, (col, label) in enumerate(zip(cols, labels), start=1):
        with col:
            if i < current:
                st.success(f"✅  {label}")
            elif i == current:
                st.info(f"▶️  {label}")
            else:
                st.markdown(
                    f"<p style='text-align:center;color:#aaa;margin:0'>{label}</p>",
                    unsafe_allow_html=True,
                )

# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image(
        "https://img.icons8.com/fluency/96/word.png",
        width=64,
    )
    st.title("Word Template Filler")
    st.caption("v2.0 — สร้างเอกสาร Word จาก Template")
    st.divider()

    mode = st.radio(
        "เลือกโหมด",
        ["📝 Single — กรอกทีละเอกสาร", "📊 Batch — สร้างจาก CSV"],
        help="Single: กรอกข้อมูลเองทีละชุด | Batch: นำเข้า CSV สร้างหลายไฟล์พร้อมกัน",
    )

    st.divider()
    st.markdown("#### 📥 ตัวอย่าง Template")
    st.caption("หนังสือรับรองการทำงาน พร้อมตัวแปรครบ")
    st.download_button(
        label="ดาวน์โหลด Sample Template",
        data=create_sample_template(),
        file_name="sample_template.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

    st.divider()
    st.markdown("#### 📤 รูปแบบไฟล์ Output")
    output_format = st.radio(
        "เลือก format ที่ต้องการ",
        ["📄 DOCX", "📕 PDF", "📄+📕 DOCX และ PDF"],
        help="PDF ต้องการ Microsoft Word หรือ LibreOffice",
    )

    st.divider()
    with st.expander("📖 วิธีสร้าง Template"):
        st.markdown("""
ใส่ตัวแปรใน Word ด้วยรูปแบบนี้:

| ตัวแปร | Input ที่ได้ |
|--------|-------------|
| `{{ name }}` | Text |
| `{{ date }}` | Date Picker |
| `{{ email }}` | Email |
| `{{ phone }}` | เบอร์โทร |
| `{{ address }}` | Textarea |
| `{{ salary }}` | ตัวเลข |
        """)

# ─── Main ──────────────────────────────────────────────────────────────────────
st.title("📄 Word Template Filler")

# ══════════════════════════════════════════════════════════════════════════════
# MODE: SINGLE
# ══════════════════════════════════════════════════════════════════════════════
if "Single" in mode:

    show_steps(1)
    st.divider()

    uploaded_file = st.file_uploader(
        "อัปโหลดไฟล์ Word Template (.docx)",
        type=["docx"],
        help="ไฟล์ต้องมีตัวแปร {{ variable }} อยู่ภายใน",
    )

    if uploaded_file:
        try:
            template_bytes = uploaded_file.read()
            doc = DocxTemplate(io.BytesIO(template_bytes))
            variables = doc.get_undeclared_template_variables()

            if not variables:
                st.warning("⚠️ ไม่พบตัวแปรใดๆ ในเทมเพลตนี้ — ลองดาวน์โหลด Sample Template ใน sidebar")
            else:
                show_steps(2)
                st.divider()

                vars_list = sorted(variables)
                st.info(
                    f"พบ **{len(vars_list)}** ตัวแปร: "
                    + "  ".join([f"`{v}`" for v in vars_list])
                )

                with st.form("single_form"):
                    st.subheader("กรอกข้อมูลสำหรับตัวแปร")
                    context: dict = {}

                    # แบ่ง 2 คอลัมน์เมื่อมีตัวแปรมากกว่า 3 ตัว
                    if len(vars_list) > 3:
                        col1, col2 = st.columns(2)
                        for i, var in enumerate(vars_list):
                            label = var.replace("_", " ").title()
                            with col1 if i % 2 == 0 else col2:
                                context[var] = get_input_widget(var, label, key=f"s_{var}")
                    else:
                        for var in vars_list:
                            label = var.replace("_", " ").title()
                            context[var] = get_input_widget(var, label, key=f"s_{var}")

                    st.divider()
                    submitted = st.form_submit_button(
                        "🚀 สร้างเอกสาร", use_container_width=True, type="primary"
                    )

                if submitted:
                    empty = [k for k, v in context.items() if not str(v).strip()]
                    if empty:
                        st.error(
                            "❌ กรุณากรอกให้ครบ ยังขาด: "
                            + "  ".join([f"`{v}`" for v in empty])
                        )
                    else:
                        with st.spinner("กำลังสร้างเอกสาร..."):
                            docx_output = render_doc(template_bytes, context)

                        show_steps(3)
                        st.success("✅ สร้างเอกสารสำเร็จ!")

                        base_name = uploaded_file.name.replace(".docx", "_filled")
                        want_docx = "DOCX" in output_format
                        want_pdf  = "PDF"  in output_format

                        if want_docx:
                            st.download_button(
                                label="⬇️ ดาวน์โหลด DOCX",
                                data=docx_output,
                                file_name=f"{base_name}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                                type="primary",
                            )

                        if want_pdf:
                            with st.spinner("กำลังแปลงเป็น PDF..."):
                                try:
                                    pdf_output = convert_to_pdf(docx_output)
                                    st.download_button(
                                        label="⬇️ ดาวน์โหลด PDF",
                                        data=pdf_output,
                                        file_name=f"{base_name}.pdf",
                                        mime="application/pdf",
                                        use_container_width=True,
                                        type="primary",
                                    )
                                except RuntimeError as pdf_err:
                                    st.error(f"❌ แปลง PDF ไม่ได้: {pdf_err}")

        except Exception as e:
            st.error(f"❌ ไม่สามารถอ่านไฟล์ได้: {e}")

    else:
        st.markdown("### 👆 เริ่มต้นด้วยการอัปโหลดไฟล์ Template ด้านบน")
        st.markdown("หรือทดลองด้วย **Sample Template** จาก sidebar ทางซ้าย")


# ══════════════════════════════════════════════════════════════════════════════
# MODE: BATCH
# ══════════════════════════════════════════════════════════════════════════════
else:
    st.subheader("📊 Batch Generation — สร้างหลายเอกสารจาก CSV")
    st.caption("อัปโหลด Template + ไฟล์ CSV แล้วระบบจะสร้างเอกสารทีละ row อัตโนมัติ")
    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        uploaded_template = st.file_uploader("1️⃣ อัปโหลด Word Template (.docx)", type=["docx"])
    with col2:
        uploaded_csv = st.file_uploader(
            "2️⃣ อัปโหลดไฟล์ข้อมูล (.csv)",
            type=["csv"],
            help="Row แรกต้องเป็น header ชื่อตรงกับตัวแปรใน Template",
        )

    # ── ถ้าอัปโหลด Template แล้ว ให้ดาวน์โหลด CSV Template ──────────────────
    if uploaded_template and not uploaded_csv:
        try:
            tmpl_b = uploaded_template.read()
            doc = DocxTemplate(io.BytesIO(tmpl_b))
            variables = doc.get_undeclared_template_variables()
            if variables:
                sample_df = pd.DataFrame(columns=sorted(variables))
                st.info(
                    f"พบตัวแปร {len(variables)} ตัว — ดาวน์โหลด CSV Template แล้วกรอกข้อมูลก่อนอัปโหลด"
                )
                st.download_button(
                    "📥 ดาวน์โหลด CSV Template (header สำเร็จรูป)",
                    data=sample_df.to_csv(index=False),
                    file_name="data_template.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
                st.dataframe(sample_df, use_container_width=True)
            # เก็บ bytes ไว้ใน session เพราะ file_uploader reset เมื่อ rerun
            st.session_state["tmpl_bytes"] = tmpl_b
        except Exception as e:
            st.error(f"❌ อ่านไฟล์ Template ไม่ได้: {e}")

    # ── ถ้าอัปโหลดครบทั้งคู่ ─────────────────────────────────────────────────
    if uploaded_template and uploaded_csv:
        try:
            tmpl_b = uploaded_template.read()
            df = pd.read_csv(uploaded_csv, dtype=str).fillna("")

            st.success(f"✅ พบข้อมูล **{len(df)} แถว** และ **{len(df.columns)} คอลัมน์**")

            with st.expander("ดูตัวอย่างข้อมูล (5 แถวแรก)", expanded=True):
                st.dataframe(df.head(), use_container_width=True)

            st.divider()
            if st.button(
                f"🚀 สร้างเอกสารทั้งหมด {len(df)} ไฟล์",
                use_container_width=True,
                type="primary",
            ):
                progress_bar = st.progress(0, text="เริ่มต้น...")
                status_text = st.empty()
                zip_buf = io.BytesIO()

                want_docx = "DOCX" in output_format
                want_pdf  = "PDF"  in output_format

                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for i, row in df.iterrows():
                        context = {col: str(val) for col, val in row.items()}
                        doc_bytes = render_doc(tmpl_b, context)

                        first_col = df.columns[0]
                        safe_name = (
                            str(row[first_col])
                            .replace("/", "-")
                            .replace(" ", "_")
                            .replace("\\", "-")[:50]
                        )
                        base = f"{int(i)+1:03d}_{safe_name}"

                        if want_docx:
                            zf.writestr(f"{base}.docx", doc_bytes)
                        if want_pdf:
                            try:
                                pdf_bytes = convert_to_pdf(doc_bytes)
                                zf.writestr(f"{base}.pdf", pdf_bytes)
                            except RuntimeError:
                                zf.writestr(f"{base}.docx", doc_bytes)  # fallback

                        pct = (int(i) + 1) / len(df)
                        progress_bar.progress(pct, text=f"สร้างแล้ว {int(i)+1}/{len(df)}")
                        status_text.caption(f"→ {base}")

                zip_buf.seek(0)
                status_text.empty()
                st.success(f"✅ สร้างครบทั้ง {len(df)} เอกสาร!")

                st.download_button(
                    label=f"⬇️ ดาวน์โหลดทั้งหมด ({len(df)} ไฟล์) เป็น ZIP",
                    data=zip_buf,
                    file_name="generated_documents.zip",
                    mime="application/zip",
                    use_container_width=True,
                    type="primary",
                )

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")
