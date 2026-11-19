import streamlit as st
from google import genai
from google.genai import types
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import datetime
import tempfile
import os
import json

st.set_page_config(page_title="ISO Meeting Minutes AI", page_icon="📋", layout="wide")

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if not st.session_state["authenticated"]:
        st.title("🔒 Private ISO Meeting Minutes Portal")
        st.write("Please enter the company access password to continue.")
        pwd = st.text_input("Company Password", type="password")
        if st.button("Login"):
            if pwd == st.secrets["APP_PASSWORD"]:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("❌ Incorrect password. Access denied.")
        return False
    return True

if not check_password():
    st.stop()

client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

def generate_iso_docx(meta, ai_data):
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    header_table = doc.add_table(rows=2, cols=2)
    header_table.style = 'Table Grid'
    header_table.cell(0, 0).text = "DOCUMENT TITLE:\nMINUTES OF THE MEETING (ISO COMPLIANT)"
    header_table.cell(0, 1).text = f"Document Control No.:\n{meta['doc_no']}"
    header_table.cell(1, 0).text = f"Meeting Title:\n{meta['title']}"
    header_table.cell(1, 1).text = f"Date & Venue:\n{meta['date']} | {meta['venue']}"

    for row in header_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    run.font.bold = True

    doc.add_paragraph()

    p_fac = doc.add_paragraph()
    p_fac.add_run("Facilitator / Operations Manager: ").bold = True
    p_fac.add_run(meta['facilitator'])

    h1 = doc.add_heading("PART I: FORMAL MINUTES OF THE MEETING", level=1)
    h1.alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.add_heading("1. Agenda / Topics Discussed", level=2)
    for item in ai_data.get("agenda", []):
        doc.add_paragraph(item, style='List Bullet')

    doc.add_heading("2. Detailed Summary of Discussions & Key Decisions", level=2)
    for summary in ai_data.get("summary", []):
        doc.add_paragraph(summary, style='List Bullet')

    doc.add_heading("3. Action Plan & Responsibilities", level=2)
    action_items = ai_data.get("action_items", [])
    if action_items:
        act_table = doc.add_table(rows=1, cols=4)
        act_table.style = 'Table Grid'
        hdr_cells = act_table.rows.cells
        hdr_cells.text = "Action Item / Task"
        hdr_cells.text = "Person Responsible"
        hdr_cells.text = "Target Date"
        hdr_cells.text = "Status"
        
        for cell in hdr_cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.bold = True

        for act in action_items:
            row_cells = act_table.add_row().cells
            row_cells.text = str(act.get("task", ""))
            row_cells.text = str(act.get("owner", ""))
            row_cells.text = str(act.get("deadline", "TBD"))
            row_cells.text = str(act.get("status", "Open"))
    else:
        doc.add_paragraph("No specific action items recorded.")

    doc.add_paragraph()

    doc.add_heading("4. Attendance & Verification (Wet Signature Required)", level=2)
    att_table = doc.add_table(rows=1, cols=3)
    att_table.style = 'Table Grid'
    att_hdr = att_table.rows.cells
    att_hdr.text = "Name of Attendee"
    att_hdr.text = "Department / Role"
    att_hdr.text = "Signature"

    for cell in att_hdr:
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.bold = True

    attendees_list = [a.strip() for a in meta['attendees'].split("\n") if a.strip()]
    for person in attendees_list:
        row_cells = att_table.add_row().cells
        if "-" in person:
            parts = person.split("-", 1)
            row_cells.text = parts.strip()
            row_cells.text = parts.strip()
        else:
            row_cells.text = person
            row_cells.text = ""
        row_cells.text = "\n\n"

    doc.add_paragraph()
    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.style = 'Table Grid'
    sig_table.cell(0, 0).text = "Prepared By (Secretary / Admin):"
    sig_table.cell(0, 1).text = "Reviewed & Approved By (Facilitator / OM):"
    sig_table.cell(1, 0).text = "\n\n___________________________\nSignature over Printed Name / Date"
    sig_table.cell(1, 1).text = f"\n\n___________________________\n{meta['facilitator']}\nSignature over Printed Name / Date"

    doc.add_page_break()

    doc.add_heading("PART II: WORD-FOR-WORD MEETING TRANSCRIPT", level=1)
    doc.add_paragraph("Note: Transcribed verbatim in original spoken language (Taglish/English/Tagalog) with speaker attribution.")
    
    transcript_text = ai_data.get("transcript", "No transcript generated.")
    for line in transcript_text.split("\n"):
        if line.strip():
            p = doc.add_paragraph()
            if ":" in line:
                speaker, speech = line.split(":", 1)
                p.add_run(speaker + ":").bold = True
                p.add_run(speech)
            else:
                p.add_run(line)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.title("📋 Private ISO Meeting Minutes AI Generator")
st.markdown("Generate audit-ready **ISO Minutes of the Meeting** + **Verbatim Taglish Transcript** in one click.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. ISO Document Details")
    doc_no = st.text_input("Document Control No.", value="ISO-MOM-2026-001")
    meeting_title = st.text_input("Meeting Title", value="Weekly Operations & Quality Alignment")
    meeting_date = st.date_input("Date", value=datetime.date.today())
    venue = st.text_input("Venue / Platform", value="Main Conference Room / Online")
    facilitator = st.text_input("Facilitator / Operations Manager", value="")
    
    st.markdown("**Attendees List** *(Format: Name - Role, one per line)*")
    attendees = st.text_area("Attendees", height=150, placeholder="Juan Dela Cruz - Operations Manager\nMaria Santos - QA Lead\nMark Reyes - Production Supervisor")

with col2:
    st.subheader("2. Meeting Audio Input")
    st.info("💡 **Tip:** Simulan ang recording sa mabilis na roll-call (*'Juan here, Maria here...'*) para 100% kilala ng AI kung sino ang nagsasalita.")
    
    audio_source = st.radio("Choose Audio Input Method:", ["🎙️ Record Live Microphone", "📁 Upload Audio File (MP3, WAV, M4A, OGG)"])
    
    audio_bytes = None
    file_ext = ".wav"

    if audio_source == "🎙️ Record Live Microphone":
        recorded_audio = st.audio_input("Click the mic icon to start/stop recording")
        if recorded_audio:
            audio_bytes = recorded_audio.read()
            file_ext = ".wav"
    else:
        uploaded_file = st.file_uploader("Upload Meeting Recording", type=["mp3", "wav", "m4a", "ogg", "aac"])
        if uploaded_file:
            audio_bytes = uploaded_file.read()
            file_ext = os.path.splitext(uploaded_file.name)

    st.markdown("---")
    generate_btn = st.button("🚀 Generate ISO Minutes & Word File", type="primary", use_container_width=True)

if generate_btn:
    if not audio_bytes:
        st.warning("⚠️ Please record or upload an audio file first.")
    elif not attendees.strip():
        st.warning("⚠️ Please enter the Attendees list for speaker recognition and the signature sheet.")
    else:
        with st.spinner("⏳ Analyzing voices & generating detailed ISO minutes..."):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_file:
                    tmp_file.write(audio_bytes)
                    tmp_path = tmp_file.name

                uploaded_gemini_file = client.files.upload(file=tmp_path)

                prompt = (
                    "You are an executive-level ISO 9001 Document Controller and Corporate Secretary in the Philippines. "
                    "Listen carefully to the entire meeting audio. The speakers use Taglish (Tagalog and English). "
                    f"Official Attendees List: {attendees}. Facilitator/OM: {facilitator}. "
                    "Return valid JSON with 4 keys: "
                    "1. 'agenda': list of strings of main topics discussed in formal English. "
                    "2. 'summary': thorough list of strings of key discussions, numbers, agreements, and decisions in formal English. "
                    "3. 'action_items': list of objects with keys 'task', 'owner', 'deadline', and 'status' ('Open'). "
                    "4. 'transcript': complete word-for-word verbatim transcript in original Taglish with speaker attribution based on voice and roll-call."
                )

                response = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=[uploaded_gemini_file, prompt],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.2
                    )
                )

                os.remove(tmp_path)
                client.files.delete(name=uploaded_gemini_file.name)

                ai_result = json.loads(response.text)

                meta_info = {
                    "doc_no": doc_no,
                    "title": meeting_title,
                    "date": meeting_date.strftime("%B %d, %Y"),
                    "venue": venue,
                    "facilitator": facilitator,
                    "attendees": attendees
                }
                docx_file = generate_iso_docx(meta_info, ai_result)

                st.success("✅ ISO Meeting Minutes successfully generated!")

                st.download_button(
                    label="📥 Download Official ISO Minutes (.docx)",
                    data=docx_file,
                    file_name=f"{doc_no}_{meeting_date.strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )

                tab1, tab2 = st.tabs(["📄 Formal ISO Minutes Preview", "🗣️ Word-for-Word Taglish Transcript"])
                with tab1:
                    st.subheader("Agenda")
                    for a in ai_result.get("agenda", []):
                        st.write(f"- {a}")
                    st.subheader("Detailed Discussion Summary")
                    for s in ai_result.get("summary", []):
                        st.write(f"- {s}")
                    st.subheader("Action Plan")
                    st.table(ai_result.get("action_items", []))
                with tab2:
                    st.text_area("Verbatim Transcript", value=ai_result.get("transcript", ""), height=400)

            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
