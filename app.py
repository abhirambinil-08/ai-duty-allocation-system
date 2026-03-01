import streamlit as st
import pandas as pd
import random
import math
from io import BytesIO
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from openpyxl.utils import get_column_letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4


# ==========================================================
# BUSINESS LOGIC
# ==========================================================

def calculate_max_duties(num_rooms, num_slots, num_teachers):
    if num_teachers == 0:
        return 0
    return math.ceil((num_rooms * num_slots) / num_teachers)


def generate_slot_duty(rooms, teachers, slots):
    max_duties = calculate_max_duties(len(rooms), slots, len(teachers))

    duty_table = {}
    teacher_count = {t: 0 for t in teachers}
    slot_assignments = {slot: [] for slot in range(slots)}

    for room in rooms:
        duty_table[room] = []
        for slot in range(slots):
            available = [
                t for t in teachers
                if teacher_count[t] < max_duties
                and t not in slot_assignments[slot]
                and t not in duty_table[room]
            ]

            if not available:
                duty_table[room].append("No Available Teacher")
            else:
                selected = random.choice(available)
                duty_table[room].append(selected)
                teacher_count[selected] += 1
                slot_assignments[slot].append(selected)

    return duty_table, max_duties


def generate_centre_duty(num_rooms, teachers):
    random.shuffle(teachers)
    data = []
    index = 0

    for room in range(1, num_rooms + 1):
        inv1 = teachers[index]
        inv2 = teachers[index + 1]
        index += 2
        data.append([room, inv1, "", inv2, ""])

    return pd.DataFrame(
        data,
        columns=["ROOM", "INVIGILATOR 1", "SIGN 1", "INVIGILATOR 2", "SIGN 2"]
    )


# ==========================================================
# EXPORT FUNCTIONS
# ==========================================================

def export_excel(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Duty List")
        worksheet = writer.sheets["Duty List"]

        for i, column in enumerate(df.columns, 1):
            max_length = max(df[column].astype(str).map(len).max(), len(column))
            worksheet.column_dimensions[get_column_letter(i)].width = max_length + 3

    buffer.seek(0)
    return buffer


# =================== WORD (CBSE STYLE) ===================

def export_word_cbse(df, school_name, school_address, centre_no, duty_date):

    buffer = BytesIO()
    doc = Document()

    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # Header
    p = doc.add_paragraph()
    run = p.add_run(school_name.upper())
    run.bold = True
    run.font.size = Pt(14)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph(school_address.upper())
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph(f"CENTRE NO.: {centre_no}")
    p.runs[0].bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph(f"DUTY LIST FOR {duty_date}")
    p.runs[0].bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(" ")

    # Table
    table = doc.add_table(rows=df.shape[0] + 2, cols=5)
    table.style = "Table Grid"

    table.rows[0].cells[0].text = "ROOM"
    table.rows[0].cells[1].text = "1"
    table.rows[0].cells[2].text = "SIGN"
    table.rows[0].cells[3].text = "2"
    table.rows[0].cells[4].text = "SIGN"

    table.rows[1].cells[1].merge(table.rows[1].cells[3])
    table.rows[1].cells[1].text = "NAME OF INVIGILATOR"

    for i in range(df.shape[0]):
        table.rows[i+2].cells[0].text = str(df.iloc[i, 0])
        table.rows[i+2].cells[1].text = str(df.iloc[i, 1])
        table.rows[i+2].cells[3].text = str(df.iloc[i, 3])

    doc.add_paragraph("\nStaff on frisking duty :")
    doc.add_paragraph("From 09:35 am to 10:00 am")

    doc.add_paragraph("\nAns. books deposit to CBSE: ____________________")

    doc.add_paragraph(
        "\nWitness duty at 09:50 am for opening of sealed question paper packets in Principal’s Room:"
    )
    doc.add_paragraph("______________________________________")

    doc.add_paragraph("\nCentre Visit Duty at 08:45 am:")
    doc.add_paragraph("______________________________________")

    doc.add_paragraph("\n")
    doc.add_paragraph(str(duty_date))

    p = doc.add_paragraph("PRINCIPAL")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.save(buffer)
    buffer.seek(0)
    return buffer


# =================== PDF (CBSE STYLE) ===================

def export_pdf_cbse(df, school_name, school_address, centre_no, duty_date):

    buffer = BytesIO()
    pdf = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []
    style = getSampleStyleSheet()

    elements.append(Paragraph(f"<b>{school_name.upper()}</b>", style["Title"]))
    elements.append(Paragraph(school_address.upper(), style["Normal"]))
    elements.append(Paragraph(f"<b>CENTRE NO.: {centre_no}</b>", style["Normal"]))
    elements.append(Paragraph(f"<b>DUTY LIST FOR {duty_date}</b>", style["Normal"]))
    elements.append(Spacer(1, 12))

    data = [df.columns.tolist()] + df.values.tolist()
    table = Table(data, repeatRows=1)

    table.setStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
    ])

    elements.append(table)
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("PRINCIPAL", style["Normal"]))

    pdf.build(elements)
    buffer.seek(0)
    return buffer


# ==========================================================
# STREAMLIT UI
# ==========================================================

st.set_page_config(layout="wide")
st.title("AI-Based Examination Duty Allocation System")

page = st.radio("Select Format", ["Slot-Based Plan", "CBSE Centre Sheet"])


# ================= SLOT PLAN =================

if page == "Slot-Based Plan":

    num_slots = st.number_input("Number of Slots", min_value=1, max_value=10)
    teachers_input = st.text_area("Teacher Names (comma separated)")
    rooms_input = st.text_area("Room Names (comma separated)")

    if st.button("Generate Slot Plan"):

        teachers = [t.strip() for t in teachers_input.split(",") if t.strip()]
        rooms = [r.strip() for r in rooms_input.split(",") if r.strip()]

        result, max_duty = generate_slot_duty(rooms, teachers, num_slots)

        df = pd.DataFrame(result).T
        df.columns = [f"Slot {i+1}" for i in range(num_slots)]
        df.insert(0, "Room", df.index)
        df.insert(0, "S.No", range(1, len(df) + 1))
        df.reset_index(drop=True, inplace=True)

        st.success(f"Max Duties per Teacher: {max_duty}")
        st.dataframe(df)


# ================= CBSE CENTRE SHEET =================

if page == "CBSE Centre Sheet":

    school_name = st.text_input("School Name")
    school_address = st.text_input("School Address")
    centre_no = st.text_input("Centre Number")
    duty_date = st.date_input("Duty Date")

    num_rooms = st.number_input("Number of Rooms", min_value=1, max_value=50)
    teachers_input = st.text_area("Teacher Names (comma separated)")

    if st.button("Generate Centre Sheet"):

        teachers = [t.strip() for t in teachers_input.split(",") if t.strip()]

        if len(teachers) < num_rooms * 2:
            st.error("Need at least 2 teachers per room.")
            st.stop()

        df = generate_centre_duty(num_rooms, teachers)
        st.dataframe(df)

        st.download_button(
            "Download Excel",
            export_excel(df),
            "Centre_Duty.xlsx"
        )

        st.download_button(
            "Download Word (CBSE Format)",
            export_word_cbse(df, school_name, school_address, centre_no, duty_date),
            "Centre_Duty.docx"
        )

        st.download_button(
            "Download PDF (CBSE Format)",
            export_pdf_cbse(df, school_name, school_address, centre_no, duty_date),
            "Centre_Duty.pdf"
        )