import streamlit as st
import pandas as pd
import random
import math
from io import BytesIO
from typing import List, Tuple
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl.utils import get_column_letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4, landscape


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
            max_length = max(
                df[column].astype(str).map(len).max(),
                len(column)
            )
            worksheet.column_dimensions[get_column_letter(i)].width = max_length + 3

    buffer.seek(0)
    return buffer


def export_word(df, header_lines):
    buffer = BytesIO()
    doc = Document()

    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    for line in header_lines:
        p = doc.add_paragraph(line)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(" ")

    table = doc.add_table(rows=df.shape[0] + 1, cols=df.shape[1])
    table.style = "Table Grid"

    for col in range(df.shape[1]):
        table.rows[0].cells[col].text = str(df.columns[col])

    for row in range(df.shape[0]):
        for col in range(df.shape[1]):
            table.rows[row + 1].cells[col].text = str(df.iat[row, col])

    doc.add_paragraph("\nPRINCIPAL")
    doc.add_paragraph(f"Generated on: {pd.Timestamp.now().date()}")

    doc.save(buffer)
    buffer.seek(0)
    return buffer


def export_pdf(df, header_lines):
    buffer = BytesIO()

    pdf = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    elements = []
    style = getSampleStyleSheet()

    for line in header_lines:
        elements.append(Paragraph(f"<b>{line}</b>", style["Title"]))
        elements.append(Spacer(1, 6))

    elements.append(Spacer(1, 12))

    data = [df.columns.tolist()] + df.values.tolist()
    col_width = (landscape(A4)[0] - 40) / len(df.columns)
    table = Table(data, colWidths=[col_width] * len(df.columns), repeatRows=1)

    table.setStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
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
st.title("🤖 AI-Based Examination Duty Allocation System")

page = st.radio("Select Duty Format", ["Slot-Based Duty Plan", "Centre Invigilator Format"])


# ==========================================================
# SLOT-BASED DUTY PLAN
# ==========================================================

if page == "Slot-Based Duty Plan":

    school_name = st.text_input("School Name")
    school_address = st.text_input("School Address")
    exam_title = st.text_input("Exam Title")
    duty_date = st.date_input("Duty Date")

    num_slots = st.number_input("Number of Slots", min_value=1, max_value=10)

    slot_timings = []
    for i in range(num_slots):
        start = st.text_input(f"Slot {i+1} Start", key=f"s{i}")
        end = st.text_input(f"Slot {i+1} End", key=f"e{i}")
        slot_timings.append(f"{start}-{end}" if start and end else "")

    teachers_input = st.text_area("Teacher Names (comma separated)")
    rooms_input = st.text_area("Room Names (comma separated)")

    if st.button("Generate Slot Duty Plan"):

        teachers = [t.strip() for t in teachers_input.split(",") if t.strip()]
        rooms = [r.strip() for r in rooms_input.split(",") if r.strip()]

        result, max_duty = generate_slot_duty(rooms, teachers, num_slots)

        df = pd.DataFrame(result).T
        df.columns = [f"Slot {i+1} ({slot_timings[i]})" for i in range(num_slots)]
        df.insert(0, "Room", df.index)
        df.insert(0, "S.No", range(1, len(df) + 1))
        df.reset_index(drop=True, inplace=True)

        st.success(f"Max Duties per Teacher: {max_duty}")
        st.dataframe(df, width="stretch")

        header = [
            school_name,
            school_address,
            f"Duty List for {duty_date} – {exam_title}"
        ]

        st.download_button("📥 Excel", export_excel(df), "Slot_Duty.xlsx")
        st.download_button("📄 Word", export_word(df, header), "Slot_Duty.docx")
        st.download_button("📑 PDF", export_pdf(df, header), "Slot_Duty.pdf")


# ==========================================================
# CENTRE INVIGILATOR FORMAT
# ==========================================================

if page == "Centre Invigilator Format":

    school_name = st.text_input("School Name")
    school_address = st.text_input("School Address")
    centre_no = st.text_input("Centre Number")
    duty_date = st.date_input("Duty Date")

    num_rooms = st.number_input("Number of Rooms", min_value=1, max_value=50)
    teachers_input = st.text_area("Teacher Names (comma separated)")

    if st.button("Generate Centre Duty Sheet"):

        teachers = [t.strip() for t in teachers_input.split(",") if t.strip()]

        if len(teachers) < num_rooms * 2:
            st.error("Need at least 2 teachers per room.")
            st.stop()

        df = generate_centre_duty(num_rooms, teachers)

        st.dataframe(df, width="stretch")

        header = [
            school_name,
            school_address,
            f"CENTRE NO: {centre_no}",
            f"DUTY LIST FOR {duty_date}"
        ]

        st.download_button("📥 Excel", export_excel(df), "Centre_Duty.xlsx")
        st.download_button("📄 Word", export_word(df, header), "Centre_Duty.docx")
        st.download_button("📑 PDF", export_pdf(df, header), "Centre_Duty.pdf")