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

def calculate_max_duties(num_rooms: int, num_slots: int, num_teachers: int) -> int:
    if num_teachers == 0:
        return 0
    total_duties = num_rooms * num_slots
    return math.ceil(total_duties / num_teachers)


def generate_duty_list(
    rooms: List[str],
    teachers: List[str],
    slots: int
) -> Tuple[dict, int]:

    max_duties = calculate_max_duties(len(rooms), slots, len(teachers))

    duty_table = {}
    teacher_count = {teacher: 0 for teacher in teachers}
    slot_assignments = {slot: [] for slot in range(slots)}

    for room in rooms:
        duty_table[room] = []

        for slot in range(slots):
            available_teachers = [
                t for t in teachers
                if teacher_count[t] < max_duties
                and t not in slot_assignments[slot]
                and t not in duty_table[room]
            ]

            if not available_teachers:
                duty_table[room].append("No Available Teacher")
            else:
                selected = random.choice(available_teachers)
                duty_table[room].append(selected)
                teacher_count[selected] += 1
                slot_assignments[slot].append(selected)

    return duty_table, max_duties


# ==========================================================
# EXPORT FUNCTIONS
# ==========================================================

def export_to_excel(df: pd.DataFrame) -> BytesIO:
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Duty List")
        worksheet = writer.sheets["Duty List"]

        for col_idx, column in enumerate(df.columns, 1):
            max_length = max(
                df[column].astype(str).map(len).max(),
                len(column)
            )
            worksheet.column_dimensions[get_column_letter(col_idx)].width = max_length + 2

    buffer.seek(0)
    return buffer


def export_to_word(df: pd.DataFrame) -> BytesIO:
    buffer = BytesIO()
    doc = Document()

    # Landscape
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    # Header Centered
    if school_name:
        p = doc.add_heading(school_name.upper(), level=0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if school_address:
        p = doc.add_paragraph(school_address)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    title = f"Duty List for {duty_date} – {exam_title}"
    p = doc.add_paragraph(title)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(" ")

    table = doc.add_table(rows=df.shape[0] + 1, cols=df.shape[1])
    table.style = "Table Grid"

    for col_num, column in enumerate(df.columns):
        table.rows[0].cells[col_num].text = str(column)

    for row_num in range(df.shape[0]):
        for col_num in range(df.shape[1]):
            table.rows[row_num + 1].cells[col_num].text = str(df.iat[row_num, col_num])

    doc.add_paragraph("\n")
    doc.add_paragraph("PRINCIPAL")
    doc.add_paragraph(f"Generated on: {pd.Timestamp.now().date()}")

    doc.save(buffer)
    buffer.seek(0)
    return buffer


def export_to_pdf(df: pd.DataFrame) -> BytesIO:
    buffer = BytesIO()

    pdf = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20,
    )

    elements = []
    style = getSampleStyleSheet()

    if school_name:
        elements.append(Paragraph(f"<b>{school_name.upper()}</b>", style["Title"]))
        elements.append(Spacer(1, 6))

    if school_address:
        elements.append(Paragraph(school_address, style["Normal"]))
        elements.append(Spacer(1, 6))

    elements.append(
        Paragraph(
            f"Duty List for {duty_date} – {exam_title}",
            style["Heading2"]
        )
    )
    elements.append(Spacer(1, 12))

    data = [df.columns.tolist()] + df.values.tolist()

    page_width = landscape(A4)[0] - 40
    col_width = page_width / len(df.columns)
    col_widths = [col_width] * len(df.columns)

    table = Table(data, colWidths=col_widths, repeatRows=1)

    table.setStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ])

    elements.append(table)
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("PRINCIPAL", style["Normal"]))
    elements.append(Paragraph(f"Generated on: {pd.Timestamp.now().date()}", style["Normal"]))

    pdf.build(elements)
    buffer.seek(0)
    return buffer


# ==========================================================
# STREAMLIT UI
# ==========================================================

st.set_page_config(page_title="AI Duty Allocation Agent", layout="wide")
st.title("🤖 AI-Based Examination Duty Allocation System")

# Institution Info
st.subheader("🏫 Institution Information")

school_name = st.text_input("School Name")
school_address = st.text_input("School Address")
exam_title = st.text_input("Examination Title")
duty_date = st.date_input("Duty Date")
principal_name = st.text_input("Principal Name")

# Slot Config
num_slots = st.number_input("Number of Examination Slots", min_value=1, max_value=12)

slot_timings = []

if num_slots > 0:
    st.subheader("⏰ Slot Timing Configuration")

    for i in range(int(num_slots)):
        col1, col2 = st.columns(2)

        with col1:
            start_time = st.text_input(f"Slot {i+1} Start Time", key=f"start_{i}")
        with col2:
            end_time = st.text_input(f"Slot {i+1} End Time", key=f"end_{i}")

        if start_time and end_time:
            slot_timings.append(f"{start_time} - {end_time}")
        else:
            slot_timings.append("")

# Teacher Input
st.subheader("👨‍🏫 Teacher Data Input")

teacher_file = st.file_uploader("Upload Teacher List (CSV / Excel)", type=["csv", "xlsx"])
manual_teacher_input = st.text_area("OR Enter Teacher Names (comma separated)")
room_input = st.text_area("🏫 Enter Room Names (comma separated)")

if st.button("🚀 Generate Duty Allocation"):

    teachers = []

    if teacher_file is not None:
        if teacher_file.name.endswith(".csv"):
            teacher_df = pd.read_csv(teacher_file)
        else:
            teacher_df = pd.read_excel(teacher_file)
        teachers = teacher_df.iloc[:, 0].dropna().astype(str).tolist()

    elif manual_teacher_input:
        teachers = [t.strip() for t in manual_teacher_input.split(",") if t.strip()]

    rooms = [r.strip() for r in room_input.split(",") if r.strip()]

    if not teachers:
        st.error("Teacher data is required.")
        st.stop()

    if not rooms:
        st.error("Room data is required.")
        st.stop()

    result, max_duties = generate_duty_list(rooms, teachers, int(num_slots))

    df = pd.DataFrame(result).T

    df.columns = [
        f"Slot {i+1} ({slot_timings[i]})" if slot_timings[i]
        else f"Slot {i+1}"
        for i in range(int(num_slots))
    ]

    df.insert(0, "Room/Class", df.index)
    df.insert(0, "S. No", range(1, len(df) + 1))
    df.reset_index(drop=True, inplace=True)

    st.success(f"Maximum Duties per Teacher (Auto-Calculated): {max_duties}")
    st.dataframe(df, width="stretch")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button("📥 Excel", export_to_excel(df), "Duty_List.xlsx")

    with col2:
        st.download_button("📄 Word", export_to_word(df), "Duty_List.docx")

    with col3:
        st.download_button("📑 PDF", export_to_pdf(df), "Duty_List.pdf")