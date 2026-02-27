import streamlit as st
import pandas as pd
import random
import math
from io import BytesIO
from typing import List, Tuple
from docx import Document
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import pagesizes


# ==========================================================
# BUSINESS LOGIC LAYER
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
# EXPORT UTILITIES
# ==========================================================

def export_to_excel(df: pd.DataFrame) -> BytesIO:
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)
    return buffer


def export_to_word(df: pd.DataFrame) -> BytesIO:
    buffer = BytesIO()
    doc = Document()
    doc.add_heading("Duty Allocation List", level=1)

    table = doc.add_table(rows=df.shape[0] + 1, cols=df.shape[1])

    for col_num, column in enumerate(df.columns):
        table.rows[0].cells[col_num].text = str(column)

    for row_num in range(df.shape[0]):
        for col_num in range(df.shape[1]):
            table.rows[row_num + 1].cells[col_num].text = str(df.iat[row_num, col_num])

    doc.save(buffer)
    buffer.seek(0)
    return buffer


def export_to_pdf(df: pd.DataFrame) -> BytesIO:
    buffer = BytesIO()
    pdf = SimpleDocTemplate(buffer, pagesize=pagesizes.A4)
    elements = []

    style = getSampleStyleSheet()
    elements.append(Paragraph("Duty Allocation List", style["Heading1"]))
    elements.append(Spacer(1, 12))

    data = [df.columns.tolist()] + df.values.tolist()

    table = Table(data)
    table.setStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])

    elements.append(table)
    pdf.build(elements)

    buffer.seek(0)
    return buffer


# ==========================================================
# STREAMLIT UI LAYER
# ==========================================================

st.set_page_config(
    page_title="AI Duty Allocation Agent",
    layout="wide"
)

st.title("🤖 AI-Based Examination Duty Allocation System")

# -------------------------
# Slot Configuration
# -------------------------

num_slots = st.number_input(
    "Number of Examination Slots",
    min_value=1,
    max_value=12,
    step=1
)

slot_timings = []

if num_slots > 0:
    st.subheader("⏰ Slot Timing Configuration")

    for i in range(int(num_slots)):
        col1, col2 = st.columns(2)

        with col1:
            start_time = st.text_input(
                f"Slot {i+1} Start Time",
                key=f"start_{i}"
            )

        with col2:
            end_time = st.text_input(
                f"Slot {i+1} End Time",
                key=f"end_{i}"
            )

        if start_time and end_time:
            slot_timings.append(f"{start_time} - {end_time}")
        else:
            slot_timings.append("")


# -------------------------
# Teacher Input
# -------------------------

st.subheader("👨‍🏫 Teacher Data Input")

teacher_file = st.file_uploader(
    "Upload Teacher List (CSV / Excel)",
    type=["csv", "xlsx"]
)

manual_teacher_input = st.text_area(
    "OR Enter Teacher Names (comma separated)"
)

room_input = st.text_area("🏫 Enter Room Names (comma separated)")


# ==========================================================
# MAIN EXECUTION
# ==========================================================

if st.button("🚀 Generate Duty Allocation"):

    teachers = []

    # File input
    if teacher_file is not None:
        try:
            if teacher_file.name.endswith(".csv"):
                teacher_df = pd.read_csv(teacher_file)
            else:
                teacher_df = pd.read_excel(teacher_file)

            teachers = teacher_df.iloc[:, 0].dropna().astype(str).tolist()

        except Exception:
            st.error("Unable to read teacher file.")

    # Manual input
    elif manual_teacher_input:
        teachers = [
            t.strip()
            for t in manual_teacher_input.split(",")
            if t.strip()
        ]

    rooms = [r.strip() for r in room_input.split(",") if r.strip()]

    if not teachers:
        st.error("Teacher data is required.")
        st.stop()

    if not rooms:
        st.error("Room data is required.")
        st.stop()

    if len(teachers) < len(rooms):
        st.warning(
            "Number of teachers is less than number of rooms. "
            "Some allocations may show 'No Available Teacher'."
        )

    result, max_duties = generate_duty_list(
        rooms,
        teachers,
        int(num_slots)
    )

    df = pd.DataFrame(result).T

    df.columns = [
        f"Slot {i+1} ({slot_timings[i]})"
        if slot_timings[i]
        else f"Slot {i+1}"
        for i in range(int(num_slots))
    ]

    df.insert(0, "Room/Class", df.index)
    df.insert(0, "S. No", range(1, len(df) + 1))
    df.reset_index(drop=True, inplace=True)

    st.success(f"Maximum Duties per Teacher (Auto-Calculated): {max_duties}")
    st.dataframe(df, use_container_width=True)

    # Export Buttons
    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            "📥 Excel",
            data=export_to_excel(df),
            file_name="Duty_List.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col2:
        st.download_button(
            "📄 Word",
            data=export_to_word(df),
            file_name="Duty_List.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    with col3:
        st.download_button(
            "📑 PDF",
            data=export_to_pdf(df),
            file_name="Duty_List.pdf",
            mime="application/pdf"
        )