import random
import math

def generate_duty_list(rooms, teachers, slots):

    total_duties = len(rooms) * slots
    total_teachers = len(teachers)

    if total_teachers == 0:
        return {}, 0

    max_duties = math.ceil(total_duties / total_teachers)

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