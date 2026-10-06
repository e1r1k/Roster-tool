import json
import calendar
import pulp
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import time

start_time = time.time()

file_path = "availability.json"

with open("availability.json") as f:
    availability_data = json.load(f)
employees = list(availability_data.keys())

with open("preferences.json") as f:
    preferences_data = json.load(f)

with open("hours.json") as f:
    hours_data = json.load(f)


# ---------------------------------------------------------------------------------------------------- #
# Preparing constants, setting up the month so that each day gets the correct shifts (weekends/weekdays)

year_inp, month_inp = 2026, 11
month = calendar.monthrange(year_inp, month_inp)  # (weekday_of_first, days_in_month)

# Required shifts by weekday
required_shifts = {
    0: ["V1", "V2", "V3"],
    1: ["V1", "V2", "V3"],
    2: ["V1", "V2", "V3"],
    3: ["V1", "V2", "V3"],
    4: ["V1", "V2", "V3", "V4"],
    5: ["V5", "V6", "V7", "V8"],
    6: ["V5", "V6", "V7", "V8"]
}

shift_hours_wdonly = {"V1":8, "V2":8, "V3":8, "V4":7, "V5":7, "V6":7, "V7":7, "V8":7}

# Build list of all shifts
shifts = []
total_hours_needed = 0
for day_nr in range(month[1]):  # days in month
    weekday = calendar.weekday(year_inp, month_inp, day_nr + 1)
    date_str = f"{year_inp}-{month_inp:02d}-{day_nr+1:02d}"
    for slot in required_shifts[weekday]:
        shifts.append((date_str, slot))  # unique ID: (date, shift_label)
        total_hours_needed += shift_hours_wdonly[slot]

print("Total hours needed: ", str(total_hours_needed))
total_hours_signed = 0
for e in hours_data:
    total_hours_signed += int(hours_data[e])
print("Total hours signed: " + str(total_hours_signed))


# ---------------------------- Building Constraints -----------------------------#

unavailable = {e: set() for e in availability_data.keys()}

for e, days in availability_data.items():
    for d, info in days.items():
        if info["unavailable"]:  # whole day blocked
            for shift in [s for (date_str, s) in shifts if date_str == d]:
                unavailable[e].add((d, shift))
        else:  # specific slots blocked
            for s in info["slots"]:
                unavailable[e].add((d, s))

preferences = {}
for e, prefs in preferences_data.items():
    weights = {}
    preferred = prefs["preferredshifts"].split(",")
    for (d, s) in shifts:
        weights[(d, s)] = 10 if s in preferred else 0
    preferences[e] = weights

min_hours = {}
max_hours = {}

for e, h in hours_data.items():
    h = int(h)
    min_hours[e] = int(0.9 * h)   # allows 10% slack
    max_hours[e] = int(1.1 * h)

shift_hours = {}
for (d, s) in shifts:
    if s in ["V1", "V2", "V3"]:
        shift_hours[(d, s)] = 8
    else:
        shift_hours[(d, s)] = 7


# ------------------------ Defining the LP problem   ----------------------------------- #
month_days = calendar.monthrange(year_inp, month_inp)[1]

x = pulp.LpVariable.dicts("assign",
                          (employees, shifts),
                          lowBound=0, upBound=1,
                          cat=pulp.LpBinary)

model = pulp.LpProblem("Roster", pulp.LpMaximize)

# Cover every shift exactly once
for (d, s) in shifts:
    model += pulp.lpSum(x[e][(d, s)] for e in employees) == 1

# Respect unavailability
for e in employees:
    for (d, s) in unavailable[e]:
        model += x[e][(d, s)] == 0

# Min/max hours per employee
for e in employees:
    model += pulp.lpSum(shift_hours[(d, s)] * x[e][(d, s)] for (d, s) in shifts) >= min_hours[e]
    model += pulp.lpSum(shift_hours[(d, s)] * x[e][(d, s)] for (d, s) in shifts) <= max_hours[e]

for e in employees:
    for day in range(1, month_days + 1):
        day_str = f"{year_inp}-{month_inp:02d}-{day:02d}"
        model += pulp.lpSum(x[e][(day_str, s)] for (d, s) in shifts if d == day_str) <= 1


# Ensure no employee is assigned more than 5 shifts per week
for e in employees:
    max_perweek = int(preferences_data[e]["max_perweek"])
    for week in range(1, 6):  # up to 5 weeks in a month
        week_days = [d for d in range(1, month_days+1)
                     if calendar.weekday(year_inp, month_inp, d) in range(7)
                     and (d-1)//7 + 1 == week]  # crude week grouping
        model += pulp.lpSum(
            x[e][(f"{year_inp}-{month_inp:02d}-{day:02d}", s)]
            for day in week_days for (d, s) in shifts if d == f"{year_inp}-{month_inp:02d}-{day:02d}"
        ) <= max_perweek

#---------------------- Solve ------------------------#
# Objective: maximize preferences
model += pulp.lpSum(preferences[e][(d, s)] * x[e][(d, s)] for e in employees for (d, s) in shifts)
model.solve(pulp.PULP_CBC_CMD(msg=0))


matrix = {emp: {d: "" for d in range(1, month_days+1)} for emp in employees}

for e in employees:
    for (d_str, s) in shifts:
        day = int(d_str.split("-")[2])  # extract day number from "YYYY-MM-DD"
        if pulp.value(x[e][(d_str, s)]) == 1:  # chosen by solver
            matrix[e][day] = s

# ----------------------Build excel- outpout ----------------------#

weekday_labels = ["m", "ti", "o", "to", "f", "l", "s"]

wb = Workbook()
ws = wb.active
ws.title = "Roster"

# Merge first row for month title
ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=month_days+1)
month_name = calendar.month_name[month_inp].upper()
ws.cell(row=1, column=2, value=month_name)
ws.cell(row=1, column=2).alignment = Alignment(horizontal="center", vertical="center")
ws.cell(row=1, column=2).font = Font(bold=True, size=12)

# Row 2: weekdays
for day in range(1, month_days+1):
    weekday = calendar.weekday(year_inp, month_inp, day)  # 0=Mon … 6=Sun
    ws.cell(row=2, column=day+1, value=weekday_labels[weekday])

# Row 3: day numbers
for day in range(1, month_days+1):
    ws.cell(row=3, column=day+1, value=day)

# Column A: employee names
for i, emp in enumerate(employees, start=4):
    ws.cell(row=i, column=1, value=emp)

# Fill roster assignments
for i, emp in enumerate(employees, start=4):
    for day in range(1, month_days+1):
        ws.cell(row=i, column=day+1, value=matrix[emp][day])

# Add totals column
col_total = month_days + 2
ws.column_dimensions["AF"].width=16
ws.cell(row=3, column=col_total, value="#Hours total:")
for i, emp in enumerate(employees, start=4):
    total = sum(shift_hours_wdonly[v] for v in matrix[emp].values() if v != "")
    ws.cell(row=i, column=col_total, value=total)

solution = pulp.LpStatus[model.status]

ws.cell(row=5+len(employees), column=1, value="Hours needed to fill:")
ws.cell(row=6+len(employees), column=1, value="Hours signed:")
ws.cell(row=5+len(employees), column=2, value=str(total_hours_needed))
ws.cell(row=6+len(employees), column=2, value=str(total_hours_signed))
ws.cell(row=7+len(employees), column=1, value="Solution found:")
ws.cell(row=7+len(employees), column=2, value=str(solution))

# -------------Formatting ------------------#
# Center align everything
for row in ws.iter_rows(min_row=1, max_row=3+len(employees),
                        min_col=2, max_col=col_total):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(size=10)

# Narrow the date columns and widen the name column
ws.column_dimensions["A"].width=20
for day in range(1, month_days + 1):
    col_letter = get_column_letter(day + 1)  # +1 because column 1 = names
    ws.column_dimensions[col_letter].width = 4  # adjust width as needed

# Shade weekends
for day in range(1, month_days+1):
    weekday = calendar.weekday(year_inp, month_inp, day)
    if weekday >= 5:  # Saturday or Sunday
        col = day + 1
        for row in range(2, 4+len(employees)):  # shade header + employee rows
            ws.cell(row=row, column=col).fill = PatternFill(start_color="DDDDDD",
                                                            end_color="DDDDDD",
                                                            fill_type="solid")


wb.save("roster.xlsx")
print(pulp.LpStatus[model.status])
end_time = time.time()
elapsed_time = end_time - start_time
print(f"Time to solve: {elapsed_time:.4f} seconds")