import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

wb = Workbook()

# Create Settings sheet
ws_settings = wb.active
ws_settings.title = "Settings"
ws_settings["A1"] = "Pull-Up Progression Settings"
ws_settings["A2"] = "Enter your Max Pull-Up PR in cell C2"
ws_settings["B2"] = "Max PR"
ws_settings["C2"] = 10  # default value; update with your current max

# Create Plan sheet with header
ws_plan = wb.create_sheet(title="Plan")
headers = ["Week", "Day", "Session Type", "Sets", "Factor", "Reps per Set", "Volume"]
ws_plan.append(headers)

# Set header style (bold)
header_font = Font(bold=True)
for cell in ws_plan[1]:
    cell.font = header_font

# Define progression factors for each week (Blocks: Weeks 1-4, 5-8, 9-12)
factors = {
    1: 0.55, 2: 0.65, 3: 0.75, 4: 0.70,
    5: 0.60, 6: 0.70, 7: 0.80, 8: 0.75,
    9: 0.65, 10: 0.75, 11: 0.85, 12: 1.00
}

# Define default session types for weeks 1-11:
# Format: (Session Name, Sets, Adjustment to target reps)
# Adjusted to work between 4-7 reps with higher set counts.
default_sessions = [
    ("Standard", 4, 0),
    ("Volume High", 5, -1),
    ("Standard", 4, 0),
    ("Volume Low", 6, -2)
]

current_row = 2  # start after header
block_weekly_rows = []  # to store weekly summary row numbers for current block
block_number = 1

for week in range(1, 13):
    week_start_row = current_row

    # For week 12, override the session structure:
    if week == 12:
        sessions_for_week = [
            ("Standard", 4, 0),
            ("Light", 4, -2),
            ("Very Light", 4, -3),
            ("Test Day", 1, 0)  # Test Day: no formula, just a prompt
        ]
    else:
        sessions_for_week = default_sessions

    # Write the four training day rows for the week
    day_index = 1
    for session in sessions_for_week:
        session_name, sets_val, adjustment = session
        ws_plan.append([week, f"Day {day_index}", session_name, sets_val, factors[week], None, None])
        cell_rep = ws_plan.cell(row=current_row, column=6)
        cell_vol = ws_plan.cell(row=current_row, column=7)
        if session_name == "Test Day":
            cell_rep.value = "Test"
            cell_vol.value = ""
        else:
            # Create formula for "Reps per Set" using the factor in column E.
            # The formula calculates:
            formula = f"=INT('Settings'!$C$2 * E{current_row}"
            if adjustment > 0:
                formula += f" + {adjustment}"
            elif adjustment < 0:
                formula += f" - {abs(adjustment)}"
            formula += ")"
            cell_rep.value = formula
            # Volume = Sets * Reps per Set
            cell_vol.value = f"=D{current_row}*F{current_row}"
        current_row += 1
        day_index += 1

    # Add weekly summary row (total volume for the week)
    week_end_row = current_row - 1
    ws_plan.append([f"Week {week} Total Volume", "", "", "", "", "", f"=SUM(G{week_start_row}:G{week_end_row})"])
    weekly_summary_row = current_row
    block_weekly_rows.append(weekly_summary_row)
    current_row += 1  # move to next row

    # If end of a block (every 4 weeks) or end of plan, add block summary and separator
    if week % 4 == 0 or week == 12:
        # Create block summary row summing the weekly totals in this block.
        sum_refs = ",".join([f"G{r}" for r in block_weekly_rows])
        ws_plan.append([f"Block {block_number} Total Volume", "", "", "", "", "", f"=SUM({sum_refs})"])
        block_summary_row = current_row
        # Apply a light fill for visibility on the block summary row
        block_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        for col in range(1, 8):
            ws_plan.cell(row=block_summary_row, column=col).fill = block_fill
        current_row += 1  # next row
        # Add an extra blank row for separation
        ws_plan.append([""] * 7)
        current_row += 1
        block_weekly_rows = []  # reset for the next block
        block_number += 1

# Hide the Factor column (Column E)
ws_plan.column_dimensions['E'].hidden = True

# Adjust column widths for readability
for col in ws_plan.columns:
    max_length = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
        if cell.value:
            length = len(str(cell.value))
            if length > max_length:
                max_length = length
    ws_plan.column_dimensions[col_letter].width = max_length + 1

# Save the workbook
wb.save("PullUpProgression.xlsx")
