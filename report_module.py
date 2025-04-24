import os
from datetime import datetime
from collections import Counter
import difflib
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt


def get_output(sheet, column, exclude_strings):
    return sum(1 for row in range(1, sheet.max_row + 1)
               if (val := sheet.cell(row=row, column=column).value) and val not in exclude_strings)

def get_accepted_rings(sheet, column_letter, string_to_find):
    return sum(1 for row in range(1, sheet.max_row + 1) if sheet[f"{column_letter}{row}"].value == string_to_find)

def get_rejected_rings(sheet, column, exclude_strings):
    return sum(1 for row in range(1, sheet.max_row + 1)
               if (val := sheet.cell(row=row, column=column).value) and val not in exclude_strings)

def get_rework_rings(sheet, column_letter, string_to_find):
    return sum(1 for row in range(1, sheet.max_row + 1) if sheet[f"{column_letter}{row}"].value == string_to_find)

def calculate_yield(okay_rings, total_rings):
    return (okay_rings / total_rings) * 100 if total_rings else 0

def get_rejection_details(sheet, column_index):
    values = [cell.value for row in sheet.iter_rows(min_col=column_index, max_col=column_index, min_row=1, max_row=sheet.max_row)
              for cell in row if cell.value and str(cell.value).strip()]
    return dict(Counter(values))

def get_cover_mismatch(sheet, column_letter, string_to_find):
    return sum(1 for row in range(1, sheet.max_row + 1) if sheet[f"{column_letter}{row}"].value == string_to_find)

def fuzzy_match(reason, keywords, cutoff=0.8):
    match = difflib.get_close_matches(reason, keywords, n=1, cutoff=cutoff)
    return match[0] if match else None

def generate_bar_chart(data_dict, output_folder):
    plt.figure(figsize=(12, 6))
    labels, values = zip(*data_dict.items()) if data_dict else ([], [])

    bars = plt.bar(labels, values, color="#3498db")

    # Add values at the top of bars
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2.0, yval + 0.5, f"{int(yval)}", ha='center', va='bottom', fontsize=9)

    plt.xticks(rotation=45, ha='right')
    plt.ylabel("Count")
    plt.title("Individual Rejection Reasons")
    plt.tight_layout()

    chart_path = os.path.join(output_folder, "rejection_chart.png")
    plt.savefig(chart_path)
    plt.close()
    return chart_path

def generate_report(file_path, report_for):
    today = datetime.today().date()
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    total_rings = get_output(sheet, 3, [])
    accepted_rings = get_accepted_rings(sheet, 'C', 'Accepted')
    rejected_rings = get_rejected_rings(sheet, 3, ['Accepted', 'REWORK', 'Cover Mismatch'])
    reworked_rings = get_rework_rings(sheet, 'C', 'REWORK')
    yield_percentage = calculate_yield(accepted_rings, total_rings)
    rejection_details = get_rejection_details(sheet, 4)
    cover_mismatch = get_cover_mismatch(sheet, 'D', 'Cover Mismatch')
    casting_keywords = ['DUST ON RESIN', 'MICRO BUBBLE', 'SPM REJECTION']
    polishing_keywords = ['SHELL COATING REMOVED', 'SIDE SCRATCH', 'SIDE SCRATCH(EMERY)',
                          'IMPROPER RESIN FINISH', 'RESIN DAMAGE', 'DISCONNECTING ISSUE',
                          'CHARGING CODE ISSUE', 'CE TAPE ISSUE']
    
    casting_keywords = [kw.lower() for kw in casting_keywords]
    polishing_keywords = [kw.lower() for kw in polishing_keywords]

    casting_rejections, polishing_rejections, other_rejections = {}, {}, {}

    for reason, count in rejection_details.items():
        reason_clean = str(reason).strip()
        reason_lower = reason_clean.lower()

        if fuzzy_match(reason_lower, casting_keywords):
            casting_rejections[reason_clean] = count
        elif fuzzy_match(reason_lower, polishing_keywords):
            polishing_rejections[reason_clean] = count
        else:
            other_rejections[reason_clean] = count    

    # Clean unwanted entries
    for key in ['Accepted', 'REWORK', 'Cover Mismatch']:
        rejection_details.pop(key, None)

    output_folder = os.path.dirname(file_path)
    output_path = os.path.join(output_folder, "output.txt")

    with open(output_path, "w") as f:
        print(f"REPORT FOR {report_for}: {today}", file=f)
        print(f"OUTPUT: {total_rings}", file=f)
        print(f"OKAY: {accepted_rings}", file=f)
        print(f"REJECTED: {rejected_rings}", file=f)
        if reworked_rings > 0:
            print(f"REWORK: {reworked_rings}", file=f)
        print(f"YIELD: {yield_percentage:.2f}%", file=f)

        print("\nREJECTION DETAILS:", file=f)
        print("\nCasting Rejections:", file=f)
        if casting_rejections:
            print(pd.DataFrame.from_dict(casting_rejections, orient='index').to_string(header=False), file=f)
        else:
            print("None", file=f)

        print("\nPolishing Issues:", file=f)
        if polishing_rejections:
            print(pd.DataFrame.from_dict(polishing_rejections, orient='index').to_string(header=False), file=f)
        else:
            print("None", file=f)

        print("\nOther Rejections:", file=f)
        if other_rejections:
            print(pd.DataFrame.from_dict(other_rejections, orient='index').to_string(header=False), file=f)
        else:
            print("None", file=f)

        print(f"\nCOVER MISMATCH: {cover_mismatch}", file=f)

    chart_path = generate_bar_chart(rejection_details, output_folder)
    os.startfile(output_path)
    os.startfile(chart_path)
    return output_path
