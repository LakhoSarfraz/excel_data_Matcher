
import pandas as pd
from rapidfuzz import fuzz
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import ColorScaleRule
import re

# ==================================================
# 1. STARTUP
# ==================================================

print("=" * 55)
print("          MULTI-COLUMN EXCEL MATCH FINDER")
print("=" * 55)

def get_file(prompt):

    while True:
        path = input(prompt).strip().strip('"')

        if Path(path).is_file():
            return path

        print("File not found. Try again.")


master_path = get_file("\nEnter MASTER Excel file path: ")
compare_path = get_file("Enter COMPARISON Excel file path: ")


# ==================================================
# 2. SELECT WORKSHEET
# ==================================================

def select_sheet(path, label):

    excel = pd.ExcelFile(path)
    names = excel.sheet_names

    print("\n" + label + " FILE WORKSHEETS")

    for i, name in enumerate(names, 1):
        print(f"{i}. {name}")

    while True:
        try:
            choice = int(input("Select worksheet number: "))

            if 1 <= choice <= len(names):
                return names[choice - 1]

        except ValueError:
            pass

        print("Invalid selection.")


master_sheet = select_sheet(master_path, "MASTER")
compare_sheet = select_sheet(compare_path, "COMPARISON")


# ==================================================
# 3. CHECK ORIGINAL HEADERS
# ==================================================

def check_headers(path, sheet_name):

    wb = load_workbook(
        path,
        read_only=True,
        data_only=True
    )

    ws = wb[sheet_name]

    headers = list(
        next(ws.iter_rows(
            min_row=1,
            max_row=1,
            values_only=True
        ))
    )

    wb.close()

    while headers and headers[-1] is None:
        headers.pop()

    if not headers:
        raise ValueError("Worksheet has no headers.")

    if any(h is None or str(h).strip() == "" for h in headers):
        raise ValueError(
            f"Blank column header found in {sheet_name}. "
            "Please correct the Excel file."
        )

    if len(headers) != len(set(headers)):
        raise ValueError(
            f"Duplicate column names found in {sheet_name}. "
            "Please make all headers unique."
        )

    return headers


master_headers = check_headers(
    master_path, master_sheet
)

compare_headers = check_headers(
    compare_path, compare_sheet
)


# ==================================================
# 4. READ EXCEL DATA
# ==================================================

master_df = pd.read_excel(
    master_path,
    sheet_name=master_sheet,
    header=0
)

compare_df = pd.read_excel(
    compare_path,
    sheet_name=compare_sheet,
    header=0
)

if master_df.empty:
    raise ValueError("Master worksheet contains no records.")

if compare_df.empty:
    raise ValueError("Comparison worksheet contains no records.")


# ==================================================
# 5. EXACT COLUMN MATCHING
# ==================================================

master_columns = list(master_df.columns)
compare_columns = list(compare_df.columns)

matching_columns = [
    col for col in master_columns
    if col in compare_columns
]

master_only = [
    col for col in master_columns
    if col not in compare_columns
]

compare_only = [
    col for col in compare_columns
    if col not in master_columns
]


print("\n" + "=" * 55)
print("              COLUMN MATCHING REPORT")
print("=" * 55)

print("\nMaster columns:", len(master_columns))
print("Comparison columns:", len(compare_columns))
print("Exact matching columns:", len(matching_columns))

print("\nMATCHED COLUMNS:")

for col in matching_columns:
    print("  +", col)

print("\nMASTER-ONLY COLUMNS:")

for col in master_only:
    print("  -", col)

print("\nCOMPARISON-ONLY COLUMNS:")

for col in compare_only:
    print("  -", col)

if not matching_columns:
    raise ValueError(
        "No exact matching column names found. Cannot compare."
    )

if master_only or compare_only:

    print("\nWARNING:")
    print("Only exact matching column names will be compared.")

else:

    print("\nAll column names match.")

input(
    "\nPress ENTER to continue or CTRL+C to cancel..."
)


# ==================================================
# 6. COLUMN WEIGHTING
# ==================================================

print("\n" + "=" * 55)
print("                 COLUMN WEIGHTING")
print("=" * 55)

print("Enter a weight for each column.")
print("Example: Name = 50, Address = 30")
print("Press ENTER for ALL columns to use equal average.")

raw_weights = []

while True:

    raw_weights = []

    for col in matching_columns:

        while True:

            value = input(
                f"Weight for '{col}' (or ENTER): "
            ).strip()

            if value == "":
                raw_weights.append(None)
                break

            try:
                weight = float(value)

                if weight < 0:
                    print("Weight cannot be negative.")
                    continue

                raw_weights.append(weight)
                break

            except ValueError:
                print("Enter a valid number.")

    if all(w is None for w in raw_weights):
        # Equal average
        weights = {
            col: 1 / len(matching_columns)
            for col in matching_columns
        }

        print("\nUsing equal average weighting.")
        break

    if any(w is None for w in raw_weights):
        print(
            "\nEnter weights for ALL columns, "
            "or press ENTER for ALL columns."
        )
        continue

    total_weight = sum(raw_weights)

    if total_weight <= 0:
        print("Total weight must be greater than zero.")
        continue

    if round(total_weight, 6) != 100:
        print(
            f"\nWeight total is {total_weight}%. "
            "Weights will be normalized to 100%."
        )

    weights = {
        col: weight / total_weight
        for col, weight in zip(
            matching_columns,
            raw_weights
        )
    }

    break


print("\nFINAL COLUMN WEIGHTS:")

for col, weight in weights.items():
    print(f"{col}: {weight * 100:.2f}%")


# ==================================================
# 7. MINIMUM MATCHING THRESHOLD
# ==================================================

while True:

    try:
        minimum_score = float(
            input("\nMinimum matching percentage (0-100): ")
        )

        if 0 <= minimum_score <= 100:
            break

        print("Enter a value between 0 and 100.")

    except ValueError:
        print("Enter a valid number.")


# ==================================================
# 8. TEXT CLEANING
# ==================================================

def clean_text(value):

    if pd.isna(value):
        return ""

    text = str(value).casefold().strip()

    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text)

    return text


# ==================================================
# 9. PREPARE RECORDS
# ==================================================

master_records = []
compare_records = []

for i in range(len(master_df)):

    master_records.append({
        "row": i + 2,
        "data": master_df.iloc[i].to_dict(),
        "clean": {
            col: clean_text(master_df.iloc[i][col])
            for col in matching_columns
        }
    })


for i in range(len(compare_df)):

    compare_records.append({
        "row": i + 2,
        "data": compare_df.iloc[i].to_dict(),
        "clean": {
            col: clean_text(compare_df.iloc[i][col])
            for col in matching_columns
        }
    })


n = len(master_records)
m = len(compare_records)

print("\nMaster records:", n)
print("Comparison records:", m)
print("Possible comparisons:", n * m)


# ==================================================
# 10. CALCULATE MULTI-COLUMN SIMILARITY
# ==================================================

def calculate_similarity(master, compare):

    weighted_total = 0
    available_weight = 0

    for col in matching_columns:

        a = master["clean"][col]
        b = compare["clean"][col]

        # Exclude missing values from both sides
        if not a or not b:
            continue

        column_score = max(
            fuzz.WRatio(a, b),
            fuzz.token_set_ratio(a, b)
        )

        weight = weights[col]

        weighted_total += column_score * weight
        available_weight += weight

    if available_weight == 0:
        return 0

    # Automatically normalize remaining weights
    final_score = weighted_total / available_weight

    return round(final_score, 2)


# ==================================================
# 11. COMPARE ALL RECORDS
# ==================================================

print("\nComparing records...")

candidates = []

compare_candidates = {
    j: [] for j in range(m)
}

for i, master in enumerate(master_records):

    for j, compare in enumerate(compare_records):

        score = calculate_similarity(
            master,
            compare
        )

        if score >= minimum_score and score > 0:

            candidates.append((score, i, j))

            compare_candidates[j].append(
                (score, i)
            )

    if (i + 1) % 100 == 0 or i + 1 == n:
        print(f"Processed: {i + 1}/{n}")


# ==================================================
# 12. GREEDY HIGHEST-FIRST MATCHING
# ==================================================

print("\nSelecting highest matches...")

candidates.sort(
    key=lambda x: (-x[0], x[1], x[2])
)

assignments = {}

used_master = set()
used_compare = set()

for score, i, j in candidates:

    if i in used_master:
        continue

    if j in used_compare:
        continue

    assignments[i] = (j, score)

    used_master.add(i)
    used_compare.add(j)


# ==================================================
# 13. PREPARE UNIQUE OUTPUT HEADERS
# ==================================================

def unique_headers(headers):

    result = []
    seen = set()

    for name in headers:

        new_name = str(name)
        original = new_name
        counter = 2

        while new_name in seen:

            new_name = f"{original} ({counter})"
            counter += 1

        seen.add(new_name)
        result.append(new_name)

    return result


# Matched Results headers
matched_headers = list(master_columns)

comparison_output_headers = []

for col in compare_columns:

    if col in master_columns:
        comparison_output_headers.append(
            "Comparison - " + str(col)
        )
    else:
        comparison_output_headers.append(str(col))


matched_headers += comparison_output_headers

matched_metadata = [
    "Best Match - Similarity %",
    "Sheet 2 Row Number"
]

matched_headers = unique_headers(
    matched_headers + matched_metadata
)


# Remaining Records headers
remaining_headers = unique_headers(
    list(compare_columns) + [
        "Sheet 2 Row Number",
        "Lost Match %",
        "Master Sheet Row Number"
    ]
)


# ==================================================
# 14. CREATE MATCHED RESULTS
# ==================================================

print("\nPreparing matched results...")

matched_rows = []

for i, master in enumerate(master_records):

    master_values = [
        master["data"][col]
        for col in master_columns
    ]

    if i in assignments:

        j, score = assignments[i]

        compare = compare_records[j]

        compare_values = [
            compare["data"][col]
            for col in compare_columns
        ]

        extra_values = [
            score,
            compare["row"]
        ]

    else:

        compare_values = [
            None for _ in compare_columns
        ]

        extra_values = [None, None]

    matched_rows.append(
        master_values + compare_values + extra_values
    )


matched_df = pd.DataFrame(
    matched_rows,
    columns=matched_headers
)


# ==================================================
# 15. CREATE REMAINING RECORDS
# ==================================================

print("\nPreparing remaining records...")

remaining_rows = []

for j, compare in enumerate(compare_records):

    if j in used_compare:
        continue

    compare_values = [
        compare["data"][col]
        for col in compare_columns
    ]

    # Find Master records that already received
    # another Comparison record.
    losing_candidates = [
        (score, i)
        for score, i in compare_candidates[j]
        if i in used_master
    ]

    if losing_candidates:

        losing_candidates.sort(
            key=lambda x: (-x[0], x[1])
        )

        lost_score, master_index = losing_candidates[0]

        lost_percentage = lost_score

        master_row_number = (
            master_records[master_index]["row"]
        )

    else:

        lost_percentage = None
        master_row_number = None

    extra_values = [
        compare["row"],
        lost_percentage,
        master_row_number
    ]

    remaining_rows.append(
        compare_values + extra_values
    )


remaining_df = pd.DataFrame(
    remaining_rows,
    columns=remaining_headers
)


# ==================================================
# 16. COLUMN MATCHING REPORT
# ==================================================

report_rows = []

for col in matching_columns:

    report_rows.append({
        "Column Name": col,
        "Master": "Yes",
        "Comparison": "Yes",
        "Weight (%)": round(weights[col] * 100, 2),
        "Status": "Matched"
    })


for col in master_only:

    report_rows.append({
        "Column Name": col,
        "Master": "Yes",
        "Comparison": "No",
        "Weight (%)": None,
        "Status": "Master only"
    })


for col in compare_only:

    report_rows.append({
        "Column Name": col,
        "Master": "No",
        "Comparison": "Yes",
        "Weight (%)": None,
        "Status": "Comparison only"
    })


report_df = pd.DataFrame(report_rows)


# ==================================================
# 17. EXPORT EXCEL
# ==================================================

output_file = Path(master_path).with_name(
    Path(master_path).stem
    + "_VS_"
    + Path(compare_path).stem
    + "_Matching_Report.xlsx"
)


with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:

    matched_df.to_excel(
        writer,
        sheet_name="Matched Results",
        index=False
    )

    remaining_df.to_excel(
        writer,
        sheet_name="Remaining Records",
        index=False
    )

    report_df.to_excel(
        writer,
        sheet_name="Column Matching Report",
        index=False
    )


# ==================================================
# 18. EXCEL FORMATTING
# ==================================================

wb = load_workbook(output_file)

header_fill = PatternFill(
    "solid",
    fgColor="305496"
)

header_font = Font(
    bold=True,
    color="FFFFFF"
)

for ws in wb.worksheets:

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:

        cell.fill = header_fill
        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    for column in ws.columns:

        max_length = 0

        for cell in column:

            value = str(cell.value or "")
            max_length = max(max_length, len(value))

        ws.column_dimensions[
            column[0].column_letter
        ].width = min(max_length + 3, 55)


# Similarity percentage formatting
ws = wb["Matched Results"]

similarity_col = matched_headers.index(
    "Best Match - Similarity %"
) + 1

ws.column_dimensions[
    ws.cell(1, similarity_col).column_letter
].width = 23

if ws.max_row >= 2:

    letter = ws.cell(
        1, similarity_col
    ).column_letter

    ws.conditional_formatting.add(
        f"{letter}2:{letter}{ws.max_row}",
        ColorScaleRule(
            start_type="num",
            start_value=0,
            start_color="F8696B",
            mid_type="num",
            mid_value=75,
            mid_color="FFEB84",
            end_type="num",
            end_value=100,
            end_color="63BE7B"
        )
    )


# Lost match percentage formatting
ws = wb["Remaining Records"]

lost_col = remaining_headers.index("Lost Match %") + 1

ws.column_dimensions[
    ws.cell(1, lost_col).column_letter
].width = 18

if ws.max_row >= 2:

    letter = ws.cell(1, lost_col).column_letter

    ws.conditional_formatting.add(
        f"{letter}2:{letter}{ws.max_row}",
        ColorScaleRule(
            start_type="num",
            start_value=0,
            start_color="F8696B",
            mid_type="num",
            mid_value=75,
            mid_color="FFEB84",
            end_type="num",
            end_value=100,
            end_color="63BE7B"
        )
    )


wb.save(output_file)


# ==================================================
# 19. FINAL SUMMARY
# ==================================================

print("\n" + "=" * 55)
print("                  COMPLETED")
print("=" * 55)

print("\nMASTER FILE:", Path(master_path).name)
print("COMPARISON FILE:", Path(compare_path).name)

print("\nMatching columns:", len(matching_columns))
print("Minimum threshold:", minimum_score, "%")

print("\nMaster records:", n)
print("Comparison records:", m)

print("Matched records:", len(assignments))
print("Unmatched Master:", n - len(assignments))
print("Remaining Comparison:", len(remaining_df))

lost_count = sum(
    1 for row in remaining_rows
    if row[-2] is not None
)

print("Remaining records with lost matches:", lost_count)

print("\nOUTPUT FILE:")
print(output_file)

print("\nThree output sheets created:")
print("1. Matched Results")
print("2. Remaining Records")
print("3. Column Matching Report")

print("\nOriginal Excel files unchanged.")
print("=" * 55)
