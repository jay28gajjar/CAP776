from openpyxl import load_workbook
from datetime import date, datetime, timedelta


# ============================================================
# 1. LOAD EXCEL WORKBOOK
# ============================================================

file = r"C:\Users\Jayyg\OneDrive\Documents\CAP776\12612265.xlsx"

workbook = load_workbook(file, data_only=True)
sheet = workbook["Daily Log"]

print("Excel file loaded successfully.")
print("Worksheet:", sheet.title)


# ============================================================
# 2. DISPLAY EXCEL HEADINGS
# ============================================================

print("\nExcel headings:")

for cell in sheet[5]:
    print(cell.value)

# ============================================================
# 3. CREATE DATA STRUCTURES
# ============================================================

records = []
errors = []

print("\nData structures created.")


# ============================================================
# 4. PROJECT DATE RANGE
# ============================================================

start_date = date(2026, 8, 17)
end_date = date(2026, 9, 21)

print("Project period:", start_date, "to", end_date)


# ============================================================
# 5. RATING CONVERSION FUNCTION
# ============================================================

def convert_rating_to_number(rating):

    if rating is None:
        return None

    rating = str(rating).strip().lower()

    # Day's Feeling
    if rating == "excellent":
        return 5
    elif rating == "good":
        return 4
    elif rating == "neutral":
        return 3
    elif rating == "low":
        return 2
    elif rating == "stressed":
        return 1

    # Satisfaction
    elif rating == "very satisfied":
        return 5
    elif rating == "satisfied":
        return 4
    elif rating == "unsatisfied":
        return 2
    elif rating == "very unsatisfied":
        return 1

    # Energy
    elif rating == "high":
        return 5
    elif rating == "medium":
        return 3
    elif rating == "low":
        return 1

    return None


# ============================================================
# 6. VALIDATE MINUTES
# ============================================================

def validate_minutes(value, name, row_num):

    try:
        value = float(value)

    except (ValueError, TypeError):
        return None, (
            "Row "
            + str(row_num)
            + ": Invalid "
            + name
        )

    if value < 0 or value > 1440:
        return None, (
            "Row "
            + str(row_number)
            + ": "
            + name
            + " must be between 0 and 1440 minutes"
        )

    return value, None


# ============================================================
# 7. READ AND VALIDATE EXCEL DATA
# ============================================================

for row_number, row in enumerate(
    sheet.iter_rows(min_row=6, values_only=True),
    start=6
):

    try:

        current_date = row[0]
        sleep_minutes = row[1]
        fitness_minutes = row[2]
        study_minutes = row[3]
        coding_minutes = row[4]
        class_minutes = row[5]
        classes_attended = row[6]
        other_minutes = row[7]
        total_tracked = row[8]
        free_minutes = row[9]
        feeling = row[10]
        satisfaction = row[11]
        energy = row[12]
        notes = row[13]

        # Skip empty rows
        if current_date is None:
            continue

        # ----------------------------------------------------
        # Validate time-related values
        # ----------------------------------------------------

        sleep_minutes, error = validate_minutes(
            sleep_minutes,
            "Sleep",
            row_number
        )

        if error:
            errors.append(error)
            continue

        fitness_minutes, error = validate_minutes(
            fitness_minutes,
            "Fitness",
            row_number
        )

        if error:
            errors.append(error)
            continue

        study_minutes, error = validate_minutes(
            study_minutes,
            "Study",
            row_number
        )

        if error:
            errors.append(error)
            continue

        coding_minutes, error = validate_minutes(
            coding_minutes,
            "Coding",
            row_number
        )

        if error:
            errors.append(error)
            continue

        class_minutes, error = validate_minutes(
            class_minutes,
            "Class",
            row_number
        )

        if error:
            errors.append(error)
            continue

        other_minutes, error = validate_minutes(
            other_minutes,
            "Other Activities",
            row_number
        )

        if error:
            errors.append(error)
            continue

        # ----------------------------------------------------
        # Validate classes attended
        # ----------------------------------------------------

        try:
            classes_attended = float(classes_attended)

            if classes_attended < 0:
                raise ValueError

        except (ValueError, TypeError):
            errors.append(
                "Row "
                + str(row_number)
                + ": Invalid Classes Attended"
            )
            continue

        # ----------------------------------------------------
        # Calculate total tracked time
        # ----------------------------------------------------

        calculated_total = (
            sleep_minutes
            + fitness_minutes
            + study_minutes
            + coding_minutes
            + class_minutes
            + other_minutes
        )

        if calculated_total > 1440:
            errors.append(
                "Row "
                + str(row_number)
                + ": Total activity time exceeds 1440 minutes"
            )
            continue

        # ----------------------------------------------------
        # Calculate free time
        # ----------------------------------------------------

        calculated_free = 1440 - calculated_total

        # ----------------------------------------------------
        # Validate ratings
        # ----------------------------------------------------

        if convert_rating_to_number(feeling) is None:
            errors.append(
                "Row "
                + str(row_number)
                + ": Invalid Day's Feeling"
            )
            continue

        if convert_rating_to_number(satisfaction) is None:
            errors.append(
                "Row "
                + str(row_number)
                + ": Invalid Satisfaction Level"
            )
            continue

        if convert_rating_to_number(energy) is None:
            errors.append(
                "Row "
                + str(row_number)
                + ": Invalid Energy Level"
            )
            continue

        # ----------------------------------------------------
        # Store valid record
        # ----------------------------------------------------

        day_data = {
            "date": current_date,
            "sleep_minutes": sleep_minutes,
            "fitness_minutes": fitness_minutes,
            "study_minutes": study_minutes,
            "coding_minutes": coding_minutes,
            "class_minutes": class_minutes,
            "classes_attended": classes_attended,
            "other_minutes": other_minutes,
            "total_tracked": calculated_total,
            "free_minutes": calculated_free,
            "feeling": feeling,
            "satisfaction": satisfaction,
            "energy": energy,
            "notes": notes
        }

        records.append(day_data)

    except Exception as error:
        errors.append(
            "Row "
            + str(row_number)
            + ": "
            + str(error)
        )


# ============================================================
# 8. DISPLAY VALIDATION RESULTS
# ============================================================

print("\nValid records:", len(records))
print("Errors:", len(errors))

if len(errors) == 0:
    print("No errors found.")

else:
    print("\nErrors found:")

    for error in errors:
        print(error)


# ============================================================
# 9. DISPLAY FIRST RECORD
# ============================================================

if len(records) > 0:

    print("\nFirst valid record:")
    print(records[0])


# ============================================================
# 10. AVERAGE CALCULATION FUNCTION
# ============================================================

def find_average(recordss, key_name):

    total = 0
    count = 0

    for record in recordss:

        value = record[key_name]

        if value is not None:
            total += value
            count += 1

    if count == 0:
        return 0

    return total / count


# ============================================================
# 11. CALCULATE AVERAGES
# ============================================================

average_sleep = find_average(records, "sleep_minutes")
average_fitness = find_average(records, "fitness_minutes")
average_study = find_average(records, "study_minutes")
average_coding = find_average(records, "coding_minutes")
average_class = find_average(records, "class_minutes")
average_other = find_average(records, "other_minutes")
average_total = find_average(records, "total_tracked")
average_free = find_average(records, "free_minutes")


print("\nAverage Sleep:", round(average_sleep, 2))
print("Average Fitness:", round(average_fitness, 2))
print("Average Study:", round(average_study, 2))
print("Average Coding:", round(average_coding, 2))
print("Average Class:", round(average_class, 2))
print("Average Other:", round(average_other, 2))
print("Average Total Tracked:", round(average_total, 2))
print("Average Free Time:", round(average_free, 2))


# ============================================================
# 12. CONVERT MINUTES TO HOURS
# ============================================================

def minutes_to_hours(minutes):
    return minutes / 60


print(
    "Average Sleep:",
    round(minutes_to_hours(average_sleep), 2),
    "hours"
)

print(
    "Average Study:",
    round(minutes_to_hours(average_study), 2),
    "hours"
)

print(
    "Average Coding:",
    round(minutes_to_hours(average_coding), 2),
    "hours"
)


# ============================================================
# 13. TPI - TECHNICAL PRODUCTIVITY INDEX
# ============================================================

def calculate_tpi(records):

    return find_average(
        records,
        "coding_minutes"
    )


# ============================================================
# 14. AAI - ACADEMIC ACTIVITY INDEX
# ============================================================

def calculate_aai(records):

    if len(records) == 0:
        return 0

    total_academic = 0

    for record in records:

        academic_time = (
            record["study_minutes"]
            + record["class_minutes"]
        )

        total_academic += academic_time

    return total_academic / len(records)


# ============================================================
# 15. PhAI - PHYSICAL ACTIVITY INDEX
# ============================================================

def calculate_phai(records):

    return find_average(
        records,
        "fitness_minutes"
    )


# ============================================================
# 16. SRI - SLEEP REGULARITY INDEX
# ============================================================

def calculate_sri(records):

    return find_average(
        records,
        "sleep_minutes"
    )


# ============================================================
# 17. ABI - AVAILABLE/BALANCE INDEX
# ============================================================

def calculate_abi(records):

    return find_average(
        records,
        "free_minutes"
    )


# ============================================================
# 18. TUI - TOTAL UTILIZATION INDEX
# ============================================================

def calculate_tui(records):

    return find_average(
        records,
        "total_tracked"
    )


# ============================================================
# 19. RATING CONVERSION FUNCTIONS
# ============================================================

def feeling_to_number(value):

    mapping = {
        "Excellent": 5,
        "Good": 4,
        "Neutral": 3,
        "Low": 2,
        "Stressed": 1
    }

    if value is None:
        return None

    return mapping.get(str(value).strip().title())


def satisfaction_to_number(value):

    mapping = {
        "Very Satisfied": 5,
        "Satisfied": 4,
        "Neutral": 3,
        "Unsatisfied": 2,
        "Very Unsatisfied": 1
    }

    if value is None:
        return None

    return mapping.get(str(value).strip().title())


def energy_to_number(value):

    mapping = {
        "High": 3,
        "Medium": 2,
        "Low": 1
    }

    if value is None:
        return None

    return mapping.get(str(value).strip().title())


# ============================================================
# 20. EI - ENERGY INDEX
# ============================================================

def calculate_ei(records):

    if len(records) == 0:
        return 0

    total_score = 0
    count = 0

    for record in records:

        feeling_score = feeling_to_number(
            record["feeling"]
        )

        satisfaction_score = satisfaction_to_number(
            record["satisfaction"]
        )

        energy_score = energy_to_number(
            record["energy"]
        )

        if (
            feeling_score is not None
            and satisfaction_score is not None
            and energy_score is not None
        ):

            daily_score = (
                feeling_score
                + satisfaction_score
                + energy_score
            ) / 3

            total_score += daily_score
            count += 1

    if count == 0:
        return 0

    return total_score / count


# ============================================================
# 21. DCI - DATA COMPLETENESS INDEX
# ============================================================

def calculate_dci(records, expected_days):

    if expected_days == 0:
        return 0

    return len(records) / expected_days


# ============================================================
# 22. CALCULATE ALL INDICES
# ============================================================

tpi = calculate_tpi(records)
aai = calculate_aai(records)
phai = calculate_phai(records)
sri = calculate_sri(records)
abi = calculate_abi(records)
tui = calculate_tui(records)
ei = calculate_ei(records)

expected_days = 36

dci = calculate_dci(
    records,
    expected_days
)


print("\nTPI :", round(tpi, 2), "minutes/day")
print("AAI :", round(aai, 2), "minutes/day")
print("PhAI:", round(phai, 2), "minutes/day")
print("SRI :", round(sri, 2), "minutes/day")
print("ABI :", round(abi, 2), "minutes/day")
print("TUI :", round(tui, 2), "minutes/day")
print("EI  :", round(ei, 2), "/ 5")
print("DCI :", round(dci * 100, 2), "%")


# ============================================================
# 23. PAI - PERSONAL ACTIVITY INDEX
# ============================================================

def calculate_pai(
    tpi,
    aai,
    phai,
    sri,
    tui,
    ei,
    dci
):

    pai = (
        (0.15 * tpi)
        + (0.20 * aai)
        + (0.15 * phai)
        + (0.20 * sri)
        + (0.15 * tui)
        + (0.10 * ei)
        + (0.05 * dci)
    )

    return pai


pai = calculate_pai(
    tpi,
    aai,
    phai,
    sri,
    tui,
    ei,
    dci
)

print("PAI:", round(pai, 2))


# ============================================================
# 24. CORRELATION FUNCTION
# ============================================================

def find_correlation(x_values, y_values):

    if len(x_values) != len(y_values):
        return None

    if len(x_values) < 2:
        return None

    x_average = sum(x_values) / len(x_values)
    y_average = sum(y_values) / len(y_values)

    numerator = 0
    x_square_sum = 0
    y_square_sum = 0

    for i in range(len(x_values)):

        x_difference = (
            x_values[i]
            - x_average
        )

        y_difference = (
            y_values[i]
            - y_average
        )

        numerator += (
            x_difference
            * y_difference
        )

        x_square_sum += (
            x_difference ** 2
        )

        y_square_sum += (
            y_difference ** 2
        )

    denominator = (
        x_square_sum
        * y_square_sum
    ) ** 0.5

    if denominator == 0:
        return None

    return numerator / denominator


# ============================================================
# 25. CREATE MATCHING LISTS
# ============================================================

def make_matching_lists(
    records,
    first_column,
    second_column
):

    first_values = []
    second_values = []

    for record in records:

        first_value = record[first_column]
        second_value = record[second_column]

        if second_column == "energy":

            second_score = energy_to_number(
                second_value
            )

        elif second_column == "satisfaction":

            second_score = satisfaction_to_number(
                second_value
            )

        elif second_column == "feeling":

            second_score = feeling_to_number(
                second_value
            )

        else:
            second_score = None

        if (
            first_value is not None
            and second_score is not None
        ):

            first_values.append(
                float(first_value)
            )

            second_values.append(
                float(second_score)
            )

    return first_values, second_values


# ============================================================
# 26. CODING VS ENERGY CORRELATION
# ============================================================

coding_values, energy_values = make_matching_lists(
    records,
    "coding_minutes",
    "energy"
)

coding_energy = find_correlation(
    coding_values,
    energy_values
)

print(
    "\nCoding vs Energy:",
    round(coding_energy, 3)
)


# ============================================================
# 27. SLEEP VS ENERGY CORRELATION
# ============================================================

sleep_values, energy_values = make_matching_lists(
    records,
    "sleep_minutes",
    "energy"
)

sleep_energy = find_correlation(
    sleep_values,
    energy_values
)

print(
    "Sleep vs Energy:",
    round(sleep_energy, 3)
)


# ============================================================
# 28. STUDY VS SATISFACTION CORRELATION
# ============================================================

study_values, satisfaction_values = make_matching_lists(
    records,
    "study_minutes",
    "satisfaction"
)

study_satisfaction = find_correlation(
    study_values,
    satisfaction_values
)

print(
    "Study vs Satisfaction:",
    round(study_satisfaction, 3)
)


# ============================================================
# 29. INTERPRET CORRELATION
# ============================================================

def interpret_correlation(value):

    if value is None:
        return "Not enough valid data"

    if value >= 0.70:
        return "Strong positive relationship"

    elif value >= 0.30:
        return "Moderate positive relationship"

    elif value > -0.30:
        return "Weak or little linear relationship"

    elif value > -0.70:
        return "Moderate negative relationship"

    else:
        return "Strong negative relationship"


print(
    "Coding vs Energy:",
    round(coding_energy, 3),
    "->",
    interpret_correlation(coding_energy)
)

print(
    "Sleep vs Energy:",
    round(sleep_energy, 3),
    "->",
    interpret_correlation(sleep_energy)
)

print(
    "Study vs Satisfaction:",
    round(study_satisfaction, 3),
    "->",
    interpret_correlation(study_satisfaction)
)


# ============================================================
# 30. CHECK DATE CONTINUITY
# ============================================================

def check_date_continuity(
    records,
    start_date,
    end_date
):

    if hasattr(start_date, "date"):
        start_date = start_date.date()

    if hasattr(end_date, "date"):
        end_date = end_date.date()

    dates = []

    for record in records:

        record_date = record["date"]

        if hasattr(record_date, "date"):
            record_date = record_date.date()

        dates.append(record_date)

    missing_dates = []
    current_date = start_date

    while current_date <= end_date:

        if current_date not in dates:
            missing_dates.append(current_date)

        current_date += timedelta(days=1)

    return missing_dates


# ============================================================
# 31. FIND MISSING DATES
# ============================================================

start_date = datetime(2026, 8, 17)
end_date = datetime(2026, 9, 21)

missing_dates = check_date_continuity(
    records,
    start_date,
    end_date
)

print("\nMissing Dates:")

if len(missing_dates) == 0:
    print("No missing dates.")

else:

    for missing_date in missing_dates:
        print(missing_date)


# ============================================================
# 32. FIND DUPLICATE DATES
# ============================================================

def find_duplicate_dates(records):

    dates = []
    duplicates = []

    for record in records:

        current_date = record["date"]

        if hasattr(current_date, "date"):
            current_date = current_date.date()

        if current_date in dates:
            duplicates.append(current_date)

        else:
            dates.append(current_date)

    return duplicates


duplicate_dates = find_duplicate_dates(records)


# ============================================================
# 33. DATA QUALITY SUMMARY
# ============================================================

expected_days = (
    end_date - start_date
).days + 1

print("\nExpected days:", expected_days)
print("Valid records:", len(records))
print("Invalid records:", len(errors))
print("Missing dates:", len(missing_dates))
print("Duplicate dates:", len(duplicate_dates))


if len(missing_dates) > 0:

    print("\nMissing dates:")

    for missing_date in missing_dates:
        print(missing_date)


if len(duplicate_dates) > 0:

    print("\nDuplicate dates:")

    for duplicate_date in duplicate_dates:
        print(duplicate_date)


# ============================================================
# 34. FINAL SUMMARY
# ============================================================

print(
    "\nMy average coding time was",
    round(average_coding, 2),
    "minutes per day."
)

print(
    "My average study time was",
    round(average_study, 2),
    "minutes per day."
)

print(
    "My average sleep time was",
    round(average_sleep, 2),
    "minutes per day."
)

print(
    "My average fitness time was",
    round(average_fitness, 2),
    "minutes per day."
)

print(
    "My average free/unaccounted time was",
    round(average_free, 2),
    "minutes per day."
)

print(
    "Coding and Energy correlation:",
    round(coding_energy, 3)
)

print(
    "Sleep and Energy correlation:",
    round(sleep_energy, 3)
)

print(
    "Study and Satisfaction correlation:",
    round(study_satisfaction, 3)
)