import csv
from datetime import date
from pathlib import Path


COLUMNS_TO_KEEP = [
    "Initiative/Campaign",
    "PLD Budget",
    "Platform",
    "Publisher",
    "Ad Type_Tactic",
    "Retail Week",
]


def week_of_year_since_june():
    today = date.today()

    if today.month > 6 or (today.month == 6 and today.day >= 1):
        june_year = today.year
    else:
        june_year = today.year - 1

    june_first = date(june_year, 6, 1)
    return (today - june_first).days // 7 + 1


def retail_week_matches_current_week(value: str): 
    value = str(value).strip()
    if not value:
        return False

    try:
        return int(float(value)) == week_of_year_since_june()
    except ValueError:
        return False

def remove_tactic_defects(row: dict[str, str]):
    row['Ad Type_Tactic'] = row['Ad Type_Tactic'].replace("'", '')
    row['Ad Type_Tactic'] = row['Ad Type_Tactic'].replace("[", '')
    row['Ad Type_Tactic'] = row['Ad Type_Tactic'].replace("]", '')


def filter_csv_rows(csv_path: Path):
    with csv_path.open(newline="", encoding="utf-8-sig") as fp:
        reader = csv.DictReader(fp)

        if reader.fieldnames is None:
            raise ValueError("CSV file is missing a header row.")

        missing_columns = [column for column in COLUMNS_TO_KEEP if column not in reader.fieldnames]
        if missing_columns:
            missing = ", ".join(missing_columns)
            raise ValueError(f"CSV file is missing expected columns: {missing}")

        for row in reader:
            remove_tactic_defects(row)
            if retail_week_matches_current_week(row["Retail Week"]):
                yield {column: row[column] for column in COLUMNS_TO_KEEP}


def write_filtered_csv_rows(csv_path: Path):
    rows = list(filter_csv_rows(csv_path))

    with csv_path.open("w", newline="", encoding="utf-8") as fp:
        writer = csv.DictWriter(fp, fieldnames=COLUMNS_TO_KEEP)
        writer.writeheader()
        writer.writerows(rows)


if __name__ == '__main__':
    csv_path = Path(__file__).resolve().parent.parent / "PLD.csv"

    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    write_filtered_csv_rows(csv_path)
