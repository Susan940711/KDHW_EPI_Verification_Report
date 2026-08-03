from __future__ import annotations

from collections import Counter
from datetime import date, datetime
from io import BytesIO
from typing import Any

import openpyxl
import pandas as pd


CODE_HEADER = "children_code"
DATE_COLUMNS = [
    "registered_date",
    "BCG",
    "OPV_first_time_dose",
    "OPV_second_time_dose",
    "OPV_third_time_dose",
    "Penta_first_time_dose",
    "Penta_second_time_dose",
    "Penta_third_time_dose",
    "MMR_first_time_dose",
    "MMR_second_time_dose",
    "IPV",
    "Rota_first_time_dose",
    "Rota_second_time_dose",
]
U1_REQUIRED_DOSES = ["BCG", "OPV1", "OPV2", "OPV3", "Penta1", "Penta2", "Penta3", "MMR1"]
ONE_TO_FIVE_REQUIRED_DOSES = ["OPV1", "OPV2", "OPV3", "Penta1", "Penta2", "Penta3", "MMR1"]


def normalize_code(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_date(value: Any):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def normalize_age_months(value: Any):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return int(value)
    text = normalize_code(value)
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def quarter_label(value: date) -> str:
    quarter = ((value.month - 1) // 3) + 1
    return f"Q{quarter} {value.year}"


def has_received(source_value: str | None) -> bool:
    return normalize_code(source_value) not in {"", "Not received yet"}


def calculate_completed_reporting_month(
    age_months: int | None,
    source_values: dict[str, str],
    reporting_values: dict[str, date | None],
) -> date | None:
    if age_months is None:
        return None

    if 0 <= age_months <= 11:
        required_doses = U1_REQUIRED_DOSES
    elif 12 <= age_months <= 59:
        required_doses = ONE_TO_FIVE_REQUIRED_DOSES
    else:
        return None

    if any(source_values.get(dose, "Not received yet") == "Not received yet" for dose in required_doses):
        return None

    required_entries: list[tuple[date, str]] = []
    for dose_key in required_doses:
        dose_source = source_values.get(dose_key, "Not received yet")
        dose_reporting_month = reporting_values.get(dose_key)
        if dose_reporting_month is not None and has_received(dose_source):
            required_entries.append((dose_reporting_month, dose_source))

    if not required_entries:
        return None

    latest_required_date, _ = max(required_entries, key=lambda item: item[0])
    return latest_required_date


def calculate_completed_dose(
    age_months: int | None,
    source_values: dict[str, str],
    reporting_values: dict[str, date | None],
) -> str:
    if age_months is None:
        return "not complete yet"

    if 0 <= age_months <= 11:
        group_label = "U1"
    elif 12 <= age_months <= 59:
        group_label = "1-5"
    else:
        return "not complete yet"

    completed_reporting_month = calculate_completed_reporting_month(age_months, source_values, reporting_values)
    if completed_reporting_month is None:
        return "not complete yet"
    return f"{group_label} completed in {quarter_label(completed_reporting_month)}"


def resolve_age_months_from_source_row(row: dict[str, Any]) -> int | None:
    dob = normalize_date(row.get("date_of_birth"))
    registered = normalize_date(row.get("registered_date"))
    if dob is None or registered is None:
        return None
    if registered < dob:
        return None

    years = (registered.year - dob.year) * 12 + (registered.month - dob.month)
    if registered.day < dob.day:
        years -= 1
    return years


def load_excel_rows(uploaded_file) -> list[dict[str, Any]]:
    workbook = openpyxl.load_workbook(BytesIO(uploaded_file.getvalue()), data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []

    headers = [normalize_code(cell) for cell in rows[0]]
    data_rows = []
    for row_values in rows[1:]:
        row = {header: row_values[idx] if idx < len(row_values) else None for idx, header in enumerate(headers)}
        data_rows.append(row)
    return data_rows


def build_verification_report(rows: list[dict[str, Any]]) -> pd.DataFrame:
    report_rows = []
    for row in rows:
        age_months = resolve_age_months_from_source_row(row)
        source_values = {
            "BCG": row.get("BCG_source", "Not received yet"),
            "OPV1": row.get("OPV1_source", "Not received yet"),
            "OPV2": row.get("OPV2_source", "Not received yet"),
            "OPV3": row.get("OPV3_source", "Not received yet"),
            "Penta1": row.get("Penta1_source", "Not received yet"),
            "Penta2": row.get("Penta2_source", "Not received yet"),
            "Penta3": row.get("Penta3_source", "Not received yet"),
            "Penta4": row.get("Penta4_source", "Not received yet"),
            "MMR1": row.get("MMR1_source", "Not received yet"),
            "MMR2": row.get("MMR2_source", "Not received yet"),
            "JE1": row.get("JE1_source", "Not received yet"),
            "JE2": row.get("JE2_source", "Not received yet"),
            "IPV": row.get("IPV_source", "Not received yet"),
            "Rota1": row.get("Rota1_source", "Not received yet"),
            "Rota2": row.get("Rota2_source", "Not received yet"),
        }
        reporting_values = {
            "BCG": normalize_date(row.get("BCG_reporting_month")),
            "OPV1": normalize_date(row.get("OPV_first_time_dose_reporting_month")),
            "OPV2": normalize_date(row.get("OPV_second_time_dose_reporting_month")),
            "OPV3": normalize_date(row.get("OPV_third_time_dose_reporting_month")),
            "Penta1": normalize_date(row.get("Penta_first_time_dose_reporting_month")),
            "Penta2": normalize_date(row.get("Penta_second_time_dose_reporting_month")),
            "Penta3": normalize_date(row.get("Penta_third_time_dose_reporting_month")),
            "Penta4": normalize_date(row.get("Penta_fourth_time_dose_reporting_month")),
            "MMR1": normalize_date(row.get("MMR_first_time_dose_reporting_month")),
            "MMR2": normalize_date(row.get("MMR_second_time_dose_reporting_month")),
            "JE1": normalize_date(row.get("JE_first_time_dose_reporting_month")),
            "JE2": normalize_date(row.get("JE_second_time_dose_reporting_month")),
            "IPV": normalize_date(row.get("IPV_reporting_month")),
            "Rota1": normalize_date(row.get("Rota_first_time_dose_reporting_month")),
            "Rota2": normalize_date(row.get("Rota_second_time_dose_reporting_month")),
        }
        completed_dose = calculate_completed_dose(age_months, source_values, reporting_values)
        report_rows.append(
            {
                "children_code": row.get(CODE_HEADER, ""),
                "age_months": age_months,
                "completed_dose": completed_dose,
                "verification_status": "OK" if completed_dose != "not complete yet" else "not complete yet",
            }
        )

    return pd.DataFrame(report_rows)


def filter_report(df: pd.DataFrame, district: str | None = None, township: str | None = None, status: str | None = None) -> pd.DataFrame:
    filtered = df.copy()
    if district:
        filtered = filtered[filtered["district"].astype(str).str.contains(district, case=False, na=False)]
    if township:
        filtered = filtered[filtered["township_name"].astype(str).str.contains(township, case=False, na=False)]
    if status:
        filtered = filtered[filtered["verification_status"].astype(str).str.contains(status, case=False, na=False)]
    return filtered


def run_full_verification(uploaded_file) -> tuple[pd.DataFrame, pd.DataFrame]:
    rows = load_excel_rows(uploaded_file)
    report_df = build_verification_report(rows)
    return pd.DataFrame(rows), report_df
