from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import hashlib
import json
import re
from io import BytesIO
from math import isfinite
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


REPORTS = {
    "All units": {"keys": ["A", "I"], "validate": ["E", "G"]},
    "Delinquent and Prepaid": {"keys": ["C"], "validate": ["D", "G", "I", "N"]},
    "Rent Roll Report": {"keys": ["C", "H"], "validate": ["N", "S", "T"]},
    "Resident Deposit": {"keys": ["A", "B"], "validate": ["C", "F", "G", "H", "I", "J"]},
    "Market Rent Evaluation": {"keys": [], "validate": []},
    "Availability Report": {"keys": [], "validate": []},
    "Rentable and Assignable Items Report": {"keys": [], "validate": []},
    "Resident by Fical Period": {"keys": [], "validate": []},
}

CODE_MAPS = {
    "Delinquent and Prepaid": {
        "HAPPD": "HAP", "FUTURE: RENT": "PMTCHECK", "FUTURE: SUBRENT": "PMTHAP",
        "LOSTRENT": "MISCCREDIT", "HOUSING": "RENT",
    },
    "Rent Roll Report": {"HVOUCH": "SUBSIDY"},
}

NAVY = "12334A"
TEAL = "147D83"
INK = "1E293B"
MUTED = "64748B"
PALE = "F3F7FA"
LINE = "DCE5EC"
RED = "FDE8E7"
GREEN = "E4F4EE"
WHITE = "FFFFFF"


def col_index(letter: str) -> int:
    """Return a zero based spreadsheet column index."""
    value = 0
    for char in letter.upper():
        value = value * 26 + ord(char) - 64
    return value - 1


def read_pdf_table(file_bytes: bytes, filename: str, report_type: str) -> pd.DataFrame:
    if report_type != "Delinquent and Prepaid":
        raise ValueError("PDF uploads are currently supported for Delinquent and Prepaid reports only. Use an Excel file for this report type.")

    try:
        import pdfplumber
    except ImportError as exc:
        raise ValueError("PDF support requires pdfplumber. Install the updated project requirements and restart the app.") from exc

    target_columns = {
        "Bldg/Unit": 2,
        "Name": 3,
        "Status": 6,
        "Code Description": 8,
        "Total Prepaid": 9,
        "Total Delinquent": 10,
        "Net Balance": 13,
    }
    def normalize_pdf_header(value: Any) -> str:
        return re.sub(r"[^a-z0-9]+", "", str(value or "").casefold())

    rows: list[list[Any]] = []
    found_table = False
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables():
                    if not table:
                        continue
                    for header_row_index, header_row in enumerate(table[:5]):
                        normalized_headers = [normalize_pdf_header(value) for value in header_row]
                        source_columns = {
                            label: next((i for i, value in enumerate(normalized_headers) if value == normalize_pdf_header(label)), None)
                            for label in target_columns
                        }
                        if any(index is None for index in source_columns.values()):
                            continue
                        found_table = True
                        if not rows:
                            pdf_headers: list[Any] = [None] * 14
                            for label, target_index in target_columns.items():
                                pdf_headers[target_index] = label
                            rows.append(pdf_headers)
                        for source_row in table[header_row_index + 1:]:
                            if not source_row or all(value is None or not str(value).strip() for value in source_row):
                                continue
                            code_index = source_columns["Code Description"]
                            code_value = str(source_row[code_index] or "").strip() if code_index < len(source_row) else ""
                            row_text = " ".join(str(value or "").strip() for value in source_row if value is not None).casefold()
                            if code_value.upper().startswith("TOTAL") or row_text.startswith("totals"):
                                continue
                            normalized_row: list[Any] = [None] * 14
                            for label, target_index in target_columns.items():
                                source_index = source_columns[label]
                                normalized_row[target_index] = source_row[source_index] if source_index < len(source_row) else None
                            rows.append(normalized_row)
                        break
    except Exception as exc:
        raise ValueError(f"Could not read the report table from {filename}. Make sure the PDF is not password protected and contains selectable text.") from exc

    if not found_table or len(rows) <= 1:
        raise ValueError(f"No Delinquent and Prepaid table was found in {filename}. Scanned/image-only PDFs are not supported yet; use a selectable-text PDF or an Excel file.")
    return pd.DataFrame(rows, dtype=object)


def read_first_sheet(file_bytes: bytes, filename: str, report_type: str | None = None) -> pd.DataFrame:
    extension = Path(filename).suffix.casefold()
    if extension == ".pdf":
        return read_pdf_table(file_bytes, filename, report_type or "")
    engine = "xlrd" if extension == ".xls" else "openpyxl"
    try:
        return pd.read_excel(BytesIO(file_bytes), sheet_name=0, header=None, dtype=object, engine=engine)
    except ImportError as exc:
        if engine == "xlrd":
            raise ValueError("Para sa .xls, kulang ang xlrd package. Gawing .xlsx ang file at i-upload ulit.") from exc
        raise
    except Exception as exc:
        raise ValueError(f"Hindi mabasa ang {filename}. Tiyaking valid ang Excel file.") from exc


MARKET_RENT_FIELDS = [
    "Unit Number", "Description/Amenities", "Unit SQFT", "Base Rent", "Amenity Value", "Market Rent",
]


def extract_market_rent_data(raw: pd.DataFrame) -> pd.DataFrame:
    """Extract the unit-level fields from a Market Rent Evaluation worksheet."""
    header_row = None
    for row_index in range(min(len(raw), 100)):
        labels = [re.sub(r"[^a-z0-9]", "", str(value).casefold()) for value in raw.iloc[row_index].tolist()]
        if any(label == "bldgunit" for label in labels):
            header_row = row_index
            break
    if header_row is None:
        raise ValueError("Could not find the Bldg/Unit header in the Market Rent Evaluation file.")

    normalized_targets = {
        "Unit Number": "bldgunit",
        "Description/Amenities": "descriptionamenities",
        "Unit SQFT": "unitsqft",
        "Base Rent": "baserent",
        "Amenity Value": "amenityvalue",
        "Market Rent": "marketrent",
    }
    positions: dict[str, int] = {}
    for col_index, value in enumerate(raw.iloc[header_row].tolist()):
        label = re.sub(r"[^a-z0-9]", "", str(value).casefold())
        for field, target in normalized_targets.items():
            if label == target:
                positions[field] = col_index
    missing = [field for field in MARKET_RENT_FIELDS if field not in positions]
    if missing:
        raise ValueError("Missing required Market Rent Evaluation columns: " + ", ".join(missing))

    records = []
    for _, source_row in raw.iloc[header_row + 1:].iterrows():
        unit_value = source_row.iloc[positions["Unit Number"]]
        if pd.isna(unit_value) or not str(unit_value).strip():
            continue
        unit_label = str(unit_value).strip()
        if unit_label.casefold().startswith(("property total", "amenity")):
            continue
        sqft = source_row.iloc[positions["Unit SQFT"]]
        base_rent = source_row.iloc[positions["Base Rent"]]
        market_rent = source_row.iloc[positions["Market Rent"]]
        if all(pd.isna(value) for value in (sqft, base_rent, market_rent)):
            continue
        record = {field: source_row.iloc[positions[field]] for field in MARKET_RENT_FIELDS}
        for numeric_field in ("Unit SQFT", "Base Rent", "Amenity Value", "Market Rent"):
            record[numeric_field] = pd.to_numeric(
                str(record[numeric_field]).replace(",", "").replace("$", ""), errors="coerce",
            )
        record["Unit Number"] = unit_label
        record["Description/Amenities"] = "" if pd.isna(record["Description/Amenities"]) else str(record["Description/Amenities"]).strip()
        records.append(record)
    if not records:
        raise ValueError("No unit-level Market Rent Evaluation records were found below the column headers.")
    return pd.DataFrame(records, columns=MARKET_RENT_FIELDS).reset_index(drop=True)


def guess_header_row(frame: pd.DataFrame, columns: list[str]) -> int:
    """Choose the row with the most useful labels in configured columns."""
    candidate_indices = sorted({col_index(c) for c in columns})
    if not candidate_indices:
        candidate_indices = list(range(min(frame.shape[1], 30)))
    best_row, best_score = 0, -1
    for row_index in range(min(len(frame), 80)):
        values = [frame.iat[row_index, i] for i in candidate_indices if i < frame.shape[1]]
        nonblank = [v for v in values if not pd.isna(v) and str(v).strip()]
        text_count = sum(isinstance(v, str) for v in nonblank)
        score = len(nonblank) * 3 + text_count
        if score > best_score:
            best_row, best_score = row_index, score
    return best_row


def clean_frame(frame: pd.DataFrame, header_row: int) -> tuple[list[str], pd.DataFrame]:
    labels = []
    used: dict[str, int] = defaultdict(int)
    for i, value in enumerate(frame.iloc[header_row].tolist()):
        label = "" if pd.isna(value) else str(value).strip()
        label = label or f"Column {get_column_letter(i + 1)}"
        used[label] += 1
        labels.append(f"{label} ({used[label]})" if used[label] > 1 else label)
    rows = frame.iloc[header_row + 1:].copy().reset_index(drop=True)
    rows = rows.loc[~rows.apply(lambda row: all(pd.isna(x) or str(x).strip() == "" for x in row), axis=1)]
    return labels, rows.reset_index(drop=True)


def normalize(value: Any, code_map: dict[str, str] | None = None) -> str:
    if value is None or (not isinstance(value, (list, tuple, dict)) and pd.isna(value)):
        return ""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.isoformat()
    if isinstance(value, bool):
        text = str(value).lower()
    elif isinstance(value, (int, float, Decimal)):
        try:
            number = Decimal(str(value))
            text = format(number.normalize(), "f") if number.is_finite() else str(value)
        except (InvalidOperation, ValueError):
            text = str(value).strip()
    else:
        text = str(value).strip()
    if code_map:
        text = code_map.get(text.upper(), text)
    return text.casefold()


def display_value(value: Any) -> Any:
    if value is None or (not isinstance(value, (list, tuple, dict)) and pd.isna(value)):
        return ""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, float) and isfinite(value) and value.is_integer():
        return int(value)
    return value


def filename_matches_report(filename: str, report_type: str) -> bool:
    ignored = {"and", "by", "report", "file", "excel", "xls", "xlsx", "old", "new", "site", "test", "copy", "final", "export", "validation"}

    def tokens(text: str) -> set[str]:
        # Split CamelCase names such as RentRollReport before normalizing punctuation.
        text = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", Path(text).stem)
        values = re.findall(r"[a-z0-9]+", text.casefold())
        aliases = {"fical": "fiscal", "units": "unit", "items": "item", "deposits": "deposit"}
        return {aliases.get(token, token) for token in values if token not in ignored and not token.isdigit()}

    expected = tokens(report_type)
    actual = tokens(filename)
    if not expected:
        return True
    required_matches = (2 * len(expected) + 2) // 3
    return len(expected & actual) >= required_matches


def money_display(value: Any) -> str:
    if value is None or (not isinstance(value, (list, tuple, dict)) and pd.isna(value)):
        return ""
    try:
        amount = Decimal(str(value).replace(",", "").replace("$", "")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return f"{amount:,.2f}"
    except (InvalidOperation, ValueError):
        return str(value)


def paired_display(record: dict[str, Any], field: str, money: bool = False) -> str:
    old_value = record.get(f"Old {field}", "")
    new_value = record.get(f"New {field}", "")
    old_display = display_value(old_value)
    new_display = display_value(new_value)
    old_text = money_display(old_value) if money else str(old_display)
    new_text = money_display(new_value) if money else str(new_display)
    same = normalize_money(old_value) == normalize_money(new_value) if money else normalize(old_value) == normalize(new_value)
    return old_text if same else f"{old_text} → {new_text}"


def result_display_rows(result: dict[str, Any], report_type: str) -> list[dict[str, Any]]:
    records = result["differences"]
    if report_type == "Resident Deposit":
        display_fields = [
            "Status", "Total Dep. Required", "Unit Deposit", "Pet Deposit",
            "Other Deposits", "Deposits On Hand",
        ]
        return [
            {"Name": record.get("Name", ""), **{
                field: paired_display(record, field, money=field != "Status")
                for field in display_fields
            }}
            for record in records
        ]
    if report_type != "Delinquent and Prepaid":
        return records
    output = [
        {
            "Bldg/Unit": record.get("Bldg/Unit", ""),
            "Name": paired_display(record, "Name"),
            "Status": paired_display(record, "Status"),
            "Code Description": paired_display(record, "Code Description"),
            "Net Balance": paired_display(record, "Net Balance", money=True),
        }
        for record in records
    ]
    if result.get("old_mismatch_net_balance_total") is not None:
        total_text = (
            f"Old {money_display(result['old_mismatch_net_balance_total'])} | "
            f"New {money_display(result['new_mismatch_net_balance_total'])} | "
            f"Difference {money_display(result['mismatch_net_balance_total_difference'])}"
        )
        output.append({"Bldg/Unit": "", "Name": "", "Status": "", "Code Description": "TOTALS", "Net Balance": total_text})
    return output


def resident_deposit_display_rows(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    fields = [
        "Name", "Status", "Total Dep. Required", "Unit Deposit", "Pet Deposit",
        "Other Deposits", "Deposits On Hand",
    ]
    return [{field: record.get(field, "") for field in fields} for record in records]


def add_net_balance_table_total(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not records or "Net Balance" not in records[0]:
        return records
    total = sum((money_decimal(record.get("Net Balance")) for record in records), Decimal("0.00"))
    footer = {column: "" for column in records[0]}
    if "Code Description" in footer:
        footer["Code Description"] = "TOTALS"
    else:
        footer["Bldg/Unit"] = "TOTALS"
    footer["Net Balance"] = money_display(total)
    return [*records, footer]


def choose_indices(old_labels: list[str], new_labels: list[str], configured: list[str], prefix: str, kind: str) -> list[tuple[int, int]]:
    """Render paired column selectors for old and new workbooks."""
    if not old_labels or not new_labels:
        return []
    if not configured:
        st.caption("Walang default setup para sa report na ito. Pumili ng parehong dami ng columns sa magkabilang file at panatilihin ang magkaparehong order.")
        left, right = st.columns(2)
        with left:
            old_selected = st.multiselect(
                f"{kind} columns · Old Site Report", range(len(old_labels)),
                format_func=lambda i: f"{get_column_letter(i + 1)} · {old_labels[i]}",
                key=f"{prefix}_old_custom",
            )
        with right:
            new_selected = st.multiselect(
                f"{kind} columns · New Site (Test Site Report)", range(len(new_labels)),
                format_func=lambda i: f"{get_column_letter(i + 1)} · {new_labels[i]}",
                key=f"{prefix}_new_custom",
            )
        if len(old_selected) != len(new_selected):
            st.warning("Pumili ng parehong bilang ng columns sa Old Site Report at New Site (Test Site Report).")
            return []
        return list(zip(old_selected, new_selected))
    old_pairs: list[tuple[int, int]] = []
    for n, letter in enumerate(configured):
        default_index = col_index(letter)
        old_default = min(default_index, len(old_labels) - 1)
        new_default = min(default_index, len(new_labels) - 1)
        left, right = st.columns(2)
        options = list(range(len(old_labels)))
        label = f"{kind} {n + 1}"
        with left:
            old_index = st.selectbox(
                f"{label} · Old Site Report", options,
                format_func=lambda i, names=old_labels: f"{get_column_letter(i + 1)} · {names[i]}",
                index=old_default, key=f"{prefix}_old_{n}",
            )
        with right:
            new_index = st.selectbox(
                f"{label} · New Site (Test Site Report)", range(len(new_labels)),
                format_func=lambda i, names=new_labels: f"{get_column_letter(i + 1)} · {names[i]}",
                index=new_default, key=f"{prefix}_new_{n}",
            )
        old_pairs.append((old_index, new_index))
    return old_pairs


def row_key(row: pd.Series, pairs: list[tuple[int, int]], side: str) -> tuple[str, ...]:
    positions = [left if side == "old" else right for left, right in pairs]
    return tuple(normalize(row.iloc[i] if i < len(row) else None) for i in positions)


def normalize_money(value: Any) -> str:
    if value is None or (not isinstance(value, (list, tuple, dict)) and pd.isna(value)):
        return ""
    text = str(value).strip().replace(",", "").replace("$", "")
    try:
        amount = Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return format(amount, ".2f")
    except (InvalidOperation, ValueError):
        return normalize(value)


def values_match(old_value: Any, new_value: Any, code_map: dict[str, str], map_code: bool = False, round_money: bool = False) -> bool:
    if round_money:
        return normalize_money(old_value) == normalize_money(new_value)
    return normalize(old_value, code_map if map_code else None) == normalize(new_value)


def is_code_field(label: str) -> bool:
    return "code" in label.casefold() or "description" in label.casefold()


def sum_net_balance(frame: pd.DataFrame, key_column: int, net_column: int) -> float:
    keys = frame.iloc[:, key_column]
    has_key = keys.notna() & keys.astype(str).str.strip().ne("")
    amounts = pd.to_numeric(
        frame.iloc[:, net_column].astype(str).str.replace(",", "", regex=False).str.replace("$", "", regex=False),
        errors="coerce",
    )
    return float(amounts[has_key].sum())


def find_header_index(headers: list[str], label: str) -> int | None:
    target = label.casefold()
    return next((i for i, header in enumerate(headers) if header.strip().casefold() == target), None)


def selected_field_pair(
    old_headers: list[str], new_headers: list[str], pairs: list[tuple[int, int]], label: str,
) -> tuple[int, int] | None:
    target = label.casefold()
    for old_i, new_i in pairs:
        old_label = old_headers[old_i].casefold() if old_i < len(old_headers) else ""
        new_label = new_headers[new_i].casefold() if new_i < len(new_headers) else ""
        if target in old_label or target in new_label:
            return old_i, new_i
    return None


def money_decimal(value: Any) -> Decimal:
    if value is None or (not isinstance(value, (list, tuple, dict)) and pd.isna(value)):
        return Decimal("0.00")
    text = str(value).strip().replace(",", "").replace("$", "")
    try:
        return Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return Decimal("0.00")


def prepare_delinquent_prepaid(
    frame: pd.DataFrame,
    headers: list[str],
    key_column: int,
    name_column: int,
    status_column: int,
    code_column: int,
    prepaid_column: int,
    delinquent_column: int,
    code_map: dict[str, str],
    map_old_codes: bool,
) -> pd.DataFrame:
    """Aggregate each unit's charge rows and offset PMT credits against RENT."""
    groups: dict[str, dict[str, Any]] = {}
    for row_number, (_, row) in enumerate(frame.iterrows()):
        unit = display_value(row.iloc[key_column] if key_column < len(row) else None)
        if not normalize(unit):
            # Ignore spreadsheet totals and footer rows, which have no Bldg/Unit ID.
            continue
        unit_key = normalize(unit)
        raw_code = str(display_value(row.iloc[code_column] if code_column < len(row) else "")).strip()
        code = code_map.get(raw_code.upper(), raw_code) if map_old_codes else raw_code
        code = str(code).strip()
        code_key = code.casefold()
        group = groups.setdefault(unit_key, {
            "unit": unit,
            "name": display_value(row.iloc[name_column] if name_column < len(row) else None),
            "status": display_value(row.iloc[status_column] if status_column < len(row) else None),
            "balances": {},
            "labels": {},
            "pmt_prepaid": {},
        })
        group["labels"].setdefault(code_key, code)
        prepaid = money_decimal(row.iloc[prepaid_column] if prepaid_column < len(row) else None)
        delinquent = money_decimal(row.iloc[delinquent_column] if delinquent_column < len(row) else None)
        group["balances"][code_key] = group["balances"].get(code_key, Decimal("0.00")) + delinquent
        if code.upper().startswith("PMT"):
            group["pmt_prepaid"][code_key] = group["pmt_prepaid"].get(code_key, Decimal("0.00")) + prepaid
        else:
            group["balances"][code_key] += prepaid

    records = []
    for group in groups.values():
        balances = group["balances"]
        rent_key = next((key for key, code in group["labels"].items() if code.upper() == "RENT"), None)
        if rent_key is not None:
            balances[rent_key] += sum(group["pmt_prepaid"].values(), Decimal("0.00"))
        else:
            for code_key, amount in group["pmt_prepaid"].items():
                balances[code_key] = balances.get(code_key, Decimal("0.00")) + amount
        for code_key, amount in balances.items():
            code_label = group["labels"].get(code_key, "")
            if rent_key is not None and code_label.upper().startswith("PMT") and amount == 0:
                continue
            records.append({
                "Bldg/Unit": group["unit"],
                "Name": group["name"],
                "Status": group["status"],
                "Code Description": code_label,
                "Net Balance": float(amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)),
            })
    return pd.DataFrame(records, columns=["Bldg/Unit", "Name", "Status", "Code Description", "Net Balance"])


def compare_frames(
    old: pd.DataFrame,
    new: pd.DataFrame,
    old_headers: list[str],
    new_headers: list[str],
    key_pairs: list[tuple[int, int]],
    validation_pairs: list[tuple[int, int]],
    code_map: dict[str, str],
) -> dict[str, Any]:
    new_by_key: dict[tuple[str, ...], list[int]] = defaultdict(list)
    old_key_counts: dict[tuple[str, ...], int] = defaultdict(int)
    for i, row in new.iterrows():
        key = row_key(row, key_pairs, "new")
        if any(key):
            new_by_key[key].append(i)
    for _, row in old.iterrows():
        key = row_key(row, key_pairs, "old")
        if any(key):
            old_key_counts[key] += 1

    code_description_pair = next((
        (old_col, new_col)
        for old_col, new_col in validation_pairs
        if "code description" in (new_headers[new_col] if new_col < len(new_headers) else "").casefold()
    ), None)

    matched: list[tuple[int, int]] = []
    missing_old: list[int] = []
    used_new: set[int] = set()
    for old_i, row in old.iterrows():
        key = row_key(row, key_pairs, "old")
        if not any(key):
            missing_old.append(old_i)
            continue
        candidates = [i for i in new_by_key.get(key, []) if i not in used_new]
        if not candidates:
            missing_old.append(old_i)
            continue
        old_row = old.iloc[old_i]
        exact = next((new_i for new_i in candidates if all(
            values_match(
                old_row.iloc[old_col] if old_col < len(old_row) else None,
                new.iloc[new_i].iloc[new_col] if new_col < len(new.iloc[new_i]) else None,
                code_map,
                is_code_field(new_headers[new_col] if new_col < len(new_headers) else ""),
                "net balance" in (new_headers[new_col] if new_col < len(new_headers) else "").casefold(),
            )
            for old_col, new_col in validation_pairs
        )), None)
        if exact is not None:
            new_i = exact
        elif old_key_counts[key] > 1 or len(new_by_key.get(key, [])) > 1:
            # Repeated keys represent multiple charge lines for the same unit.
            # Do not pair unrelated charges just because they share a unit ID.
            same_code = [
                candidate for candidate in candidates
                if code_description_pair is not None and values_match(
                    old_row.iloc[code_description_pair[0]] if code_description_pair[0] < len(old_row) else None,
                    new.iloc[candidate].iloc[code_description_pair[1]] if code_description_pair[1] < len(new.iloc[candidate]) else None,
                    code_map,
                    True,
                )
            ]
            if not same_code:
                missing_old.append(old_i)
                continue
            new_i = same_code[0]
        else:
            new_i = candidates[0]
        matched.append((old_i, new_i))
        used_new.add(new_i)
    missing_new = [i for i in range(len(new)) if i not in used_new]

    key_labels = []
    for n, (left_i, right_i) in enumerate(key_pairs):
        left_label = old_headers[left_i] if left_i < len(old_headers) else f"Column {get_column_letter(left_i + 1)}"
        right_label = new_headers[right_i] if right_i < len(new_headers) else f"Column {get_column_letter(right_i + 1)}"
        key_labels.append(right_label or left_label or f"Key {n + 1}")

    differences: list[dict[str, Any]] = []
    for old_i, new_i in matched:
        old_row, new_row = old.iloc[old_i], new.iloc[new_i]
        key_values = {key_labels[n]: display_value(old_row.iloc[a] if a < len(old_row) else None) for n, (a, _) in enumerate(key_pairs)}
        mismatch_row = dict(key_values)
        has_mismatch = False
        for n, (old_col, new_col) in enumerate(validation_pairs):
            old_val = old_row.iloc[old_col] if old_col < len(old_row) else None
            new_val = new_row.iloc[new_col] if new_col < len(new_row) else None
            label = new_headers[new_col] if new_col < len(new_headers) else f"Column {get_column_letter(new_col + 1)}"
            label = label or f"Validation {n + 1}"
            mismatch_row[f"Old {label}"] = display_value(old_val)
            mismatch_row[f"New {label}"] = display_value(new_val)
            if not values_match(old_val, new_val, code_map, is_code_field(label), "net balance" in label.casefold()):
                has_mismatch = True
        if has_mismatch:
            differences.append(mismatch_row)

    net_balance_pair = next((
        (old_col, new_col)
        for old_col, new_col in validation_pairs
        if "net balance" in (new_headers[new_col] if new_col < len(new_headers) else "").casefold()
    ), None)
    if net_balance_pair:
        old_net_total = sum_net_balance(old, key_pairs[0][0], net_balance_pair[0])
        new_net_total = sum_net_balance(new, key_pairs[0][1], net_balance_pair[1])
        old_balance_label = old_headers[net_balance_pair[0]] if net_balance_pair[0] < len(old_headers) else "Net Balance"
        new_balance_label = new_headers[net_balance_pair[1]] if net_balance_pair[1] < len(new_headers) else "Net Balance"
        old_mismatch_total = sum(money_decimal(record.get(f"Old {old_balance_label}")) for record in differences)
        new_mismatch_total = sum(money_decimal(record.get(f"New {new_balance_label}")) for record in differences)
    else:
        old_net_total = new_net_total = None
        old_mismatch_total = new_mismatch_total = None
    return {
        "matched": matched,
        "missing_old": missing_old,
        "missing_new": missing_new,
        "differences": differences,
        "key_labels": key_labels,
        "old_headers": old_headers,
        "new_headers": new_headers,
        "old_rows": old,
        "new_rows": new,
        "old_net_balance_total": old_net_total,
        "new_net_balance_total": new_net_total,
        "net_balance_total_difference": new_net_total - old_net_total if old_net_total is not None and new_net_total is not None else None,
        "old_mismatch_net_balance_total": old_mismatch_total,
        "new_mismatch_net_balance_total": new_mismatch_total,
        "mismatch_net_balance_total_difference": new_mismatch_total - old_mismatch_total if old_mismatch_total is not None and new_mismatch_total is not None else None,
    }


def row_records(frame: pd.DataFrame, indexes: list[int], headers: list[str], output_columns: list[int]) -> list[dict[str, Any]]:
    records = []
    for idx in indexes:
        row = frame.iloc[idx]
        record = {headers[i]: display_value(row.iloc[i]) if i < len(row) else "" for i in output_columns if i < len(headers)}
        records.append(record)
    return records


def market_rent_output_rows(result: dict[str, Any]) -> list[dict[str, Any]]:
    old_frame: pd.DataFrame = result["market_rent_old"]
    new_frame: pd.DataFrame = result["market_rent_new"]
    old_by_index = {old_i: new_i for old_i, new_i in result["matched"]}
    new_by_index = {new_i: old_i for old_i, new_i in result["matched"]}
    pairs: list[tuple[int | None, int | None]] = [(old_i, new_i) for old_i, new_i in result["matched"]]
    pairs.extend((old_i, None) for old_i in result["missing_old"] if old_i not in old_by_index)
    pairs.extend((None, new_i) for new_i in result["missing_new"] if new_i not in new_by_index)
    output = []
    for old_i, new_i in pairs:
        old_rec = old_frame.iloc[old_i].to_dict() if old_i is not None else {}
        new_rec = new_frame.iloc[new_i].to_dict() if new_i is not None else {}
        status = "Missing from New Site" if new_i is None else "Missing from Old Site" if old_i is None else "Matched"
        if old_i is not None and new_i is not None:
            status = "Mismatched" if any(
                not values_match(old_rec.get(field), new_rec.get(field), {}, False, field in {"Base Rent", "Amenity Value", "Market Rent"})
                for field in MARKET_RENT_FIELDS[1:]
            ) else "Matched"
        row = {"Unit Number": old_rec.get("Unit Number", new_rec.get("Unit Number", "")), "Comparison Status": status}
        for field in MARKET_RENT_FIELDS[1:]:
            row[f"Old {field}"] = old_rec.get(field, "")
            row[f"New {field}"] = new_rec.get(field, "")
        output.append(row)
    return output


def market_amenity_summary(result: dict[str, Any]) -> list[list[Any]]:
    grouped: dict[str, dict[str, Any]] = {}
    grand_old_rent = grand_new_rent = Decimal("0")
    grand_old_count = grand_new_count = 0
    for site, frame in (("Old", result["market_rent_old"]), ("New", result["market_rent_new"])):
        for _, record in frame.iterrows():
            description = str(record.get("Description/Amenities", "") or "").strip()
            if not description or "no floor plan" in description.casefold() or "no unit amenit" in description.casefold():
                continue
            item = grouped.setdefault(description, {"Old Rent": Decimal("0"), "Old Units": 0, "New Rent": Decimal("0"), "New Units": 0})
            rent = money_decimal(record.get("Amenity Value"))
            item[f"{site} Rent"] += rent
            item[f"{site} Units"] += 1
            if site == "Old":
                grand_old_rent += rent
                grand_old_count += 1
            else:
                grand_new_rent += rent
                grand_new_count += 1
    rows = [[name, float(values["Old Rent"]), values["Old Units"], float(values["New Rent"]), values["New Units"]]
            for name, values in sorted(grouped.items(), key=lambda pair: pair[0].casefold())]
    rows.append(["TOTAL", float(grand_old_rent), grand_old_count, float(grand_new_rent), grand_new_count])
    return rows


def market_field_total(frame: pd.DataFrame, field: str) -> float:
    amounts = pd.to_numeric(frame[field], errors="coerce").fillna(0)
    return float(amounts.sum())


def resident_deposit_total(frame: pd.DataFrame, headers: list[str], field: str) -> float:
    name_index = find_header_index(headers, "Name")
    amount_index = find_header_index(headers, field)
    if name_index is None or amount_index is None:
        return 0.0
    names = frame.iloc[:, name_index]
    has_resident = names.notna() & names.astype(str).str.strip().ne("")
    amounts = pd.to_numeric(
        frame.iloc[:, amount_index].astype(str).str.replace(",", "", regex=False).str.replace("$", "", regex=False),
        errors="coerce",
    ).fillna(0)
    return float(amounts[has_resident].sum())


def add_market_rent_sheets(wb: Workbook, result: dict[str, Any]) -> None:
    unit_sheet = wb.create_sheet("Market Rent Units")
    unit_rows = market_rent_output_rows(result)
    unit_headers = ["Unit Number", "Comparison Status"] + [f"{site} {field}" for field in MARKET_RENT_FIELDS[1:] for site in ("Old", "New")]
    unit_sheet.append(unit_headers)
    for record in unit_rows:
        unit_sheet.append([record.get(header, "") for header in unit_headers])
    unit_sheet.freeze_panes = "A2"
    style_sheet(unit_sheet)
    for row in unit_sheet.iter_rows(min_row=2):
        if row[1].value == "Mismatched":
            for cell in row:
                cell.font = Font(name="Aptos", size=10, bold=True, color="B42318")
        elif row[1].value != "Matched":
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=RED)
    for col_index, header in enumerate(unit_headers, start=1):
        if "Rent" in header or "Value" in header:
            for row_index in range(2, unit_sheet.max_row + 1):
                unit_sheet.cell(row_index, col_index).number_format = "#,##0.00"
        elif "SQFT" in header:
            for row_index in range(2, unit_sheet.max_row + 1):
                unit_sheet.cell(row_index, col_index).number_format = "#,##0"

    amenity_sheet = wb.create_sheet("Amenities Summary")
    amenity_sheet.append(["Amenity / Description", "Old Site Total Amenity Rent", "Old Site Unit Assignments", "New Site Total Amenity Rent", "New Site Unit Assignments"])
    for row in market_amenity_summary(result):
        amenity_sheet.append(row)
    amenity_sheet.freeze_panes = "A2"
    style_sheet(amenity_sheet)
    for row_index in range(2, amenity_sheet.max_row + 1):
        for col_index in (2, 4):
            amenity_sheet.cell(row_index, col_index).number_format = "#,##0.00"
        if row_index == amenity_sheet.max_row:
            for cell in amenity_sheet[row_index]:
                cell.fill = PatternFill("solid", fgColor=PALE)
                cell.font = Font(name="Aptos", size=10, bold=True, color=NAVY)


def make_report(result: dict[str, Any], report_type: str, task: str) -> bytes:
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    summary.append(["OS2OS VALIDATION · COMPARISON SUMMARY"])
    summary.merge_cells("A1:D1")
    summary.append(["Task", task, "Report type", report_type])
    summary.append(["Matched records", len(result["matched"]), "Compared fields", len(result["validation_pairs"])])
    summary.append(["Mismatched records", len(result["differences"]), "Missing from New Site (Test Site Report)", len(result["missing_old_records"])])
    summary.append(["Missing from Old Site Report", len(result["missing_new_records"]), "Result", "Differences found" if result["differences"] or result["missing_old_records"] or result["missing_new_records"] else "All matched values agree"])
    if result.get("old_mismatch_net_balance_total") is not None:
        summary.append(["Old Site Mismatch Net Balance Total", float(result["old_mismatch_net_balance_total"]), "New Site Mismatch Net Balance Total", float(result["new_mismatch_net_balance_total"])])
        summary.append(["Mismatch Net Balance Difference", float(result["mismatch_net_balance_total_difference"])])
    if report_type == "Market Rent Evaluation":
        old_market = result["market_rent_old"]
        new_market = result["market_rent_new"]
        summary.append(["Old Site Total SQFT", market_field_total(old_market, "Unit SQFT"), "New Site Total SQFT", market_field_total(new_market, "Unit SQFT")])
        summary.append(["Old Site Total Base Rent", market_field_total(old_market, "Base Rent"), "New Site Total Base Rent", market_field_total(new_market, "Base Rent")])
        for column_index in (2, 4):
            summary.cell(row=summary.max_row - 1, column=column_index).number_format = "#,##0"
            summary.cell(row=summary.max_row, column=column_index).number_format = "#,##0.00"
    summary.append([])
    summary.append(["Matching is based on the selected key columns. Only the first worksheet in each upload is read."])
    note_row = summary.max_row
    summary.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=4)

    def add_table(name: str, records: list[dict[str, Any]], empty_message: str) -> None:
        sheet = wb.create_sheet(name)
        if records:
            columns = list(records[0].keys())
            sheet.append(columns)
            for record in records:
                sheet.append([record.get(column, "") for column in columns])
            tint = RED
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    cell.fill = PatternFill("solid", fgColor=tint)
        else:
            sheet.append([empty_message])
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        style_sheet(sheet)
        if name == "Differences":
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    if any(cell.value == "TOTALS" for cell in row):
                        cell.fill = PatternFill("solid", fgColor=PALE)
                        cell.font = Font(name="Aptos", size=10, bold=True, color=NAVY)
                    else:
                        cell.font = Font(name="Aptos", size=10, bold=True, color="B42318")
        elif name in {"Missing from New", "Missing from Old Site Report"}:
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    if any(item.value == "TOTALS" for item in row):
                        cell.fill = PatternFill("solid", fgColor=PALE)
                        cell.font = Font(name="Aptos", size=10, bold=True, color=NAVY)
                    else:
                        cell.font = Font(name="Aptos", size=10, bold=True, color="B42318")
        for column_index, header in enumerate(columns if records else [], start=1):
            if "net balance" in header.casefold():
                for row_index in range(2, sheet.max_row + 1):
                    sheet.cell(row=row_index, column=column_index).number_format = "#,##0.00"

    if result.get("old_net_balance_total") is not None:
        total_row = 6
        summary.cell(row=total_row, column=2).number_format = "#,##0.00"
        summary.cell(row=total_row, column=4).number_format = "#,##0.00"
        summary.cell(row=total_row + 1, column=2).number_format = "#,##0.00"

    add_table("Differences", result_display_rows(result, report_type), "No value differences found.")
    missing_old_records = result["missing_old_records"]
    missing_new_records = result["missing_new_records"]
    if report_type == "Resident Deposit":
        missing_old_records = resident_deposit_display_rows(missing_old_records)
        missing_new_records = resident_deposit_display_rows(missing_new_records)
    add_table("Missing from New", add_net_balance_table_total(missing_old_records), "No rows are missing from the New Site (Test Site Report).")
    add_table("Missing from Old Site Report", add_net_balance_table_total(missing_new_records), "No rows are missing from the Old Site Report.")
    if report_type == "Market Rent Evaluation":
        add_market_rent_sheets(wb, result)
    style_sheet(summary, title=True)
    summary.column_dimensions["A"].width = 29
    summary.column_dimensions["B"].width = 36
    summary.column_dimensions["C"].width = 26
    summary.column_dimensions["D"].width = 36
    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def style_sheet(sheet, title: bool = False) -> None:
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.freeze_panes = "A2" if not title else None
    for row in sheet.iter_rows():
        for cell in row:
            cell.font = Font(name="Aptos", size=10, color=INK)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="hair", color=LINE))
    if title:
        cell = sheet["A1"]
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name="Aptos Display", size=16, bold=True, color=WHITE)
        cell.alignment = Alignment(vertical="center")
        sheet.row_dimensions[1].height = 32
        for cell in sheet[2]:
            cell.fill = PatternFill("solid", fgColor=PALE)
            cell.font = Font(name="Aptos", size=10, bold=True, color=TEAL)
    elif sheet.max_row >= 1:
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
    for column in sheet.columns:
        letter = get_column_letter(column[0].column)
        width = min(42, max(14, max((len(str(c.value or "")) for c in column[:250]), default=10) + 2))
        sheet.column_dimensions[letter].width = width
    if sheet.max_row > 1:
        sheet.auto_filter.ref = sheet.dimensions


def red_bold_table(records: list[dict[str, Any]]):
    frame = pd.DataFrame(records)
    def format_money(value: Any) -> Any:
        if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
            return f"{value:,.2f}"
        return value

    resident_deposit_money_fields = {
        "total dep. required", "unit deposit", "pet deposit", "other deposits", "deposits on hand",
    }
    money_columns = {
        column: format_money for column in frame.columns
        if "net balance" in column.casefold() or column.casefold() in resident_deposit_money_fields
    }
    def row_style(row):
        if row.get("Bldg/Unit") == "TOTALS" or row.get("Code Description") == "TOTALS":
            return ["color: #12334A; font-weight: bold; background-color: #F3F7FA"] * len(row)
        return ["color: #B42318; font-weight: bold"] * len(row)
    return frame.style.format(money_columns).apply(row_style, axis=1)


st.set_page_config(page_title="OS To OS Validation Tool Report", page_icon="◈", layout="wide")

logo_path = Path(__file__).with_name("OS2OS Validation Tool Finder.png")
if logo_path.is_file():
    st.sidebar.image(str(logo_path), width=145)

mapping_file = Path(__file__).with_name("code_mappings.json")
saved_mappings: dict[str, str] = {
    old.upper(): new
    for report_mappings in CODE_MAPS.values()
    for old, new in report_mappings.items()
}
if mapping_file.exists():
    try:
        with mapping_file.open("r", encoding="utf-8") as file:
            loaded_mappings = json.load(file)
        if isinstance(loaded_mappings, dict):
            saved_mappings = {
                str(old).strip().upper(): str(new).strip()
                for old, new in loaded_mappings.items()
                if str(old).strip() and str(new).strip()
            }
    except (OSError, json.JSONDecodeError):
        st.sidebar.error("Hindi mabasa ang saved code mappings. Ayusin o alisin ang code_mappings.json.")

st.sidebar.title("Dashboard")
page = st.sidebar.radio(
    "Menu",
    ["Home", "Mapping Code", "OS to OS Validation", "Merge File"],
    index=0,
    key="dashboard_navigation",
    label_visibility="collapsed",
)

st.markdown(f"""
<style>
    .stApp {{ background: radial-gradient(ellipse at 85% 0%, #e1f5f2 0, transparent 34%), linear-gradient(145deg, #f7f9fc 0%, #eef4f8 100%); color: #{INK}; }}
    .block-container {{ max-width: 1200px; padding-top: 1.7rem; padding-bottom: 3.5rem; }}
    h1, h2, h3 {{ color: #{NAVY}; letter-spacing: -0.025em; }}
    [data-testid="stSidebar"] {{ background: linear-gradient(180deg, #12334a 0%, #174c5d 100%); }}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2 {{ color: white; }}
    [data-testid="stSidebar"] label {{ color: #e3f2f4; }}
    [data-testid="stSidebar"] [role="radiogroup"] {{ gap: .35rem; }}
    [data-testid="stSidebar"] [role="radio"] {{ border-radius: 10px; padding: .4rem .55rem; }}
    [data-testid="stSidebar"] [role="radio"]:hover {{ background: #ffffff18; }}
    .hero {{ background: linear-gradient(115deg, #{NAVY} 0%, #176d79 62%, #24a19a 100%); color: white; padding: 2rem 2.2rem; border-radius: 22px; margin-bottom: 1.6rem; box-shadow: 0 18px 42px #12334a25; position: relative; overflow: hidden; }}
    .hero:after {{ content: ""; position: absolute; width: 250px; height: 250px; right: -55px; top: -120px; border: 1px solid #ffffff30; border-radius: 50%; box-shadow: 0 0 0 28px #ffffff0d, 0 0 0 58px #ffffff0a; }}
    .hero h1 {{ color: white; font-size: clamp(2rem, 4vw, 3rem); line-height: 1.1; margin: 0; position: relative; z-index: 1; }}
    .hero p {{ color: #e0f1f2; margin: .65rem 0 0; font-size: 1rem; position: relative; z-index: 1; }}
    .hero .eyebrow {{ color: #9ee5d8; font-size: .76rem; font-weight: 750; letter-spacing: .15em; text-transform: uppercase; margin-bottom: .5rem; }}
    .home-hero {{ background: linear-gradient(120deg, #{NAVY}, #176d79 70%, #24a19a); color: white; padding: clamp(1.6rem, 4vw, 2.8rem); border-radius: 24px; margin: .2rem 0 1.5rem; box-shadow: 0 18px 42px #12334a25; }}
    .home-hero .home-kicker {{ color: #a4eee0; font-size: .76rem; font-weight: 750; letter-spacing: .16em; text-transform: uppercase; }}
    .home-hero h1 {{ color: white; font-size: clamp(2.1rem, 4vw, 3.2rem); margin: .5rem 0; line-height: 1.08; }}
    .home-hero p {{ color: #e2f2f3; max-width: 720px; font-size: 1.06rem; margin: 0; }}
    .home-stage {{ position: relative; overflow: hidden; padding: clamp(.7rem, 2.4vw, 1.4rem); border-radius: 24px; margin: .2rem 0 1.5rem; background: radial-gradient(circle at 9% 82%, #e9b18f 0 6%, transparent 6.2%), radial-gradient(circle at 90% 18%, #f0bc9e 0 7%, transparent 7.2%), linear-gradient(135deg, #f6c5aa 0%, #ffe2d1 55%, #f6cbb4 100%); box-shadow: 0 18px 42px #8e4d2a20; }}
    .home-stage:before {{ content: ""; position: absolute; width: 230px; height: 230px; right: 7%; bottom: -170px; border-radius: 50%; background: #e7a98a88; }}
    .home-panel {{ position: relative; z-index: 1; min-height: 390px; display: grid; grid-template-columns: 1.05fr .95fr; align-items: center; gap: clamp(1.5rem, 4vw, 3.2rem); padding: clamp(1.6rem, 5vw, 4.2rem); border: 1px solid #fff9; border-radius: 22px; background: linear-gradient(120deg, #fffaf5ed, #fff4ecdf); box-shadow: 0 16px 38px #79492b20, inset 0 1px 0 #fff; }}
    .home-copy {{ position: relative; z-index: 2; }}
    .home-kicker {{ display: inline-flex; align-items: center; gap: .5rem; color: #a64b2b; font-size: .72rem; font-weight: 800; letter-spacing: .14em; text-transform: uppercase; }}
    .home-kicker:before {{ content: ""; width: 8px; height: 8px; border-radius: 50%; background: #ef8a58; box-shadow: 0 0 0 4px #ef8a5828; }}
    .home-panel h1 {{ color: #29211f; font-size: clamp(2.2rem, 4.3vw, 3.8rem); line-height: 1.03; letter-spacing: -.045em; margin: .85rem 0 1rem; }}
    .home-panel h1 span {{ color: #c65b37; }}
    .home-panel p {{ color: #685955; max-width: 480px; font-size: 1.02rem; line-height: 1.75; margin: 0; }}
    .home-tags {{ display: flex; flex-wrap: wrap; gap: .55rem; margin-top: 1.4rem; }}
    .home-tags span {{ color: #754533; background: #f8e9df; border: 1px solid #efd6c6; border-radius: 999px; padding: .46rem .72rem; font-size: .78rem; font-weight: 650; }}
    .report-visual {{ position: relative; min-height: 320px; display: grid; place-items: center; }}
    .report-art {{ display: block; width: min(125%, 560px); max-width: none; height: auto; margin-left: -8%; overflow: visible; filter: drop-shadow(0 18px 18px #8d563022); }}
    @media (max-width: 760px) {{ .home-panel {{ grid-template-columns: 1fr; gap: .5rem; }} .report-visual {{ min-height: 255px; }} .report-art {{ width: min(100%, 450px); margin-left: 0; }} }}
    [data-testid="stSidebar"] [data-testid="stImage"] img {{ border-radius: 12px; }}
    div[data-testid="stVerticalBlockBorderWrapper"] {{ border-color: #d8e4ea; border-radius: 16px; background: #ffffffc9; box-shadow: 0 8px 22px #18344a0a; }}
    div[data-testid="stMetric"] {{ background: #fff; border: 1px solid #dce7ed; border-radius: 16px; padding: 1rem 1.1rem; box-shadow: 0 8px 20px #18344a0b; }}
    div[data-testid="stMetricLabel"] {{ color: #{MUTED}; }}
    div[data-testid="stMetricValue"] {{ color: #{NAVY}; font-weight: 700; }}
    .stButton>button {{ border-radius: 11px; min-height: 2.7rem; font-weight: 600; transition: transform .15s ease, box-shadow .15s ease; }}
    .stButton>button:hover {{ transform: translateY(-1px); box-shadow: 0 7px 16px #12334a1c; border-color: #56a9a7; }}
    .stButton>button[kind="primary"] {{ background: linear-gradient(100deg, #{TEAL}, #20a095); border-color: #{TEAL}; color: white; box-shadow: 0 8px 18px #147d8326; }}
    [data-testid="stFileUploaderDropzone"] {{ background: #fff; border: 1px dashed #b8ccd6; border-radius: 15px; transition: border-color .15s ease, background .15s ease; }}
    [data-testid="stFileUploaderDropzone"]:hover {{ border-color: #{TEAL}; background: #f4fffd; }}
    .step-label {{ color: #{TEAL}; font-size: .75rem; font-weight: 800; letter-spacing: .14em; text-transform: uppercase; margin-bottom: .4rem; }}
    .stCaption {{ color: #{MUTED}; }}
    .stTabs [data-baseweb="tab-list"] {{ gap: .5rem; border-bottom: 1px solid #dce5ec; }}
    .stTabs [data-baseweb="tab"] {{ border-radius: 10px 10px 0 0; padding: .65rem 1rem; }}
    [data-testid="stDataFrame"] {{ border: 1px solid #dce5ec; border-radius: 14px; overflow: hidden; }}
</style>
""", unsafe_allow_html=True)

if page == "Mapping Code":
    st.title("Mapping Code")
    st.markdown("**Old Site Code Mapping to New Site Code Mapping**")
    st.caption('Edit the codes, click "-" to delete a row, and click "Save mappings" to save your changes.')
    if "mapping_rows" not in st.session_state:
        st.session_state["mapping_rows"] = [
            {"id": index, "old": old, "new": new}
            for index, (old, new) in enumerate(saved_mappings.items())
        ]
        st.session_state["next_mapping_row_id"] = len(st.session_state["mapping_rows"])

    def sync_mapping_rows() -> list[dict[str, Any]]:
        rows = st.session_state["mapping_rows"]
        for row in rows:
            old_key = f"mapping_old_{row['id']}"
            new_key = f"mapping_new_{row['id']}"
            if old_key in st.session_state:
                row["old"] = st.session_state[old_key]
            if new_key in st.session_state:
                row["new"] = st.session_state[new_key]
        return rows

    sync_mapping_rows()
    heading_cols = st.columns([0.44, 0.44, 0.12])
    heading_cols[0].caption("Old Site")
    heading_cols[1].caption("New Site")
    for row in st.session_state["mapping_rows"]:
        old_col, new_col, delete_col = st.columns([0.44, 0.44, 0.12])
        with old_col:
            st.text_input("Old Site Code", value=row["old"], key=f"mapping_old_{row['id']}", label_visibility="collapsed", placeholder="Old code")
        with new_col:
            st.text_input("New Site Code", value=row["new"], key=f"mapping_new_{row['id']}", label_visibility="collapsed", placeholder="New code")
        with delete_col:
            if st.button("-", key=f"delete_mapping_{row['id']}", help="Delete this mapping row"):
                sync_mapping_rows()
                st.session_state["mapping_rows"] = [item for item in st.session_state["mapping_rows"] if item["id"] != row["id"]]
                st.rerun()

    if st.button("+ Add mapping", key="add_mapping_row"):
        sync_mapping_rows()
        new_id = st.session_state["next_mapping_row_id"]
        st.session_state["next_mapping_row_id"] = new_id + 1
        st.session_state["mapping_rows"].append({"id": new_id, "old": "", "new": ""})
        st.rerun()

    if st.button("Save mappings", key="save_mappings", type="primary"):
        mapping_rows = sync_mapping_rows()
        clean_rows = []
        incomplete_rows = 0
        for row in mapping_rows:
            old_code = str(row["old"]).strip()
            new_code = str(row["new"]).strip()
            if not old_code and not new_code:
                continue
            if not old_code or not new_code:
                incomplete_rows += 1
                continue
            clean_rows.append((old_code.upper(), new_code))
        old_codes = [old for old, _ in clean_rows]
        if len(old_codes) != len(set(old_codes)):
            st.error("May duplicate na Old Site Code. Bawat Old Site Code ay dapat isang beses lang.")
        elif incomplete_rows:
            st.error("Kumpletuhin o burahin ang mga row na kulang ang Old Site o New Site code bago mag-save.")
        else:
            try:
                saved_mappings = dict(clean_rows)
                with mapping_file.open("w", encoding="utf-8") as file:
                    json.dump(saved_mappings, file, indent=2, ensure_ascii=False)
                st.success("Nai-save ang code mappings.")
            except OSError as exc:
                st.error(f"Hindi ma-save ang mappings: {exc}")

custom_code_map = saved_mappings
if page == "Mapping Code":
    st.stop()
if page == "Home":
    st.markdown("""
    <section class="home-stage">
      <div class="home-panel">
        <div class="home-copy">
          <div class="home-kicker">OS2OS · Report validation workspace</div>
          <h1>Make every report comparison <span>clear.</span></h1>
          <p>Compare exports from your old and new systems, spot mismatched values, and find records missing from either file—all in one review-ready workflow.</p>
          <div class="home-tags"><span>Excel reports</span><span>PDF for Delinquent &amp; Prepaid</span><span>Downloadable results</span></div>
        </div>
        <div class="report-visual">
          <svg class="report-art" viewBox="0 0 560 390" role="img" aria-label="Illustration of a man using a magnifying glass to inspect Excel and PDF reports">
            <defs>
              <linearGradient id="boardFill" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#fffefa"/><stop offset="1" stop-color="#fff2e9"/></linearGradient>
              <linearGradient id="shirtFill" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#ffc86d"/><stop offset="1" stop-color="#f19a48"/></linearGradient>
              <linearGradient id="pantsFill" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#397d91"/><stop offset="1" stop-color="#24556d"/></linearGradient>
              <linearGradient id="hairFill" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#59413a"/><stop offset="1" stop-color="#302b31"/></linearGradient>
              <filter id="softShadow" x="-30%" y="-30%" width="160%" height="180%"><feDropShadow dx="0" dy="9" stdDeviation="9" flood-color="#744a36" flood-opacity=".16"/></filter>
            </defs>
            <circle cx="290" cy="186" r="154" fill="#f5c4a8" opacity=".28"/>
            <circle cx="500" cy="73" r="17" fill="#e99362" opacity=".78"/>
            <circle cx="65" cy="288" r="24" fill="#e9a37e" opacity=".38"/>
            <g filter="url(#softShadow)">
              <rect x="267" y="83" width="260" height="215" rx="22" fill="url(#boardFill)" stroke="#f0d9c9" stroke-width="2"/>
              <circle cx="292" cy="110" r="5" fill="#e99462"/>
              <text x="306" y="114" fill="#8c6251" font-family="Arial, sans-serif" font-size="11" font-weight="700" letter-spacing="1.5">FILE VALIDATION</text>
              <rect x="438" y="99" width="67" height="23" rx="11.5" fill="#e4f2e9"/>
              <text x="471.5" y="114" text-anchor="middle" fill="#39836a" font-family="Arial, sans-serif" font-size="9" font-weight="700">READY</text>
              <path d="M285 137H509" stroke="#f0e5dc" stroke-width="1.5"/>
              <!-- Excel report -->
              <path d="M300 154H341L356 169V250H300Z" fill="#fff" stroke="#dce8df" stroke-width="2"/>
              <path d="M341 154V170H356" fill="#e6f3e9" stroke="#dce8df" stroke-width="2" stroke-linejoin="round"/>
              <rect x="306" y="178" width="22" height="28" rx="3" fill="#27845b"/>
              <path d="M312 185L322 199M322 185L312 199" stroke="#fff" stroke-width="2.3" stroke-linecap="round"/>
              <text x="334" y="188" fill="#52705e" font-family="Arial, sans-serif" font-size="8" font-weight="700">EXCEL</text>
              <path d="M334 196H349M334 202H349M306 215H349M306 222H349M306 229H349" stroke="#d5e5d9" stroke-width="2" stroke-linecap="round"/>
              <!-- PDF report -->
              <path d="M370 154H411L426 169V250H370Z" fill="#fff" stroke="#eadbd6" stroke-width="2"/>
              <path d="M411 154V170H426" fill="#f9e9e5" stroke="#eadbd6" stroke-width="2" stroke-linejoin="round"/>
              <path d="M378 180H398L404 186V204H378Z" fill="#d85f4a"/>
              <text x="391" y="197" text-anchor="middle" fill="#fff" font-family="Arial, sans-serif" font-size="10" font-weight="700">PDF</text>
              <path d="M378 215H418M378 222H418M378 229H411" stroke="#eadbd6" stroke-width="2" stroke-linecap="round"/>
              <!-- magnifying glass over both report types -->
              <path d="M447 229L479 262" stroke="#31566a" stroke-width="13" stroke-linecap="round"/>
              <path d="M447 229L479 262" stroke="#6f9ba0" stroke-width="5" stroke-linecap="round"/>
              <circle cx="431" cy="213" r="39" fill="#fffefa" fill-opacity=".87" stroke="#31566a" stroke-width="8"/>
              <circle cx="431" cy="213" r="30" fill="#f3faf7" fill-opacity=".72" stroke="#fff" stroke-width="2"/>
              <path d="M418 214L427 223L445 202" fill="none" stroke="#4caa88" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/>
              <path d="M286 276H504" stroke="#eee2d9" stroke-width="2" stroke-linecap="round"/>
            </g>
            <ellipse cx="182" cy="351" rx="74" ry="12" fill="#9a654a" opacity=".13"/>
            <g stroke-linecap="round" stroke-linejoin="round">
              <!-- legs and shoes -->
              <path d="M159 229Q171 232 184 230L181 324Q170 329 156 325Z" fill="url(#pantsFill)"/>
              <path d="M187 230Q202 232 218 229L225 325Q210 329 194 324Z" fill="url(#pantsFill)"/>
              <path d="M154 337Q165 332 181 340L180 350Q164 356 148 348Q147 342 154 337Z" fill="#fff8ed" stroke="#d8b79f" stroke-width="2"/>
              <path d="M211 337Q222 332 235 337L242 347Q230 354 213 349Q208 345 211 337Z" fill="#fff8ed" stroke="#d8b79f" stroke-width="2"/>
              <!-- short hair, neck and face -->
              <path d="M151 91Q148 53 181 49Q213 51 216 84L210 94L207 78Q197 77 191 69Q177 82 154 78Z" fill="url(#hairFill)"/>
              <path d="M171 112L171 132Q182 143 195 131L193 110Z" fill="#e9aa7b"/>
              <ellipse cx="183" cy="88" rx="31" ry="38" fill="#f3bd8e"/>
              <path d="M153 83Q153 55 180 52Q205 52 212 75Q198 70 190 65Q177 78 155 76Z" fill="#493631"/>
              <circle cx="173" cy="89" r="2.7" fill="#463932"/><circle cx="195" cy="89" r="2.7" fill="#463932"/>
              <path d="M177 103Q183 108 190 102" fill="none" stroke="#a85e52" stroke-width="2.2"/>
              <!-- arms -->
              <path d="M162 151Q147 159 140 180L125 194Q119 200 124 207Q130 211 137 204L163 187L177 163Z" fill="#f3bd8e"/>
              <path d="M213 153Q229 154 239 168L252 183Q260 189 255 196Q249 201 242 194L220 180L206 166Z" fill="#f3bd8e"/>
              <!-- shirt -->
              <path d="M163 137Q181 128 199 136L217 149L210 204Q190 216 157 202L154 157Z" fill="url(#shirtFill)"/>
              <path d="M176 136L184 149L192 136" fill="#ffe0a2" opacity=".95"/>
              <!-- trousers waistband and seams -->
              <path d="M157 199Q183 211 211 202L218 237Q190 250 157 238Z" fill="#2c6880"/>
              <path d="M171 240L168 321M202 240L209 321" fill="none" stroke="#70a5ad" stroke-width="2" opacity=".75"/>
              <!-- tablet -->
              <g transform="rotate(-8 133 185)">
                <rect x="115" y="167" width="43" height="34" rx="5" fill="#31546c"/>
                <rect x="119" y="171" width="35" height="26" rx="3" fill="#eef7f5"/>
                <path d="M125 180H147M125 185H141" stroke="#9bbdb4" stroke-width="1.8" stroke-linecap="round"/>
                <path d="M125 191L130 195L138 186" fill="none" stroke="#55a891" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
              </g>
            </g>
            <circle cx="120" cy="101" r="7" fill="#ed9a64" opacity=".8"/>
            <path d="M118 100L120 96L122 100L126 101L122 103L120 107L118 103L114 101Z" fill="#f4bd6e"/>
          </svg>
        </div>
      </div>
    </section>
    """, unsafe_allow_html=True)
    st.subheader("A simple three-step workflow")
    home_cols = st.columns(3)
    with home_cols[0]:
        with st.container(border=True):
            st.markdown("#### ① Choose a report")
            st.write("Open **OS to OS Validation** and select the report type you need to review.")
    with home_cols[1]:
        with st.container(border=True):
            st.markdown("#### ② Add both files")
            st.write("Upload matching Old Site and New Site Excel reports, or selectable-text PDFs for Delinquent and Prepaid.")
    with home_cols[2]:
        with st.container(border=True):
            st.markdown("#### ③ Review results")
            st.write("Review differences and missing records, then download the comparison report.")
    st.stop()
if page == "Merge File":
    st.title("Merge File")
    st.info("Coming soon — under construction.")
    st.stop()

task = "OS to OS"


@st.dialog("Report type warning")
def show_report_type_warning(report_type: str, filenames: list[str]) -> None:
    st.warning("Validation was stopped because these file names do not appear to match the selected report type.")
    st.write(f"Selected report type: **{report_type}**")
    for filename in filenames:
        st.write(f"• {filename}")
    st.caption("Choose the matching report type or upload the correct Old Site and New Site reports, then compare again.")
    if st.button("OK", type="primary"):
        st.rerun()

st.markdown(f"""
<div class="hero"><h1>OS To OS Validation Tool Report</h1><p>Compare two site exports, find missing records, and review field level differences.</p></div>
""", unsafe_allow_html=True)

comparison_signature = None
st.markdown('<div class="step-label">1 · Configure comparison</div>', unsafe_allow_html=True)
st.caption("Select a report type, then upload both matching Excel files or selectable-text PDFs for Delinquent and Prepaid.")
top_left, top_right = st.columns([1, 1])
with top_left:
    report_type = st.selectbox("Report type", ["Please select report", *REPORTS], index=0, key="report_type_select_v2")
with top_right:
    st.info("Upload Old Site and New Site files that match the report type selected here.", icon="ℹ️")

st.markdown('<div class="step-label" style="margin-top:1rem">2 · Add two different files</div>', unsafe_allow_html=True)
upload_left, upload_right = st.columns(2)
with upload_left:
    st.markdown("#### Old Site Report")
    old_upload = st.file_uploader("Upload the Old Site Report", type=["xls", "xlsx", "pdf"], key="old_file", label_visibility="collapsed")
    st.caption("Previous system export")
with upload_right:
    st.markdown("#### New Site (Test Site Report)")
    new_upload = st.file_uploader("Upload the New Site (Test Site Report)", type=["xls", "xlsx", "pdf"], key="new_file", label_visibility="collapsed")
    st.caption("Current system export")

if report_type == "Please select report":
    if old_upload or new_upload:
        st.info("Please select a report type before uploading or comparing files.")
elif old_upload and new_upload:
    config = REPORTS[report_type]
    try:
        old_raw = read_first_sheet(old_upload.getvalue(), old_upload.name, report_type)
        new_raw = read_first_sheet(new_upload.getvalue(), new_upload.name, report_type)
        if report_type == "Market Rent Evaluation":
            old_header_row = new_header_row = 0
            old_rows = extract_market_rent_data(old_raw)
            new_rows = extract_market_rent_data(new_raw)
            old_headers = new_headers = MARKET_RENT_FIELDS
            key_pairs = [(0, 0)]
            validation_pairs = [(index, index) for index in range(1, len(MARKET_RENT_FIELDS))]
            st.markdown('<div class="step-label" style="margin-top:1rem">3 · Market rent fields</div>', unsafe_allow_html=True)
            st.caption("Units are matched by Bldg/Unit. Description/Amenities, square footage, base rent, amenity value, and market rent are included in the comparison and downloadable workbook.")
        else:
            old_header_guess = guess_header_row(old_raw, config["keys"] + config["validate"])
            new_header_guess = guess_header_row(new_raw, config["keys"] + config["validate"])
            st.markdown('<div class="step-label" style="margin-top:1rem">3 · Confirm the report layout</div>', unsafe_allow_html=True)
            with st.expander("Advanced column setup", expanded=False):
                row_col1, row_col2 = st.columns(2)
                with row_col1:
                    old_header_row = st.number_input("Old Site Report header row", min_value=1, max_value=max(1, len(old_raw)), value=old_header_guess + 1, key="old_header") - 1
                with row_col2:
                    new_header_row = st.number_input("New Site (Test Site Report) header row", min_value=1, max_value=max(1, len(new_raw)), value=new_header_guess + 1, key="new_header") - 1
            old_headers, old_rows = clean_frame(old_raw, int(old_header_row))
            new_headers, new_rows = clean_frame(new_raw, int(new_header_row))
            with st.expander("Matching and validation columns", expanded=False):
                st.caption("Defaults are based on the report setup in your workbook. Adjust the paired columns if either export uses a different layout.")
                st.markdown("**Match rows using**")
                report_key = report_type.replace(" ", "_")
                if report_type == "Resident Deposit":
                    # New namespace drops previously saved field choices so Status uses its configured column.
                    report_key += "_status_fields"
                key_pairs = choose_indices(old_headers, new_headers, config["keys"], f"{report_key}_keys", "Match key")
                st.markdown("**Compare values in**")
                validation_pairs = choose_indices(old_headers, new_headers, config["validate"], f"{report_key}_validation", "Field")
        comparison_signature = (
            hashlib.sha256(old_upload.getvalue()).hexdigest(),
            hashlib.sha256(new_upload.getvalue()).hexdigest(),
            report_type, task, int(old_header_row), int(new_header_row),
            tuple(key_pairs), tuple(validation_pairs), tuple(sorted(custom_code_map.items())),
        )
        compare_col, clear_col, status_col = st.columns([1, 1, 3])
        with compare_col:
            run = st.button("Compare & validate", type="primary", use_container_width=True)
        with clear_col:
            clear_results = st.button("Clear results", use_container_width=True)
        with status_col:
            st.caption(f"{len(old_rows):,} Old Site Report records · {len(new_rows):,} New Site (Test Site Report) records · Excel: first worksheet | PDF: detected report tables")
        if clear_results:
            st.session_state.pop("comparison_result", None)
            st.session_state.pop("comparison_meta", None)
            st.session_state.pop("comparison_signature", None)
            st.rerun()
        if run:
            filename_mismatches = [
                upload.name
                for upload in (old_upload, new_upload)
                if not filename_matches_report(upload.name, report_type)
            ]
            if filename_mismatches:
                st.session_state.pop("comparison_result", None)
                st.session_state.pop("comparison_meta", None)
                st.session_state.pop("comparison_signature", None)
                show_report_type_warning(report_type, filename_mismatches)
            elif not key_pairs:
                st.error("Pumili ng kahit isang matching column.")
            elif not validation_pairs:
                st.error("Walang validation fields na naka-configure para sa report type na ito.")
            else:
                code_map = custom_code_map
                compare_old, compare_new = old_rows, new_rows
                compare_old_headers, compare_new_headers = old_headers, new_headers
                compare_key_pairs, compare_validation_pairs = key_pairs, validation_pairs
                compare_code_map = code_map
                if report_type == "Delinquent and Prepaid":
                    name_pair = selected_field_pair(old_headers, new_headers, validation_pairs, "name")
                    status_pair = selected_field_pair(old_headers, new_headers, validation_pairs, "status")
                    code_pair = selected_field_pair(old_headers, new_headers, validation_pairs, "code description")
                    prepaid_old = find_header_index(old_headers, "Total Prepaid")
                    prepaid_new = find_header_index(new_headers, "Total Prepaid")
                    delinquent_old = find_header_index(old_headers, "Total Delinquent")
                    delinquent_new = find_header_index(new_headers, "Total Delinquent")
                    if not name_pair or not status_pair or not code_pair:
                        raise ValueError("Delinquent and Prepaid report needs Name, Status, and Code Description validation columns.")
                    if None in (prepaid_old, prepaid_new, delinquent_old, delinquent_new):
                        raise ValueError("Hindi makita ang Total Prepaid at Total Delinquent columns sa isa o parehong report.")
                    old_key_column, new_key_column = key_pairs[0]
                    compare_old_headers = compare_new_headers = ["Bldg/Unit", "Name", "Status", "Code Description", "Net Balance"]
                    compare_old = prepare_delinquent_prepaid(
                        old_rows, old_headers, old_key_column, name_pair[0], status_pair[0], code_pair[0],
                        prepaid_old, delinquent_old, code_map, True,
                    )
                    compare_new = prepare_delinquent_prepaid(
                        new_rows, new_headers, new_key_column, name_pair[1], status_pair[1], code_pair[1],
                        prepaid_new, delinquent_new, {}, False,
                    )
                    compare_key_pairs = [(0, 0)]
                    compare_validation_pairs = [(1, 1), (2, 2), (3, 3), (4, 4)]
                    compare_code_map = {}

                result = compare_frames(
                    compare_old, compare_new, compare_old_headers, compare_new_headers,
                    compare_key_pairs, compare_validation_pairs, compare_code_map,
                )
                old_output_columns = list(dict.fromkeys([old_i for old_i, _ in compare_key_pairs + compare_validation_pairs]))
                new_output_columns = list(dict.fromkeys([new_i for _, new_i in compare_key_pairs + compare_validation_pairs]))
                missing_old_records = row_records(compare_old, result["missing_old"], compare_old_headers, old_output_columns)
                missing_new_records = row_records(compare_new, result["missing_new"], compare_new_headers, new_output_columns)
                result["missing_old_records"] = missing_old_records
                result["missing_new_records"] = missing_new_records
                result["validation_pairs"] = compare_validation_pairs
                if report_type == "Market Rent Evaluation":
                    result["market_rent_old"] = old_rows
                    result["market_rent_new"] = new_rows
                st.session_state["comparison_result"] = result
                st.session_state["comparison_meta"] = (report_type, task)
                st.session_state["comparison_signature"] = comparison_signature
    except Exception as exc:
        st.error(str(exc))
elif old_upload or new_upload:
    st.info("Upload both files to begin the comparison.")

if "comparison_result" in st.session_state and st.session_state.get("comparison_signature") == comparison_signature:
    result = st.session_state["comparison_result"]
    result_type, result_task = st.session_state["comparison_meta"]
    st.divider()
    st.markdown('<div class="step-label">Comparison result</div>', unsafe_allow_html=True)
    is_clean = not (result["differences"] or result["missing_old"] or result["missing_new"])
    if is_clean:
        st.success("All matched values agree. Walang nawawalang rows.")
    else:
        st.warning("Differences were found or records are missing from the other file.")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Matched records", f"{len(result['matched']):,}")
    c2.metric("Mismatched records", f"{len(result['differences']):,}")
    c3.metric("Missing from new", f"{len(result['missing_old']):,}")
    c4.metric("Missing from old", f"{len(result['missing_new']):,}")
    if result_type == "Market Rent Evaluation":
        old_market = result["market_rent_old"]
        new_market = result["market_rent_new"]
        old_sqft = market_field_total(old_market, "Unit SQFT")
        new_sqft = market_field_total(new_market, "Unit SQFT")
        old_base_rent = market_field_total(old_market, "Base Rent")
        new_base_rent = market_field_total(new_market, "Base Rent")
        total_sqft_old, total_sqft_new, total_rent_old, total_rent_new = st.columns(4)
        total_sqft_old.metric("Old Site Total SQFT", f"{old_sqft:,.0f}")
        total_sqft_new.metric("New Site Total SQFT", f"{new_sqft:,.0f}")
        total_rent_old.metric("Old Site Total Base Rent", f"${old_base_rent:,.2f}")
        total_rent_new.metric("New Site Total Base Rent", f"${new_base_rent:,.2f}")
    if result_type == "Resident Deposit":
        old_rows, new_rows = result["old_rows"], result["new_rows"]
        old_headers, new_headers = result["old_headers"], result["new_headers"]
        old_required = resident_deposit_total(old_rows, old_headers, "Total Dep. Required")
        new_required = resident_deposit_total(new_rows, new_headers, "Total Dep. Required")
        old_on_hand = resident_deposit_total(old_rows, old_headers, "Deposits On Hand")
        new_on_hand = resident_deposit_total(new_rows, new_headers, "Deposits On Hand")
        old_unit_deposit = resident_deposit_total(old_rows, old_headers, "Unit Deposit")
        new_unit_deposit = resident_deposit_total(new_rows, new_headers, "Unit Deposit")
        old_pet_deposit = resident_deposit_total(old_rows, old_headers, "Pet Deposit")
        new_pet_deposit = resident_deposit_total(new_rows, new_headers, "Pet Deposit")
        old_other_deposits = resident_deposit_total(old_rows, old_headers, "Other Deposits")
        new_other_deposits = resident_deposit_total(new_rows, new_headers, "Other Deposits")
        required_old, required_new, on_hand_old, on_hand_new = st.columns(4)
        required_old.metric("Old Site Total Dep. Required", f"${old_required:,.2f}")
        required_new.metric("New Site Total Dep. Required", f"${new_required:,.2f}")
        on_hand_old.metric("Old Site Deposits On Hand", f"${old_on_hand:,.2f}")
        on_hand_new.metric("New Site Deposits On Hand", f"${new_on_hand:,.2f}")
        st.markdown("**Deposit breakdown totals · Old Site**")
        old_component_cards = st.columns(3)
        old_component_cards[0].metric("Unit Deposit", f"${old_unit_deposit:,.2f}")
        old_component_cards[1].metric("Pet Deposit", f"${old_pet_deposit:,.2f}")
        old_component_cards[2].metric("Other Deposits", f"${old_other_deposits:,.2f}")
        st.markdown("**Deposit breakdown totals · New Site**")
        new_component_cards = st.columns(3)
        new_component_cards[0].metric("Unit Deposit", f"${new_unit_deposit:,.2f}")
        new_component_cards[1].metric("Pet Deposit", f"${new_pet_deposit:,.2f}")
        new_component_cards[2].metric("Other Deposits", f"${new_other_deposits:,.2f}")
    if result_type == "Delinquent and Prepaid" and result.get("old_net_balance_total") is not None and result.get("new_net_balance_total") is not None:
        total_old_col, total_new_col = st.columns(2)
        total_old_col.metric("Old Site Net Balance Total", f"${money_display(result['old_net_balance_total'])}")
        total_new_col.metric("New Site Net Balance Total", f"${money_display(result['new_net_balance_total'])}")
    tab_diff, tab_missing_new, tab_missing_old = st.tabs(["Mismatched records", "Missing from new", "Missing from Old Site Report"])
    with tab_diff:
        displayed_differences = result_display_rows(result, result_type)
        if displayed_differences:
            differences_view = red_bold_table(displayed_differences)
            st.dataframe(differences_view, use_container_width=True, hide_index=True)
        else:
            st.caption("No mismatched records found.")
    with tab_missing_new:
        st.caption("Records from the Old Site Report not found in the New Site (Test Site Report).")
        if result["missing_old_records"]:
            missing_old_records = result["missing_old_records"]
            if result_type == "Resident Deposit":
                missing_old_records = resident_deposit_display_rows(missing_old_records)
            missing_new_view = red_bold_table(add_net_balance_table_total(missing_old_records))
            st.dataframe(missing_new_view, use_container_width=True, hide_index=True)
        else:
            st.caption("Walang missing records.")
    with tab_missing_old:
        st.caption("Records from the New Site (Test Site Report) not found in the Old Site Report.")
        if result["missing_new_records"]:
            missing_new_records = result["missing_new_records"]
            if result_type == "Resident Deposit":
                missing_new_records = resident_deposit_display_rows(missing_new_records)
            missing_old_view = red_bold_table(add_net_balance_table_total(missing_new_records))
            st.dataframe(missing_old_view, use_container_width=True, hide_index=True)
        else:
            st.caption("Walang missing records.")
    report_bytes = make_report(result, result_type, result_task)
    st.download_button("Download comparison report", data=report_bytes, file_name="OS2OS_Validation_Results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

st.markdown("---")
