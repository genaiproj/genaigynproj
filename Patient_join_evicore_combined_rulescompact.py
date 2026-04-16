#!/usr/bin/env python3
"""
Patient_join_evicore_combined_mapping_inputfolder.py

This version drops the old normalized diabetes/obesity CSV rule inputs.
Instead, it reads the shared lookup workbook from the input folder and uses
that sheet as the source of truth for expected_count, actual_tracking,
and actual_tracking_details.

What changed:
- expected_count is now based on the test description schedule, not CPT volume.
- actual_tracking counts completed schedule occurrences, not the number of CPT codes hit.
- actual_tracking_details still shows the matching CPT/date evidence.
- mapping lookup is read from input/MultipleConditions_Evicore_Mapping.xls
  (or whatever is passed in with --mapping-blob).
"""

# Update : 13-04-2026 for soring and renam output file as per mail from Satarupa dated 13-04-2026

import argparse
import os
import re
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from io import BytesIO
from typing import Any, Dict, Iterable, List, Optional, Tuple, Union

import pandas as pd
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient

try:
    from zoneinfo import ZoneInfo
except Exception:  # pragma: no cover
    ZoneInfo = None


CPT_RE = re.compile(r"\b(\d{5})\b")
#STEP4_TS_RE = re.compile(r"step4_(\d{8})_(\d{6})", re.IGNORECASE)
#added 13-04-2026
STEP4_TS_RE = re.compile(r"WHF_(\d{8})_(\d{6})", re.IGNORECASE)
NUMBERED_ITEM_RE = re.compile(r"(?:(?<=\n)|^)[ \t]*\d+[.)][ \t]*")

DIABETES_GROUPS = {
    "pre_existing",
    "diet_controlled",
    "oral_or_insulin",
}

COLUMN_MAP = {
    "Pre existing Diabetes": ("DIABETES", "pre_existing"),
    "Diabetes with diet control": ("DIABETES", "diet_controlled"),
    "Diabetes with oral medication or Insulin": ("DIABETES", "oral_or_insulin"),
    "Obesity -class 1": ("OBESITY", "OBESITY|OBESITY_CLASS_I"),
    "Obesity -class 2": ("OBESITY", "OBESITY|OBESITY_CLASS_II"),
    "Obesity- class 3": ("OBESITY", "OBESITY|OBESITY_CLASS_III"),
    "Obesity -class 4": ("OBESITY", "OBESITY|OBESITY_CLASS_IV"),
    "Hypertension": ("HYPERTENSION", "HYPERTENSION"),
}


@dataclass(frozen=True)
class BlobLoc:
    container: str
    blob: str


@dataclass
class Rule:
    program: str
    applies_to_groups: str
    test_name: str
    cpt_codes: str
    start_week: Optional[float]
    end_week: Optional[float]
    frequency_type: str
    frequency_value: str
    stop_condition: str = ""
    special_notes: str = ""


def log(msg: str) -> None:
    print(msg)


def normalize_patient_id(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def normalize_cpt(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    match = CPT_RE.search(str(value))
    return match.group(1) if match else str(value).strip()


def to_dt(value: Any) -> pd.Timestamp:
    return pd.to_datetime(value, errors="coerce")


def strip_tz(series: pd.Series) -> pd.Series:
    s = pd.to_datetime(series, errors="coerce")
    try:
        if getattr(s.dt, "tz", None) is not None:
            return s.dt.tz_convert(None)
    except Exception:
        pass
    return s


def fmt_mdy(ts: pd.Timestamp) -> str:
    if pd.isna(ts):
        return ""
    try:
        return ts.strftime("%-m/%-d/%Y")
    except Exception:
        return ts.strftime("%m/%d/%Y")


def as_float(value: Any) -> Optional[float]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null"}:
        return None
    try:
        return float(text)
    except Exception:
        return None


def blob_service_client(account_url: str) -> BlobServiceClient:
    conn = os.getenv("AZURE_STORAGE_CONNECTION_STRING", "").strip()
    if conn:
        return BlobServiceClient.from_connection_string(conn)
    cred = DefaultAzureCredential(exclude_interactive_browser_credential=True)
    return BlobServiceClient(account_url=account_url, credential=cred)


def download_bytes(bsc: BlobServiceClient, loc: BlobLoc) -> bytes:
    bc = bsc.get_blob_client(container=loc.container, blob=loc.blob)
    return bc.download_blob().readall()


def upload_bytes(bsc: BlobServiceClient, loc: BlobLoc, data: bytes) -> None:
    bc = bsc.get_blob_client(container=loc.container, blob=loc.blob)
    bc.upload_blob(data, overwrite=True)


def step4_dt_from_name(name: str) -> Optional[datetime]:
    match = STEP4_TS_RE.search(name or "")
    if not match:
        return None
    try:
        return datetime.strptime(f"{match.group(1)}_{match.group(2)}", "%Y%m%d_%H%M%S").replace(tzinfo=timezone.utc)
    except Exception:
        return None


def pick_latest_step4(bsc: BlobServiceClient, container: str, prefix: str) -> str:
    prefix = (prefix or "").lstrip("/")
    if prefix and not prefix.endswith("/"):
        prefix += "/"

    cc = bsc.get_container_client(container)
    candidates: List[Tuple[Optional[datetime], Optional[datetime], str]] = []

    #for blob in cc.list_blobs(name_starts_with=prefix + "step4_"):
    for blob in cc.list_blobs(name_starts_with=prefix + "WHF"):
        name = blob.name
        lowered = name.lower()
        #if "care_tracking_output" not in lowered:
        #added 13-04-2026
        if "current_pregnant_patients" not in lowered:
            continue
        if not (lowered.endswith(".csv") or lowered.endswith(".csv.csv")):
            continue

        parsed_ts = step4_dt_from_name(name)
        last_modified = blob.last_modified
        if last_modified and last_modified.tzinfo is None:
            last_modified = last_modified.replace(tzinfo=timezone.utc)
        candidates.append((parsed_ts, last_modified, name))

    if not candidates:
        raise FileNotFoundError(
            f"No Step4 Care Tracking output found in container '{container}' under prefix '{prefix or '(root)'}'."
        )

    candidates.sort(
        key=lambda item: (
            item[0] or datetime.min.replace(tzinfo=timezone.utc),
            item[1] or datetime.min.replace(tzinfo=timezone.utc),
        )
    )
    return candidates[-1][2]


def load_step4_df(raw_csv: bytes) -> pd.DataFrame:
    df = pd.read_csv(BytesIO(raw_csv))

    if "patient_id" not in df.columns:
        for candidate in ["patientid", "PatientID", "PATIENT_ID", "MRN"]:
            if candidate in df.columns:
                df = df.rename(columns={candidate: "patient_id"})
                break
    if "patient_id" not in df.columns:
        raise KeyError(f"Step4 file is missing patient_id. Columns found: {list(df.columns)}")

    df["patient_id"] = df["patient_id"].apply(normalize_patient_id)

    if "start_of_pregnancy_date" not in df.columns:
        raise KeyError("Step4 file is missing start_of_pregnancy_date.")
    df["start_of_pregnancy_date"] = strip_tz(df["start_of_pregnancy_date"])

    if "BMI_at_pregnancy_start_num" not in df.columns:
        if "BMI_at_pregnancy_start" in df.columns:
            df["BMI_at_pregnancy_start_num"] = pd.to_numeric(df["BMI_at_pregnancy_start"], errors="coerce")
        else:
            df["BMI_at_pregnancy_start_num"] = pd.NA

    for col in ["category", "category_reason", "pregnancy_complications", "has_fgr"]:
        if col not in df.columns:
            df[col] = ""

    return df


def load_encounters_df(raw_xlsx: bytes) -> pd.DataFrame:
    enc = pd.read_excel(BytesIO(raw_xlsx))

    def find_col(candidates: Iterable[str]) -> str:
        lowered = {c.lower(): c for c in enc.columns}
        for cand in candidates:
            if cand in enc.columns:
                return cand
            if cand.lower() in lowered:
                return lowered[cand.lower()]
        raise KeyError(f"Missing encounter column. Tried {list(candidates)}. Found {list(enc.columns)}")

    pid_col = find_col(["patientid", "patient_id", "PatientID", "PATIENT_ID", "MRN"])
    cpt_col = find_col(["CPTCode", "CPT_CODE", "CPT", "cpt"])
    dos_col = find_col(["DOS", "Date of Service", "Service Date"])

    enc = enc.rename(columns={pid_col: "enc_patient_id", cpt_col: "enc_cpt", dos_col: "enc_dos"})
    enc["enc_patient_id"] = enc["enc_patient_id"].apply(normalize_patient_id)
    enc["enc_cpt"] = enc["enc_cpt"].apply(normalize_cpt)
    enc["enc_dos"] = strip_tz(enc["enc_dos"])

    return enc[["enc_patient_id", "enc_cpt", "enc_dos"]].copy()


def build_enc_index(enc: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    out: Dict[str, pd.DataFrame] = {}
    enc = enc.dropna(subset=["enc_patient_id"]).copy()
    enc["enc_patient_id"] = enc["enc_patient_id"].astype(str)
    for patient_id, group in enc.groupby("enc_patient_id"):
        out[patient_id] = group.sort_values(["enc_dos", "enc_cpt"]).reset_index(drop=True)
    return out


def parse_range(value: str) -> Optional[Tuple[int, int]]:
    nums = [int(x) for x in re.findall(r"(\d+)", str(value or ""))]
    if len(nums) >= 2:
        return min(nums[0], nums[1]), max(nums[0], nums[1])
    if len(nums) == 1:
        return nums[0], nums[0]
    return None


def parse_row_week(cell_value: Any) -> Tuple[Optional[float], Optional[float]]:
    if cell_value is None or (isinstance(cell_value, float) and pd.isna(cell_value)):
        return None, None

    text = str(cell_value).strip().lower()
    nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", text)]
    if not nums:
        return None, None

    if len(nums) >= 2:
        return min(nums[0], nums[1]), max(nums[0], nums[1])

    week = nums[0]
    if ">=" in text or ">" in text:
        return week, 40.0
    return week, week


def split_mapping_cell(cell_value: Any) -> List[str]:
    if cell_value is None or (isinstance(cell_value, float) and pd.isna(cell_value)):
        return []

    text = str(cell_value).strip()
    if not text:
        return []

    if NUMBERED_ITEM_RE.search(text):
        parts = [p.strip(" \n\t.-") for p in NUMBERED_ITEM_RE.split(text) if p.strip()]
        if parts:
            return parts

    lines = [line.strip(" \n\t.-") for line in text.splitlines() if line.strip()]
    return lines or [text]


def clean_test_name(text: str) -> str:
    cleaned = re.sub(r"\[[^\]]*\]", "", text)
    cleaned = re.sub(r"['\"]", "", cleaned)
    cleaned = re.sub(r"\b\d{5}\b", "", cleaned)
    cleaned = re.sub(r"\bweekly\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bevery\s+\d+(?:\s*(?:to|-|/)\s*\d+)?\s+weeks?\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bbetween\s+\d+\s+to\s+\d+\s+weeks?\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bat diagnosis\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.strip(" -,.:")
    return cleaned or "Mapped test"


def parse_frequency(text: str, row_start: Optional[float], row_end: Optional[float]) -> Tuple[str, str, Optional[float], Optional[float]]:
    lowered = text.lower()

    between_match = re.search(r"between\s+(\d+(?:\.\d+)?)\s+to\s+(\d+(?:\.\d+)?)\s+weeks?", lowered)
    if between_match:
        start = float(between_match.group(1))
        end = float(between_match.group(2))
        return "once_in_window", "1", start, end

    if "weekly" in lowered:
        start = row_start
        end = 40.0 if row_end is not None and row_end >= row_start else max(row_end or 0.0, 40.0)
        return "weekly", "1", start, end

    every_match = re.search(r"every\s+(\d+)(?:\s*(?:to|-|/)\s*(\d+))?\s+weeks?", lowered)
    if every_match:
        first = every_match.group(1)
        second = every_match.group(2)
        freq_type = "every_n_to_m_weeks" if second else "every_n_weeks"
        freq_value = f"{first}-{second}" if second else first
        start = row_start
        end = 40.0 if row_end is None or row_end < row_start else row_end
        if end == row_start:
            end = 40.0
        return freq_type, freq_value, start, end

    if "at diagnosis" in lowered:
        start = row_start
        end = 40.0
        return "once", "1", start, end

    if row_start is not None and row_end is not None and row_end > row_start:
        return "once_in_window", "1", row_start, row_end

    return "once", "1", row_start, row_end


def parse_rule_item(program: str, applies_to_groups: str, row_week: Any, item_text: str) -> Rule:
    row_start, row_end = parse_row_week(row_week)
    freq_type, freq_value, start_week, end_week = parse_frequency(item_text, row_start, row_end)
    cpt_codes = sorted(set(CPT_RE.findall(item_text)))

    notes: List[str] = []
    lowered = item_text.lower()
    if "fgr only" in lowered:
        notes.append("FGR only")
    if "mfm" in lowered:
        notes.append("MFM")
    if "gap if not done within" in lowered:
        notes.append("Gap logic came from mapping note")
    if "should trigger when the diagnosis is done" in lowered:
        notes.append("Trigger when diagnosis is present")

    test_name = clean_test_name(item_text)
    if "fgr only" in lowered and "Umbilical artery Doppler" not in test_name:
        test_name = f"{test_name} (FGR only)"

    return Rule(
        program=program,
        applies_to_groups=applies_to_groups,
        test_name=test_name,
        cpt_codes="|".join(cpt_codes),
        start_week=start_week,
        end_week=end_week,
        frequency_type=freq_type,
        frequency_value=freq_value,
        stop_condition="",
        special_notes="; ".join(notes),
    )


def load_mapping_rules(raw_excel: bytes) -> pd.DataFrame:
    mapping = pd.read_excel(BytesIO(raw_excel))
    rules: List[Rule] = []

    week_col = None
    for candidate in mapping.columns:
        if str(candidate).strip().lower() == "gestational age":
            week_col = candidate
            break
    if week_col is None:
        raise KeyError(f"Could not find 'gestational age' column in mapping workbook. Found: {list(mapping.columns)}")

    for _, row in mapping.iterrows():
        row_week = row.get(week_col)
        for column_name, (program, applies_to_groups) in COLUMN_MAP.items():
            if column_name not in mapping.columns:
                continue
            cell_value = row.get(column_name)
            for item in split_mapping_cell(cell_value):
                rule = parse_rule_item(program, applies_to_groups, row_week, item)
                if not rule.cpt_codes:
                    continue
                rules.append(rule)

    if not rules:
        raise ValueError("No rules could be parsed from the mapping workbook.")

    rules_df = pd.DataFrame([r.__dict__ for r in rules])

    # Light cleanup for obviously duplicated rows caused by workbook formatting.
    rules_df = rules_df.drop_duplicates(
        subset=[
            "program",
            "applies_to_groups",
            "test_name",
            "cpt_codes",
            "start_week",
            "end_week",
            "frequency_type",
            "frequency_value",
        ]
    ).reset_index(drop=True)

    return rules_df


def has_preexisting_diabetes(row: pd.Series) -> bool:
    blob = " ".join(
        [
            str(row.get("category", "") or "").lower(),
            str(row.get("category_reason", "") or "").lower(),
            str(row.get("pregnancy_complications", "") or "").lower(),
        ]
    )
    return any(key in blob for key in ["pre-existing", "pre existing", "preexisting", "type 1", "t1dm"])


def is_diabetes_patient(row: pd.Series) -> bool:
    blob = " ".join(
        [
            str(row.get("category", "") or "").lower(),
            str(row.get("category_reason", "") or "").lower(),
            str(row.get("pregnancy_complications", "") or "").lower(),
        ]
    )
    return "diabetes" in blob


def is_obesity_patient(row: pd.Series) -> bool:
    bmi = as_float(row.get("BMI_at_pregnancy_start_num"))
    if bmi is not None and bmi >= 30:
        return True
    blob = " ".join(
        [
            str(row.get("category", "") or "").lower(),
            str(row.get("category_reason", "") or "").lower(),
            str(row.get("pregnancy_complications", "") or "").lower(),
        ]
    )
    return "obesity" in blob


def obesity_class_from_bmi(bmi: Optional[float]) -> str:
    if bmi is None:
        return ""
    if 30 <= bmi < 35:
        return "OBESITY|OBESITY_CLASS_I"
    if 35 <= bmi < 40:
        return "OBESITY|OBESITY_CLASS_II"
    if bmi >= 40:
        return "OBESITY|OBESITY_CLASS_III"
    return ""


def diabetes_group_from_step4(row: pd.Series) -> str:
    blob = " ".join(
        [
            str(row.get("category", "") or "").lower(),
            str(row.get("category_reason", "") or "").lower(),
            str(row.get("pregnancy_complications", "") or "").lower(),
        ]
    )
    if "diet" in blob:
        return "diet_controlled"
    if any(key in blob for key in ["insulin", "oral", "medication", "metformin", "glyburide"]):
        return "oral_or_insulin"
    return "pre_existing"


def rule_applies_to_patient(rule: pd.Series, patient: pd.Series, program: str) -> bool:
    if program == "DIABETES":
        return str(rule.get("applies_to_groups", "")).strip().lower() == diabetes_group_from_step4(patient)

    if program == "OBESITY":
        bmi = as_float(patient.get("BMI_at_pregnancy_start_num"))
        bmi_group = obesity_class_from_bmi(bmi)
        return str(rule.get("applies_to_groups", "")).strip().upper() == bmi_group

    return False


def compute_window(preg_start: pd.Timestamp, start_week: Optional[float], end_week: Optional[float]) -> Tuple[pd.Timestamp, pd.Timestamp]:
    start_dt = preg_start if start_week is None else preg_start + timedelta(weeks=float(start_week))
    end_dt = preg_start + timedelta(weeks=40 if end_week is None else float(end_week))
    return start_dt, end_dt
     

def gestational_week_at_run(preg_start: pd.Timestamp, run_ts: pd.Timestamp) -> float:
    return max(0.0, (run_ts - preg_start).days / 7.0)
    

def expected_count_by_run_date(
    run_ts: pd.Timestamp,
    win_start: pd.Timestamp,
    win_end: pd.Timestamp,
    frequency_type: str,
    frequency_value: str,
) -> int:
    if pd.isna(run_ts) or pd.isna(win_start) or pd.isna(win_end):
        return 0
    if run_ts < win_start:
        return 0

    effective_end = min(run_ts, win_end)
    if effective_end < win_start:
        return 0

    if frequency_type in {"once", "once_in_window"}:
        return 1

    elapsed_days = max(0, (effective_end - win_start).days)
    elapsed_weeks = int(elapsed_days // 7)

    if frequency_type == "weekly":
        return elapsed_weeks + 1

    if frequency_type == "every_n_weeks":
        interval = max(1, int(as_float(frequency_value) or 1))
        return (elapsed_weeks // interval) + 1

    if frequency_type == "every_n_to_m_weeks":
        bounds = parse_range(frequency_value)
        interval = bounds[1] if bounds else 1
        return (elapsed_weeks // interval) + 1

    return 0


def collect_matching_hits(
    enc_p: pd.DataFrame,
    cpt_codes: List[str],
    win_start: pd.Timestamp,
    win_end: pd.Timestamp,
    run_ts: pd.Timestamp,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if enc_p.empty or not cpt_codes:
        empty = pd.DataFrame(columns=["enc_cpt", "enc_dos"])
        return empty, empty

    mask = (
        enc_p["enc_cpt"].isin(cpt_codes)
        & (enc_p["enc_dos"] >= win_start)
        & (enc_p["enc_dos"] <= win_end)
    )
    all_hits = enc_p.loc[mask, ["enc_cpt", "enc_dos"]].dropna().sort_values(["enc_dos", "enc_cpt"]).copy()
    past_hits = all_hits[all_hits["enc_dos"] <= run_ts].copy()
    return all_hits, past_hits


def count_actual_occurrences(
    hits: pd.DataFrame,
    win_start: pd.Timestamp,
    run_ts: pd.Timestamp,
    frequency_type: str,
    frequency_value: str,
) -> int:
    if hits.empty:
        return 0

    if frequency_type in {"once", "once_in_window"}:
        return 1

    working = hits.copy()
    working["days_from_start"] = (working["enc_dos"] - win_start).dt.days.clip(lower=0)

    if frequency_type == "weekly":
        working["bucket"] = working["days_from_start"] // 7
        return int(working["bucket"].nunique())

    if frequency_type == "every_n_weeks":
        interval = max(1, int(as_float(frequency_value) or 1))
        working["bucket"] = working["days_from_start"] // (interval * 7)
        return int(working["bucket"].nunique())

    if frequency_type == "every_n_to_m_weeks":
        bounds = parse_range(frequency_value)
        interval = bounds[1] if bounds else 1
        working["bucket"] = working["days_from_start"] // (interval * 7)
        return int(working["bucket"].nunique())

    return 0


def summarize_gap(expected_count: int, actual_count: int, run_ts: pd.Timestamp, win_start: pd.Timestamp) -> Tuple[str, str]:
    if run_ts < win_start:
        return "No", "not due yet"
    if actual_count >= expected_count:
        return "No", "met by encounter data"
    return "Yes", f"expected {expected_count}, found {actual_count}"


def build_tracking_details(hits: pd.DataFrame) -> str:
    if hits.empty:
        return ""
    return " | ".join(f"{row.enc_cpt}:{fmt_mdy(row.enc_dos)}" for row in hits.itertuples(index=False))

# Added as per discussion 27-03-2026
def special_gap_details(cpt_codes: Optional[Union[str, List[str]]]) -> str:
    if not cpt_codes:
        return ""

    # Normalize to set of strings (robust + fast lookup)
    if isinstance(cpt_codes, str):
        cpt_set = {cpt_codes}
    else:
        cpt_set = set(map(str, cpt_codes))

    if "76811" in cpt_set:
        return "check MFM records"

    if cpt_set & {"76825", "76827"}:
        return "check external records"

    return ""


def apply_rules_for_program(
    program: str,
    step4_df: pd.DataFrame,
    rules_df: pd.DataFrame,
    encounters_df: pd.DataFrame,
    run_ts: pd.Timestamp,
) -> pd.DataFrame:
    enc_by_pid = build_enc_index(encounters_df)
    rows: List[Dict[str, Any]] = []

    if program == "DIABETES":
        cohort = step4_df[step4_df.apply(is_diabetes_patient, axis=1)].copy()
    elif program == "OBESITY":
        cohort = step4_df[step4_df.apply(is_obesity_patient, axis=1)].copy()
    else:
        cohort = pd.DataFrame(columns=step4_df.columns)

    rules_for_program = rules_df[rules_df["program"] == program].copy()

    for _, patient in cohort.iterrows():
        patient_id = normalize_patient_id(patient.get("patient_id", ""))
        preg_start = to_dt(patient.get("start_of_pregnancy_date"))
        if not patient_id or pd.isna(preg_start):
            continue

        enc_p = enc_by_pid.get(patient_id, pd.DataFrame(columns=["enc_patient_id", "enc_cpt", "enc_dos"]))
        bmi = as_float(patient.get("BMI_at_pregnancy_start_num"))
        fgr_flag = str(patient.get("has_fgr", "") or "").strip().upper()
        
        for _, rule in rules_for_program.iterrows():
            if not rule_applies_to_patient(rule, patient, program):
                continue

            if "FGR only" in str(rule.get("special_notes", "")) and fgr_flag != "Y":
                continue

            cpt_codes = [code for code in str(rule.get("cpt_codes", "")).split("|") if code]
            if not cpt_codes:
                continue

            win_start, win_end = compute_window(preg_start, rule.get("start_week"), rule.get("end_week"))
            all_hits, past_hits = collect_matching_hits(enc_p, cpt_codes, win_start, win_end, run_ts)

            expected_count = expected_count_by_run_date(
                run_ts=run_ts,
                win_start=win_start,
                win_end=win_end,
                frequency_type=str(rule.get("frequency_type", "")),
                frequency_value=str(rule.get("frequency_value", "")),
            )
            actual_tracking = min(
                expected_count,
                count_actual_occurrences(
                    hits=past_hits,
                    win_start=win_start,
                    run_ts=run_ts,
                    frequency_type=str(rule.get("frequency_type", "")),
                    frequency_value=str(rule.get("frequency_value", "")),
                ),
            )
            tracking_details = build_tracking_details(past_hits)
            gap_flag, gap_details = summarize_gap(expected_count, actual_tracking, run_ts, win_start)

            preg_comp = str(patient.get("pregnancy_complications", "") or "")
            bmi_group = obesity_class_from_bmi(bmi)
            if "Obesity Complication" in preg_comp and bmi_group and bmi_group not in preg_comp:
                preg_comp = f"{preg_comp}, {bmi_group}"

            # Added as per discussion dated 27-03-2026
            spl_gap_dtls = special_gap_details(cpt_codes)
            if spl_gap_dtls:
                gap_details =  spl_gap_dtls
                gap_flag =""

    
            rows.append(
                {
                    "process_run_date": fmt_mdy(run_ts),
                    "patient_id": patient_id,
                    "first_name": patient.get("first_name", ""),
                    "last_name": patient.get("last_name", ""),
                    "date_of_birth": fmt_mdy(to_dt(patient.get("date_of_birth"))),
                    "payer": patient.get("payer", ""),
                    "start_of_pregnancy_date": fmt_mdy(preg_start),
                    # Added current_gestational_age as per discussion dated 28-03-2026
                    "current_gestational_age": patient.get("current_gestational_age",""),
                    "pregnancy_complications": preg_comp,
                    "BMI_at_pregnancy_start": bmi if bmi is not None else "",
                    "has_fgr": fgr_flag,
                    "test_name": rule.get("test_name", ""),
                    "cpt_codes": rule.get("cpt_codes", ""),
                    #"start_week": rule.get("start_week", ""),
                    #"end_week": rule.get("end_week", ""),
                    "frequency_type": rule.get("frequency_type", ""),
                    "frequency_value": rule.get("frequency_value", ""),
                    #"stop_condition": rule.get("stop_condition", ""),
                    #"special_notes": rule.get("special_notes", ""),
                    "window_start": fmt_mdy(win_start),
                    "window_end": fmt_mdy(win_end),
                    #"frequency_hits": len(past_hits),
                    #"expected_count": expected_count,
                    #"actual_tracking": actual_tracking,
                    "actual_tracking_details": tracking_details,
                    #"is_there_a_gap": gap_flag,
                    "gap_details": gap_details,
                    "expected_test_count" : expected_count,
                    "actual_test_count": actual_tracking,
                    "gap": gap_flag,
                }
            )
            
    #added 13-04-2026        
    # 1. Sort by gestational age DESCENDING first
    rows.sort(key=lambda r: r.get("current_gestational_age", ""), reverse=True)

    # 2. Sort by patient_id ASCENDING (stable sort keeps age order for same ID)
    #rows.sort(key=lambda r: r.get("patient_id", ""))

    return pd.DataFrame(rows)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create combined eviCore join output using the mapping workbook from the input folder."
    )
    parser.add_argument("--account-name", required=True)
    parser.add_argument("--encounters-container", default="input")
    parser.add_argument("--encounters-blob", default="EncounterData.xlsx")
    parser.add_argument("--mapping-container", default="input")
    parser.add_argument("--mapping-blob", default="MultipleConditions_Evicore_Mapping.xls")
    parser.add_argument("--step4-container", default="output")
    parser.add_argument("--step4-prefix", default="")
    parser.add_argument("--output-container", default="output")
    parser.add_argument("--output-prefix", default="")
    parser.add_argument("--process-run-date", default="", help="Optional YYYY-MM-DD. Default is today in America/Chicago.")
    parser.add_argument("--local-out", default="", help="Optional local output path.")
    args = parser.parse_args()

    if args.process_run_date.strip():
        run_date = datetime.strptime(args.process_run_date.strip(), "%Y-%m-%d").date()
    else:
        if ZoneInfo is not None:
            run_date = datetime.now(ZoneInfo("America/Chicago")).date()
        else:
            run_date = datetime.now().date()

    run_ts = pd.Timestamp(datetime(run_date.year, run_date.month, run_date.day, 23, 59, 59))

    account_url = f"https://{args.account_name}.blob.core.windows.net"
    bsc = blob_service_client(account_url)

    step4_blob = pick_latest_step4(bsc, args.step4_container, args.step4_prefix)
    log(f"[JOIN] Latest Step4: container={args.step4_container} blob={step4_blob}")

    step4_raw = download_bytes(bsc, BlobLoc(args.step4_container, step4_blob))
    encounters_raw = download_bytes(bsc, BlobLoc(args.encounters_container, args.encounters_blob))
    mapping_raw = download_bytes(bsc, BlobLoc(args.mapping_container, args.mapping_blob))

    step4_df = load_step4_df(step4_raw)
    encounters_df = load_encounters_df(encounters_raw)
    rules_df = load_mapping_rules(mapping_raw)
     
    diabetes_out = apply_rules_for_program("DIABETES", step4_df, rules_df, encounters_df, run_ts)
    obesity_out = apply_rules_for_program("OBESITY", step4_df, rules_df, encounters_df, run_ts)

    out_df = pd.concat([diabetes_out, obesity_out], ignore_index=True)
    if out_df.empty:
        log("[JOIN] No output rows were produced.")
        return

#Columns are commented and "current_gestational_age" is added as per discussion dated 27-03-2026
    ordered_cols = [
        "process_run_date",
        "patient_id",
        "first_name",
        "last_name",
        "date_of_birth",
        "payer",
        "start_of_pregnancy_date",
        "current_gestational_age",
        "pregnancy_complications",
        "BMI_at_pregnancy_start",
        "has_fgr",
        "test_name",
        "cpt_codes",
        #"start_week",
        #"end_week",
        "frequency_type",
        "frequency_value",
        #"stop_condition",
        #"special_notes",
        "window_start",
        "window_end",
        #"frequency_hits",
        #"expected_count",
        #"actual_tracking",
        "actual_tracking_details",
        #"is_there_a_gap",
        "gap_details",
        "expected_test_count" ,
        "actual_test_count",
        "gap",
    ]
    for col in ordered_cols:
        if col not in out_df.columns:
            out_df[col] = ""
    extras = [c for c in out_df.columns if c not in ordered_cols]
    out_df = out_df[ordered_cols + extras]

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    #local_path = args.local_out.strip() or f"evicore_join_combined_{timestamp}.csv"
    #added 13-04-2026
    local_path = args.local_out.strip() or f"HighRiskPregnancy_tracking_evicore_{timestamp}.csv"
    out_df.to_csv(local_path, index=False)
    log(f"[JOIN] Wrote local: {local_path} ({len(out_df)} rows)")

    output_prefix = (args.output_prefix or "").lstrip("/")
    if output_prefix and not output_prefix.endswith("/"):
        output_prefix += "/"
    #out_blob = f"{output_prefix}evicore_join_combined_{timestamp}.csv"
    out_blob = f"{output_prefix}HighRiskPregnancy_tracking_evicore_{timestamp}.csv"

    with open(local_path, "rb") as fh:
        upload_bytes(bsc, BlobLoc(args.output_container, out_blob), fh.read())
    log(f"[JOIN] Uploaded: container={args.output_container} blob={out_blob}")


if __name__ == "__main__":
    main()
