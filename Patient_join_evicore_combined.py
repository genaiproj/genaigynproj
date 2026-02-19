#!/usr/bin/env python3
# Patient_join_evicore_combined.py
# Fix 2026-01-31: normalize timestamps so we do not compare tz-aware with tz-naive values.
# Root cause: Encounter DOS comes in tz-naive from Excel, but we were creating a tz-aware run_ts.
#
# Discussion changes applied (2026-01-30):
# 1) If a future-dated CPT exists after the process run date, "Is there a gap?" -> "Check for next appointment".
# 2) For CPT 76811 or 76827, "Is there a gap?" -> "Not in WHF encounters, check MFM chart notes".
#    Optional rules column: spl_notes (example values: MFM_ONLY, NOT_IN_WHF).
# 4) If patient is pre-existing diabetic and obese, keep diabetes only (obesity suppressed).
# 5) Frequency ranges enforce the lower boundary (2-4 weeks means due by 2 weeks).
# 6) Weekly cycles are checked bucket-by-bucket up to the process run date.

import argparse
import os
import re
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient

try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None  # fallback

GRACE_WEEKS = 3
CPT_MFM_ONLY = {"76811", "76827"}
_CPT_RE = re.compile(r"\b(\d{5})\b")
_STEP4_TS_RE = re.compile(r"step4_(\d{8})_(\d{6})", re.IGNORECASE)


def log(msg: str) -> None:
    print(msg)


@dataclass(frozen=True)
class BlobLoc:
    container: str
    blob: str


def _normalize_patient_id(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def _normalize_cpt(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip()
    m = _CPT_RE.search(s)
    return m.group(1) if m else s


def _to_dt(x) -> pd.Timestamp:
    return pd.to_datetime(x, errors="coerce")


def _strip_tz(series: pd.Series) -> pd.Series:
    # Keep everything tz-naive for consistent comparisons.
    s = pd.to_datetime(series, errors="coerce")
    try:
        if getattr(s.dt, "tz", None) is not None:
            # tz-aware -> naive
            return s.dt.tz_convert(None)
    except Exception:
        pass
    return s


def _fmt_mdy(ts: pd.Timestamp) -> str:
    if pd.isna(ts):
        return ""
    try:
        return ts.strftime("%-m/%-d/%Y")
    except Exception:
        return ts.strftime("%m/%d/%Y")


def _as_float_or_none(x: Any) -> Optional[float]:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return None
        s = str(x).strip()
        if not s or s.lower() in ("nan", "none", "null"):
            return None
        return float(s)
    except Exception:
        return None


def _parse_week(v) -> Optional[float]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return None
    m = re.search(r"(-?\d+(?:\.\d+)?)", s)
    return float(m.group(1)) if m else None


def _parse_rule_cpts(raw) -> List[str]:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return []
    codes = _CPT_RE.findall(str(raw))
    return sorted(set(codes)) if codes else []


def _compute_window(
    preg_start: pd.Timestamp,
    start_week: Optional[float],
    end_week: Optional[float],
) -> Tuple[pd.Timestamp, pd.Timestamp]:
    # 2026-01-31: treat NaN like missing. Some CSV exports keep empty numeric cells as NaN.
    if start_week is None or (isinstance(start_week, float) and pd.isna(start_week)):
        win_start = preg_start
    else:
        win_start = preg_start + timedelta(weeks=float(start_week))

    if end_week is None or (isinstance(end_week, float) and pd.isna(end_week)):
        win_end = preg_start + timedelta(weeks=40)
    else:
        win_end = preg_start + timedelta(weeks=float(end_week))

    return win_start, win_end



def _step4_dt_from_name(name: str) -> Optional[datetime]:
    m = _STEP4_TS_RE.search(name)
    if not m:
        return None
    ymd, hms = m.group(1), m.group(2)
    try:
        return datetime.strptime(f"{ymd}_{hms}", "%Y%m%d_%H%M%S").replace(tzinfo=timezone.utc)
    except Exception:
        return None


def _blob_service_client(account_url: str) -> BlobServiceClient:
    conn = os.getenv("AZURE_STORAGE_CONNECTION_STRING", "").strip()
    if conn:
        return BlobServiceClient.from_connection_string(conn)
    cred = DefaultAzureCredential(exclude_interactive_browser_credential=True)
    return BlobServiceClient(account_url=account_url, credential=cred)


def _download_bytes(bsc: BlobServiceClient, loc: BlobLoc) -> bytes:
    bc = bsc.get_blob_client(container=loc.container, blob=loc.blob)
    return bc.download_blob().readall()


def _upload_bytes(bsc: BlobServiceClient, loc: BlobLoc, data: bytes) -> None:
    bc = bsc.get_blob_client(container=loc.container, blob=loc.blob)
    bc.upload_blob(data, overwrite=True)


def pick_latest_step4(bsc: BlobServiceClient, container: str, prefix: str) -> str:
    prefix = (prefix or "").lstrip("/")
    if prefix and not prefix.endswith("/"):
        prefix += "/"

    cc = bsc.get_container_client(container)
    candidates: List[Tuple[Optional[datetime], Optional[datetime], str]] = []

    for b in cc.list_blobs(name_starts_with=prefix + "step4_"):
        name = b.name
        low = name.lower()
        if "care_tracking_output" not in low:
            continue
        if not (low.endswith(".csv") or low.endswith(".csv.csv")):
            continue

        ts = _step4_dt_from_name(name)
        lm = b.last_modified
        if lm and lm.tzinfo is None:
            lm = lm.replace(tzinfo=timezone.utc)

        candidates.append((ts, lm, name))

    if not candidates:
        raise FileNotFoundError(
            f"No Step4 Care Tracking output found in container '{container}' under prefix '{prefix or '(root)'}'."
        )

    candidates.sort(key=lambda x: (x[0] or datetime.min.replace(tzinfo=timezone.utc), x[1] or datetime.min.replace(tzinfo=timezone.utc)))
    return candidates[-1][2]


def load_step4_df(raw_csv: bytes) -> pd.DataFrame:
    df = pd.read_csv(BytesIO(raw_csv))

    if "patient_id" not in df.columns:
        for cand in ["patientid", "PatientID", "PATIENT_ID", "MRN"]:
            if cand in df.columns:
                df = df.rename(columns={cand: "patient_id"})
                break
    if "patient_id" not in df.columns:
        raise KeyError(f"Step4 file is missing patient_id. Columns found: {list(df.columns)}")

    df["patient_id"] = df["patient_id"].apply(_normalize_patient_id)

    if "start_of_pregnancy_date" not in df.columns:
        raise KeyError("Step4 file is missing start_of_pregnancy_date.")
    df["start_of_pregnancy_date"] = _strip_tz(df["start_of_pregnancy_date"])


    if "BMI_at_pregnancy_start_num" not in df.columns:
        if "BMI_at_pregnancy_start" in df.columns:
            df["BMI_at_pregnancy_start_num"] = pd.to_numeric(df["BMI_at_pregnancy_start"], errors="coerce")
        else:
            df["BMI_at_pregnancy_start_num"] = pd.NA

    for col in ["category", "category_reason", "pregnancy_complications"]:
        if col not in df.columns:
            df[col] = ""

    return df


def load_rules_df(raw_csv: bytes, rules_name: str) -> pd.DataFrame:
    df = pd.read_csv(BytesIO(raw_csv))

    required = ["condition_id", "test_name", "cpt_codes", "start_week", "end_week", "frequency_type", "frequency_value"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"{rules_name}: rules file missing required columns: {missing}")

    df = df.copy()
    df["condition_id"] = df["condition_id"].fillna("").astype(str).str.strip()
    df["test_name"] = df["test_name"].fillna("").astype(str).str.strip()

    df["start_week"] = df["start_week"].apply(_parse_week)
    df["end_week"] = df["end_week"].apply(_parse_week)

    df["frequency_type"] = df["frequency_type"].fillna("").astype(str).str.strip().str.lower()
    df["frequency_value"] = df["frequency_value"].replace({pd.NA: "", None: ""}).astype(str).str.strip()

    for col in ["bmi_min", "bmi_max"]:
        if col in df.columns:
            df[col] = df[col].apply(_as_float_or_none)
        else:
            df[col] = None

    for col in ["stop_condition", "special_notes", "guideline_section", "page_ref"]:
        if col not in df.columns:
            df[col] = ""
        else:
            df[col] = df[col].fillna("").astype(str).str.strip()

    if "spl_notes" not in df.columns:
        df["spl_notes"] = ""
    else:
        df["spl_notes"] = df["spl_notes"].fillna("").astype(str).str.strip()

    return df.reset_index(drop=True)


def load_encounters_df(raw_xlsx: bytes) -> pd.DataFrame:
    enc = pd.read_excel(BytesIO(raw_xlsx))

    def find_col(cands):
        lower = {c.lower(): c for c in enc.columns}
        for c in cands:
            if c in enc.columns:
                return c
            if c.lower() in lower:
                return lower[c.lower()]
        raise KeyError(f"Missing column. Tried {cands}. Found {list(enc.columns)}")

    pid_col = find_col(["patientid", "patient_id", "PatientID", "PATIENT_ID", "MRN"])
    cpt_col = find_col(["CPTCode", "CPT_CODE", "CPT", "cpt"])
    dos_col = find_col(["DOS", "Date of Service", "Service Date"])

    enc = enc.rename(columns={pid_col: "enc_patient_id", cpt_col: "enc_cpt", dos_col: "enc_dos"})
    enc["enc_patient_id"] = enc["enc_patient_id"].apply(_normalize_patient_id)
    enc["enc_cpt"] = enc["enc_cpt"].apply(_normalize_cpt)
    enc["enc_dos"] = _strip_tz(enc["enc_dos"])

    return enc[["enc_patient_id", "enc_cpt", "enc_dos"]].copy()


def build_enc_index(enc: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    out: Dict[str, pd.DataFrame] = {}
    enc = enc.dropna(subset=["enc_patient_id"]).copy()
    enc["enc_patient_id"] = enc["enc_patient_id"].astype(str)
    for pid, g in enc.groupby("enc_patient_id"):
        out[pid] = g.sort_values("enc_dos").reset_index(drop=True)
    return out


def _parse_range_weeks(freq_value: str) -> Optional[Tuple[int, int]]:
    if not freq_value:
        return None
    s = str(freq_value).strip().lower()
    nums = [int(x) for x in re.findall(r"(\d+)", s)]
    if len(nums) >= 2:
        return (min(nums[0], nums[1]), max(nums[0], nums[1]))
    if len(nums) == 1:
        return (nums[0], nums[0])
    return None



MONTH_ABBR_TO_NUM = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}


def _parse_excel_range_artifact(s: str) -> Optional[Tuple[int, int]]:
    """Handle Excel artifacts like '6-Mar' that sometimes come from '3-6'."""
    if not s or not isinstance(s, str):
        return None
    t = s.strip().lower()
    m = re.match(r"^(\d{1,2})\s*[-/]\s*([a-z]{3,4})$", t)
    if not m:
        return None
    a = int(m.group(1))
    mon = MONTH_ABBR_TO_NUM.get(m.group(2)[:4], None)
    if not mon:
        return None
    lo, hi = sorted([a, mon])
    return (lo, hi)


def normalize_frequency(freq_type: str, freq_value: str) -> Tuple[str, str]:
    """Map messy rule inputs into: once, weekly, every_n_weeks, every_n_to_m_weeks."""
    ft = (freq_type or "").strip().lower()
    fv_raw = "" if freq_value is None else str(freq_value).strip()
    fv = fv_raw.lower()

    if not ft and fv:
        if "week" in fv or re.search(r"\b\d+\b", fv):
            ft = "every"

    if not ft or ft == "once":
        return "once", fv_raw

    if ft in ("every", "q", "interval", "every_week", "everyweeks"):
        if "to" in fv or "-" in fv or "/" in fv:
            rng = _parse_range_weeks(fv_raw) or _parse_excel_range_artifact(fv)
            if rng:
                lo, hi = rng
                return "every_n_to_m_weeks", f"{lo}-{hi}"
        nums = re.findall(r"(\d+)", fv)
        if nums:
            return "every_n_weeks", nums[0]
        return "every_n_weeks", fv_raw

    if ft in ("other", "misc", "custom"):
        if "weekly" in fv:
            if "twice weekly" in fv or "2x" in fv or "two times a week" in fv:
                return "weekly", "2"
            # "up to twice" -> allowed max, requirement min stays 1
            return "weekly", "1"
        if "week" in fv:
            return "every", fv_raw
        return "once", fv_raw

    if ft == "weekly":
        if "up to twice" in fv or "up_to_twice" in fv or "upto" in fv:
            return "weekly", "1"
        if "twice weekly" in fv or "2x" in fv or "two times a week" in fv:
            return "weekly", "2"
        nums = re.findall(r"(\d+)", fv)
        return ("weekly", nums[0]) if nums else ("weekly", "1")

    if ft == "every_n_to_m_weeks":
        rng = _parse_range_weeks(fv_raw) or _parse_excel_range_artifact(fv)
        if rng:
            lo, hi = rng
            return "every_n_to_m_weeks", f"{lo}-{hi}"
        return "every_n_to_m_weeks", fv_raw

    if ft == "every_n_weeks":
        nums = re.findall(r"(\d+)", fv)
        return ("every_n_weeks", nums[0]) if nums else ("every_n_weeks", fv_raw)

    return ft, fv_raw

def _effective_end(win_end: pd.Timestamp, run_ts: pd.Timestamp) -> pd.Timestamp:
    return min(win_end, run_ts)


def _bucketize_weeks(d: pd.Series, win_start: pd.Timestamp) -> pd.Series:
    return ((d - win_start).dt.days // 7).clip(lower=0)


def _gap_status_once(past_hits: pd.DataFrame, future_hits: pd.DataFrame, win_end: pd.Timestamp, run_ts: pd.Timestamp) -> Tuple[str, str]:
    if not past_hits.empty:
        return "No", f"{len(past_hits)} hit(s)"
    if not future_hits.empty:
        first = future_hits.sort_values("enc_dos").iloc[0]["enc_dos"]
        return "Check for next appointment", f"future date {_fmt_mdy(first)}"
    if run_ts >= win_end:
        return "Yes", "no evidence by run date"
    return "No", "not due yet"


def _gap_status_weekly(
    past_hits: pd.DataFrame,
    future_hits: pd.DataFrame,
    win_start: pd.Timestamp,
    win_end: pd.Timestamp,
    run_ts: pd.Timestamp,
    need_per_week: int,
) -> Tuple[str, str]:
    if run_ts < win_start:
        return "No", "not due yet"

    eff_end = _effective_end(win_end, run_ts)
    total_days = max(1, int((eff_end - win_start).days) + 1)
    total_weeks = max(1, int((total_days + 6) // 7))

    if past_hits.empty:
        if not future_hits.empty:
            first = future_hits.sort_values("enc_dos").iloc[0]["enc_dos"]
            return "Check for next appointment", f"future date {_fmt_mdy(first)}"
        return "Yes", "0 hits"

    idx = _bucketize_weeks(past_hits["enc_dos"], win_start)
    counts = idx.value_counts().to_dict()

    bad = []
    for w in range(total_weeks):
        c = int(counts.get(w, 0))
        if c < need_per_week:
            bad.append(f"wk{w+1}:{c}/{need_per_week}")

    if bad:
        if not future_hits.empty:
            first = future_hits.sort_values("enc_dos").iloc[0]["enc_dos"]
            return "Check for next appointment", f"missing {bad[0]} (future {_fmt_mdy(first)})"
        return "Yes", " | ".join(bad[:12]) + (" | ..." if len(bad) > 12 else "")

    return "No", f"weekly ok ({need_per_week}/week)"


def _gap_status_every_n_weeks(
    past_hits: pd.DataFrame,
    future_hits: pd.DataFrame,
    win_start: pd.Timestamp,
    win_end: pd.Timestamp,
    run_ts: pd.Timestamp,
    block_weeks: int,
) -> Tuple[str, str]:
    if run_ts < win_start:
        return "No", "not due yet"

    eff_end = _effective_end(win_end, run_ts)
    total_days = max(1, int((eff_end - win_start).days) + 1)
    total_weeks = max(1, int((total_days + 6) // 7))
    n_blocks = max(1, int((total_weeks + block_weeks - 1) // block_weeks))

    if past_hits.empty:
        if not future_hits.empty:
            first = future_hits.sort_values("enc_dos").iloc[0]["enc_dos"]
            return "Check for next appointment", f"future date {_fmt_mdy(first)}"
        return "Yes", "0 hits"

    idx = _bucketize_weeks(past_hits["enc_dos"], win_start)
    block_idx = (idx // block_weeks)
    counts = block_idx.value_counts().to_dict()

    bad = []
    for b in range(n_blocks):
        c = int(counts.get(b, 0))
        if c < 1:
            bad.append(f"blk{b+1}:0/1")

    if bad:
        if not future_hits.empty:
            first = future_hits.sort_values("enc_dos").iloc[0]["enc_dos"]
            return "Check for next appointment", f"missing {bad[0]} (future {_fmt_mdy(first)})"
        return "Yes", " | ".join(bad[:12]) + (" | ..." if len(bad) > 12 else "")

    return "No", f"q{block_weeks}w ok (>=1/block)"


def _special_message_for_rule(cpt_list: List[str], spl_notes: str) -> str:
    if CPT_MFM_ONLY.intersection(set(cpt_list)):
        return "Not in WHF encounters, check MFM chart notes"
    s = (spl_notes or "").strip().lower()
    if not s:
        return ""
    if "mfm" in s or "not in whf" in s or "mfm_only" in s or "not_in_whf" in s:
        return "Not in WHF encounters, check MFM chart notes"
    return ""


def compute_gap_status(
    hits_all: pd.DataFrame,
    win_start: pd.Timestamp,
    win_end: pd.Timestamp,
    freq_type: str,
    freq_value: str,
    run_ts: pd.Timestamp,
    special_msg: str,
) -> Tuple[str, str]:
    if special_msg:
        return special_msg, "special handling"

    if hits_all is None or hits_all.empty:
        past_hits = pd.DataFrame(columns=["enc_dos"])
        future_hits = pd.DataFrame(columns=["enc_dos"])
    else:
        past_hits = hits_all[hits_all["enc_dos"] <= run_ts].copy()
        future_hits = hits_all[hits_all["enc_dos"] > run_ts].copy()

    ft, fv = normalize_frequency(freq_type, freq_value)

    if not ft or ft == "once":
        return _gap_status_once(past_hits, future_hits, win_end, run_ts)

    if ft == "weekly":
        need = int(_as_float_or_none(fv) or 1)
        return _gap_status_weekly(past_hits, future_hits, win_start, win_end, run_ts, need)

    if ft == "every_n_weeks":
        block = int(_as_float_or_none(fv) or 1)
        block = max(1, block)
        return _gap_status_every_n_weeks(past_hits, future_hits, win_start, win_end, run_ts, block)

    if ft == "every_n_to_m_weeks":
        rng = _parse_range_weeks(fv)
        block = 1 if not rng else max(1, int(rng[0]))
        return _gap_status_every_n_weeks(past_hits, future_hits, win_start, win_end, run_ts, block)

    if not past_hits.empty:
        return "No", f"{len(past_hits)} hit(s)"
    if not future_hits.empty:
        first = future_hits.sort_values("enc_dos").iloc[0]["enc_dos"]
        return "Check for next appointment", f"future date {_fmt_mdy(first)}"
    if run_ts >= win_end:
        return "Yes", "no evidence by run date"
    return "No", "not due yet"


def _has_preexisting_diabetes(row: pd.Series) -> bool:
    cat = str(row.get("category", "") or "").lower()
    comp = str(row.get("pregnancy_complications", "") or "").lower()

    if "diabetes|" in cat and ("pre" in cat or "pre-existing" in cat or "pre existing" in cat):
        return True
    if "pre-existing diabetes" in comp or "pre existing diabetes" in comp:
        return True
    return False


def _is_diabetes_patient(row: pd.Series) -> bool:
    cat = str(row.get("category", "") or "").lower()
    comp = str(row.get("pregnancy_complications", "") or "").lower()
    if "diabetes|" in cat:
        return True
    if "gestational diabetes" in comp or "diabetes" in comp:
        return True
    return False


def _is_obesity_patient(row: pd.Series) -> bool:
    if _has_preexisting_diabetes(row):
        return False
    cat = str(row.get("category", "") or "").lower()
    comp = str(row.get("pregnancy_complications", "") or "").lower()
    bmi = _as_float_or_none(row.get("BMI_at_pregnancy_start_num"))
    if bmi is not None and bmi >= 30:
        return True
    if "obesity|" in cat:
        return True
    if "obesity" in comp:
        return True
    return False


def _apply_rules_for_program(
    program: str,
    step4: pd.DataFrame,
    rules: pd.DataFrame,
    encounters: pd.DataFrame,
    rules_blob: str,
    step4_blob: str,
    run_ts: pd.Timestamp,
) -> pd.DataFrame:
    enc_by_pid = build_enc_index(encounters)
    results: List[Dict[str, Any]] = []

    if program == "DIABETES":
        cohort = step4[step4.apply(_is_diabetes_patient, axis=1)].copy()
    elif program == "OBESITY":
        cohort = step4[step4.apply(_is_obesity_patient, axis=1)].copy()
    else:
        raise ValueError(f"Unknown program: {program}")

    for _, p in cohort.iterrows():
        pid = _normalize_patient_id(p.get("patient_id", ""))
        preg_start = _to_dt(p.get("start_of_pregnancy_date"))   
        fgr_yn = p.get("has_fgr","")
        if not pid or pd.isna(preg_start):
            continue

        bmi = _as_float_or_none(p.get("BMI_at_pregnancy_start_num"))
        enc_p = enc_by_pid.get(pid, pd.DataFrame(columns=["enc_patient_id", "enc_cpt", "enc_dos"]))

        for _, r in rules.iterrows():
            bmi_min = _as_float_or_none(r.get("bmi_min"))
            bmi_max = _as_float_or_none(r.get("bmi_max"))
            if bmi_min is not None or bmi_max is not None:
                if bmi is None:
                    continue
                if bmi_min is not None and bmi < bmi_min:
                    continue
                if bmi_max is not None and bmi > bmi_max:
                    continue

            cpt_list = _parse_rule_cpts(r.get("cpt_codes"))
            if not cpt_list:
                continue

            win_start, win_end = _compute_window(preg_start, r.get("start_week"), r.get("end_week"))

            grace_start = win_start - timedelta(weeks=GRACE_WEEKS)
            grace_end = win_end + timedelta(weeks=GRACE_WEEKS)

            hits = pd.DataFrame(columns=["enc_cpt", "enc_dos"])
            if not enc_p.empty:
                mask = (
                    enc_p["enc_cpt"].isin(cpt_list)
                    & (enc_p["enc_dos"] >= grace_start)
                    & (enc_p["enc_dos"] <= grace_end)
                )
                hits = enc_p.loc[mask, ["enc_cpt", "enc_dos"]].dropna().copy()
                hits = hits.sort_values(["enc_dos", "enc_cpt"])
                hits["dos_key"] = hits["enc_dos"].dt.strftime("%Y-%m-%d")
                hits = hits.drop_duplicates(subset=["dos_key"]).drop(columns=["dos_key"])

            tracking = ""
            if not hits.empty:
                tracking = " | ".join(
                    f"{row.enc_cpt}:{_fmt_mdy(row.enc_dos)}"
                    for row in hits.itertuples(index=False)
                )

            freq_type = str(r.get("frequency_type", "")).strip().lower()
            freq_value = str(r.get("frequency_value", "")).strip()
            special_msg = _special_message_for_rule(cpt_list, str(r.get("spl_notes", "")).strip())
            status, details = compute_gap_status(
                hits_all=hits,
                win_start=win_start,
                win_end=win_end,
                freq_type=freq_type,
                freq_value=freq_value,
                run_ts=run_ts,
                special_msg=special_msg,
            )

            results.append({
               # "program": program,
                "process_run_date": _fmt_mdy(run_ts),
               # "rules_blob": rules_blob,
               # "step4_blob": step4_blob,
                
                "patient_id": pid,
                "first_name": p.get("first_name", ""),
                "last_name": p.get("last_name", ""),
                "date_of_birth": _fmt_mdy(_to_dt(p.get("date_of_birth"))),
                "start_of_pregnancy_date": _fmt_mdy(preg_start),
                "payer": p.get("payer", ""),
                #"pregnancy_complications": p.get("pregnancy_complications", ""),
                "pregnancy_complications": p.get("category_reason", ""),
                #"category": p.get("category", ""),
                #"category_reason": p.get("category_reason", ""),
                "BMI_at_pregnancy_start": bmi if bmi is not None else "",
                "has_fgr" : fgr_yn,

                #"condition_id": str(r.get("condition_id", "")).strip(),
                #"bmi_min": bmi_min if bmi_min is not None else "",
                #"bmi_max": bmi_max if bmi_max is not None else "",
                "test_name": str(r.get("test_name", "")).strip(),
                "cpt_codes": str(r.get("cpt_codes", "")).strip(),
                "start_week": r.get("start_week") if r.get("start_week") is not None else "",
                "end_week": r.get("end_week") if r.get("end_week") is not None else "",
                "frequency_type": freq_type,
                "frequency_value": freq_value,
                "stop_condition": str(r.get("stop_condition", "")).strip(),
                "special_notes": str(r.get("special_notes", "")).strip(),
                #"spl_notes": str(r.get("spl_notes", "")).strip(),
                #"guideline_section": str(r.get("guideline_section", "")).strip(),
                #"page_ref": str(r.get("page_ref", "")).strip(),

                "window_start": _fmt_mdy(win_start),
                "window_end": _fmt_mdy(win_end),
                "frequency_hits": int(len(hits)),
                "actual_tracking": tracking,
                #"is_there_a_gap": status,
                "is_there_a_gap": "No" if "76820" in str(r.get("cpt_codes", "")) and fgr_yn =="N" else status ,
                "gap_details": details,
                "gap":  "No" if "76820" in str(r.get("cpt_codes", "")) and fgr_yn =="N" else status,
            })

    return pd.DataFrame(results)


def main() -> None:
    ap = argparse.ArgumentParser(description="Run diabetes + obesity eviCore joins with one common output layout.")
    ap.add_argument("--account-name", required=True)

    ap.add_argument("--encounters-container", default="input")
    ap.add_argument("--encounters-blob", default="EncounterData.xlsx")

    ap.add_argument("--step4-container", default="output")
    ap.add_argument("--step4-prefix", default="")

    ap.add_argument("--diabetes-rules-container", default="input")
    ap.add_argument("--diabetes-rules-blob", default="diabetes_rules_norm.csv")

    ap.add_argument("--obesity-rules-container", default="input")
    ap.add_argument("--obesity-rules-blob", default="obesity_rules_norm.csv")

    ap.add_argument("--output-container", default="output")
    ap.add_argument("--output-prefix", default="")

    ap.add_argument("--process-run-date", default="", help="Optional YYYY-MM-DD. Default is today (America/Chicago).")
    ap.add_argument("--local-out", default="", help="Optional local output path.")

    args = ap.parse_args()

    if args.process_run_date.strip():
        try:
            run_dt = datetime.strptime(args.process_run_date.strip(), "%Y-%m-%d").date()
        except Exception:
            raise SystemExit("process-run-date must be YYYY-MM-DD")
    else:
        if ZoneInfo is not None:
            run_dt = datetime.now(ZoneInfo("America/Chicago")).date()
        else:
            run_dt = datetime.now().date()

    # Important: keep tz-naive so it matches encounter and step4 dates
    run_ts = pd.Timestamp(datetime(run_dt.year, run_dt.month, run_dt.day, 23, 59, 59))

    account_url = f"https://{args.account_name}.blob.core.windows.net"
    bsc = _blob_service_client(account_url)

    step4_blob = pick_latest_step4(bsc, args.step4_container, args.step4_prefix)
    log(f"[JOIN] Latest Step4: container={args.step4_container} blob={step4_blob}")

    step4_raw = _download_bytes(bsc, BlobLoc(args.step4_container, step4_blob))
    enc_raw = _download_bytes(bsc, BlobLoc(args.encounters_container, args.encounters_blob))

    dia_rules_raw = _download_bytes(bsc, BlobLoc(args.diabetes_rules_container, args.diabetes_rules_blob))
    obe_rules_raw = _download_bytes(bsc, BlobLoc(args.obesity_rules_container, args.obesity_rules_blob))

    step4_df = load_step4_df(step4_raw)
    enc_df = load_encounters_df(enc_raw)

    dia_rules_df = load_rules_df(dia_rules_raw, "diabetes")
    obe_rules_df = load_rules_df(obe_rules_raw, "obesity")

    dia_out = _apply_rules_for_program("DIABETES", step4_df, dia_rules_df, enc_df, args.diabetes_rules_blob, step4_blob, run_ts)
    obe_out = _apply_rules_for_program("OBESITY", step4_df, obe_rules_df, enc_df, args.obesity_rules_blob, step4_blob, run_ts)

    out_df = pd.concat([dia_out, obe_out], ignore_index=True)

    if out_df.empty:
        log("[JOIN] No output rows (no cohort patients found or no rules matched).")
        return

    out_cols = [
        # "program", 
        "process_run_date", 
        # "rules_blob",
        # "step4_blob",
        "patient_id", 
        "first_name", 
        "last_name", 
        "date_of_birth", 
        "payer",
        "start_of_pregnancy_date", 
        "pregnancy_complications", 
        # "category", 
        # "category_reason",
        "BMI_at_pregnancy_start",
        "has_fgr",
        # "condition_id", 
        # "bmi_min",
        # "bmi_max",
        "test_name", 
        "cpt_codes",
        "start_week", 
        "end_week", 
        "frequency_type",
        "frequency_value",
        "stop_condition", 
        "special_notes", 
        # "spl_notes", 
        # "guideline_section", 
        # "page_ref",
        "window_start", 
        "window_end",
        "frequency_hits", 
        "actual_tracking",
        "is_there_a_gap",
        "gap_details",
        "gap",
    ]
    for c in out_cols:
        if c not in out_df.columns:
            out_df[c] = ""
    extras = [c for c in out_df.columns if c not in out_cols]
    out_df = out_df[out_cols + extras]

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    local_path = args.local_out.strip() or f"evicore_join_combined_{ts}.csv"
    out_df.to_csv(local_path, index=False)
    log(f"[JOIN] Wrote local: {local_path} ({len(out_df)} rows)")

    out_prefix = (args.output_prefix or "").lstrip("/")
    if out_prefix and not out_prefix.endswith("/"):
        out_prefix += "/"
    out_blob = f"{out_prefix}evicore_join_combined_{ts}.csv"

    with open(local_path, "rb") as f:
        _upload_bytes(bsc, BlobLoc(args.output_container, out_blob), f.read())

    log(f"[JOIN] Uploaded: container={args.output_container} blob={out_blob}")


if __name__ == "__main__":
    main()
