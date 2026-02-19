#!/usr/bin/env python3
"""
Patient_join_evicore_obesity.py

Joins the latest Step4 pregnancy/BMI care-tracking output with:
  - EncounterData.xlsx (encounters with CPT codes and dates of service)
  - obesity_rules_norm.csv (normalized eviCore obesity rules)

Behavior
  - Finds newest Step4 file under step4-container/step4-prefix that contains Care_Tracking_output
  - Uses BMI_at_pregnancy_start from Step4 and keeps BMI >= 30 as obesity cohort
  - Filters rules to condition_id containing "OBESITY" and applies bmi_min/bmi_max checks (if present)
  - Matches CPTs within the rule window (gestational start/end weeks) with a small grace window
  - Flags gaps for once, weekly, every_n_weeks, every_n_to_m_weeks

Example (Cloud Shell)
  python3 Patient_join_evicore_obesity.py \
    --account-name genafuncapp \
    --encounters-container input \
    --encounters-blob EncounterData.xlsx \
    --rules-container input \
    --rules-blob obesity_rules_norm.csv \
    --step4-container output \
    --step4-prefix "" \
    --output-container output \
    --output-prefix ""
"""

import argparse
import os
import re
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient


GRACE_WEEKS = 3
_CPT_RE = re.compile(r"\b(\d{5})\b")
_STEP4_TS_RE = re.compile(r"step4_(\d{8})_(\d{6})", re.IGNORECASE)


def log(*args, **kwargs):
    print(*args, **kwargs)


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


def _fmt_mdy(ts: pd.Timestamp) -> str:
    if pd.isna(ts):
        return ""
    try:
        return ts.strftime("%-m/%-d/%Y")
    except Exception:
        return ts.strftime("%m/%d/%Y")


def _find_col(df: pd.DataFrame, candidates: List[str]) -> str:
    lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand in df.columns:
            return cand
        if cand.lower() in lower:
            return lower[cand.lower()]
    raise KeyError(f"Missing column. Tried: {candidates}. Found: {list(df.columns)}")


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
    preg_start: pd.Timestamp, start_week: Optional[float], end_week: Optional[float]
) -> Tuple[pd.Timestamp, pd.Timestamp]:
    win_start = (
        preg_start + timedelta(weeks=float(start_week))
        if start_week is not None and not pd.isna(start_week)
        else preg_start
    )
    win_end = (
        preg_start + timedelta(weeks=float(end_week))
        if end_week is not None and not pd.isna(end_week)
        else preg_start + timedelta(weeks=40)
    )
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

        if "bmicare_tracking_output" not in low and "care_tracking_output" not in low:
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

    candidates.sort(
        key=lambda x: (
            x[0] or datetime.min.replace(tzinfo=timezone.utc),
            x[1] or datetime.min.replace(tzinfo=timezone.utc),
        )
    )
    return candidates[-1][2]


def load_step4_df(raw_csv: bytes) -> pd.DataFrame:
    df = pd.read_csv(BytesIO(raw_csv))

    if "patient_id" not in df.columns:
        pid_col = _find_col(df, ["patient_id", "patientid", "PatientID", "PATIENT_ID", "MRN"])
        df = df.rename(columns={pid_col: "patient_id"})
    df["patient_id"] = df["patient_id"].apply(_normalize_patient_id)

    for c in ["start_of_pregnancy_date", "date_of_birth", "BMI_DOS"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    return df


def load_rules_df(raw_csv: bytes) -> pd.DataFrame:
    df = pd.read_csv(BytesIO(raw_csv))

    required = [
        "condition_id",
        "test_name",
        "cpt_codes",
        "start_week",
        "end_week",
        "frequency_type",
        "frequency_value",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"obesity_rules_norm.csv missing required columns: {missing}")

    df = df.copy()
    df["condition_id"] = df["condition_id"].fillna("").astype(str).str.strip()
    df["test_name"] = df["test_name"].fillna("").astype(str).str.strip()
    df["start_week"] = df["start_week"].apply(_parse_week)
    df["end_week"] = df["end_week"].apply(_parse_week)
    df["frequency_type"] = df["frequency_type"].fillna("").astype(str).str.strip().str.lower()
    df["frequency_value"] = df["frequency_value"].replace({pd.NA: "", None: ""}).astype(str).str.strip()

    if "bmi_min" in df.columns:
        df["bmi_min"] = df["bmi_min"].apply(_as_float_or_none)
    else:
        df["bmi_min"] = None

    if "bmi_max" in df.columns:
        df["bmi_max"] = df["bmi_max"].apply(_as_float_or_none)
    else:
        df["bmi_max"] = None

    for col in ["stop_condition", "special_notes", "guideline_section"]:
        if col not in df.columns:
            df[col] = ""
        else:
            df[col] = df[col].fillna("").astype(str).str.strip()

    return df.reset_index(drop=True)


def load_encounters_df(raw_xlsx: bytes) -> pd.DataFrame:
    enc = pd.read_excel(BytesIO(raw_xlsx))

    pid_col = _find_col(enc, ["patientid", "patient_id", "PatientID", "PATIENT_ID", "MRN"])
    cpt_col = _find_col(enc, ["CPTCode", "CPT_CODE", "CPT", "cpt", "CPT codes"])
    dos_col = _find_col(enc, ["DOS", "Date of Service", "Service Date"])

    enc = enc.rename(columns={pid_col: "enc_patient_id", cpt_col: "enc_cpt", dos_col: "enc_dos"})
    enc["enc_patient_id"] = enc["enc_patient_id"].apply(_normalize_patient_id)
    enc["enc_cpt"] = enc["enc_cpt"].apply(_normalize_cpt)
    enc["enc_dos"] = pd.to_datetime(enc["enc_dos"], errors="coerce")

    return enc[["enc_patient_id", "enc_cpt", "enc_dos"]].copy()


def build_enc_index(enc: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    out: Dict[str, pd.DataFrame] = {}
    enc = enc.dropna(subset=["enc_patient_id"]).copy()
    enc["enc_patient_id"] = enc["enc_patient_id"].astype(str)
    for pid, g in enc.groupby("enc_patient_id"):
        out[pid] = g.sort_values("enc_dos").reset_index(drop=True)
    return out


def obesity_class_from_bmi(bmi: float) -> str:
    if bmi < 30:
        return ""
    if bmi < 35:
        return "OBESITY_CLASS_I"
    if bmi < 40:
        return "OBESITY_CLASS_II"
    return "OBESITY_CLASS_III"


def _parse_every_n_to_m(freq_value: str) -> Optional[int]:
    if not freq_value:
        return None
    s = str(freq_value).strip().lower()
    m = re.search(r"(\d+)\s*(?:to|-)\s*(\d+)", s)
    if not m:
        return None
    a = int(m.group(1))
    b = int(m.group(2))
    return max(a, b)


def compliance_gap(
    hits: pd.DataFrame,
    win_start: pd.Timestamp,
    win_end: pd.Timestamp,
    freq_type: str,
    freq_value: str,
) -> Tuple[str, str]:
    freq_type = (freq_type or "").strip().lower()
    fv = _as_float_or_none(freq_value)
    fv_int = int(fv) if fv is not None else None

    if hits is None or hits.empty:
        return "Yes", "0 hits"

    if not freq_type or freq_type == "once":
        return "No", f"{len(hits)} hits"

    total_days = max(1, int((win_end - win_start).days) + 1)
    total_weeks = max(1, int((total_days + 6) // 7))

    d = hits["enc_dos"].copy()
    idx = ((d - win_start).dt.days // 7).clip(lower=0)

    if freq_type == "weekly":
        need = fv_int if fv_int is not None else 1
        counts = idx.value_counts().to_dict()
        bad = []
        for w in range(total_weeks):
            c = int(counts.get(w, 0))
            if c < need:
                bad.append(f"wk{w+1}:{c}/{need}")
        if bad:
            return "Yes", " | ".join(bad[:12]) + (" | ..." if len(bad) > 12 else "")
        return "No", f"weekly ok ({need}/week)"

    if freq_type == "every_n_weeks":
        block = fv_int if fv_int is not None else 1
        block_idx = (idx // block)
        n_blocks = max(1, int((total_weeks + block - 1) // block))
        counts = block_idx.value_counts().to_dict()
        bad = []
        for b in range(n_blocks):
            c = int(counts.get(b, 0))
            if c < 1:
                bad.append(f"blk{b+1}:0/1")
        if bad:
            return "Yes", " | ".join(bad[:12]) + (" | ..." if len(bad) > 12 else "")
        return "No", f"q{block}w ok (>=1/block)"

    if freq_type == "every_n_to_m_weeks":
        max_block = _parse_every_n_to_m(freq_value) or 6
        block_idx = (idx // max_block)
        n_blocks = max(1, int((total_weeks + max_block - 1) // max_block))
        counts = block_idx.value_counts().to_dict()
        bad = []
        for b in range(n_blocks):
            c = int(counts.get(b, 0))
            if c < 1:
                bad.append(f"blk{b+1}:0/1")
        if bad:
            return "Yes", " | ".join(bad[:12]) + (" | ..." if len(bad) > 12 else "")
        return "No", f"q{max_block}w ok (max interval from range)"

    return "No", f"{len(hits)} hits"


def select_rules_for_patient(rules: pd.DataFrame, bmi: float) -> pd.DataFrame:
    df = rules.copy()
    df = df[df["condition_id"].str.contains("obesity", case=False, na=False)].copy()

    m1 = df["bmi_min"].isna() | (df["bmi_min"] <= bmi)
    m2 = df["bmi_max"].isna() | (bmi <= df["bmi_max"])
    df = df[m1 & m2].copy()

    return df.reset_index(drop=True)


def build_output(step4: pd.DataFrame, rules: pd.DataFrame, encounters: pd.DataFrame) -> pd.DataFrame:
    enc_by_pid = build_enc_index(encounters)

    comp_col = "pregnancy_complications" if "pregnancy_complications" in step4.columns else None

    bmi_col = None
    for cand in ["BMI_at_pregnancy_start", "BMI_at_pregnancy_start_num", "bmi_at_pregnancy_start", "BMI"]:
        if cand in step4.columns:
            bmi_col = cand
            break
    if bmi_col is None:
        raise KeyError("Step4 output is missing BMI_at_pregnancy_start (or a compatible BMI column).")


    results: List[Dict[str, Any]] = []

    for _, p in step4.iterrows():
        pid = _normalize_patient_id(p.get("patient_id", ""))
        preg_start = _to_dt(p.get("start_of_pregnancy_date"))
        if not pid or pd.isna(preg_start):
            continue

        bmi = _as_float_or_none(p.get(bmi_col))
        if bmi is None or bmi < 30:
            continue

        ob_class = obesity_class_from_bmi(float(bmi))
        rules_p = select_rules_for_patient(rules, float(bmi))
        if rules_p.empty:
            continue

        enc_p = enc_by_pid.get(pid, pd.DataFrame(columns=["enc_patient_id", "enc_cpt", "enc_dos"]))

        for _, r in rules_p.iterrows():
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

                hits["dos_key"] = hits["enc_dos"].dt.strftime("%Y-%m-%d")
                hits = (
                    hits.sort_values(["enc_dos", "enc_cpt"])
                    .drop_duplicates(subset=["dos_key"])
                    .drop(columns=["dos_key"])
                )

            freq = int(len(hits))
            tracking = ""
            if freq > 0:
                tracking = " | ".join(
                    f"{row.enc_cpt}:{_fmt_mdy(row.enc_dos)}"
                    for row in hits.sort_values("enc_dos").itertuples(index=False)
                )

            freq_type = str(r.get("frequency_type", "")).strip().lower()
            freq_value = str(r.get("frequency_value", "")).strip()

            gap_flag, gap_details = compliance_gap(
                hits=hits,
                win_start=win_start,
                win_end=win_end,
                freq_type=freq_type,
                freq_value=freq_value,
            )

            results.append({
                "patient_id": pid,
                "first_name": p.get("first_name", ""),
                "last_name": p.get("last_name", ""),
                "date_of_birth": _fmt_mdy(_to_dt(p.get("date_of_birth"))),
                "start_of_pregnancy_date": _fmt_mdy(preg_start),
                "payer": p.get("payer", ""),
                "pregnancy_complications": p.get(comp_col, "") if comp_col else "",
                "BMI_at_pregnancy_start": bmi,
                "obesity_class": ob_class,
                "rule_condition_id": str(r.get("condition_id", "")).strip(),
                "CPT codes": str(r.get("cpt_codes", "")).strip(),
                "Test": str(r.get("test_name", "")).strip(),
                "When": str(r.get("stop_condition", "")).strip(),
                "starting date according to Evicore": _fmt_mdy(win_start),
                "Latest date according to Evicore": _fmt_mdy(win_end),
                "Frequency": freq,
                "Actual tracking": tracking,
                "Is there a gap?": gap_flag,
                "Gap details": gap_details,
            })

    out_df = pd.DataFrame(results)
    if out_df.empty:
        return out_df

    col_order = [
        "patient_id", "first_name", "last_name", "date_of_birth", "start_of_pregnancy_date",
        "payer", "pregnancy_complications", "BMI_at_pregnancy_start", "obesity_class", "rule_condition_id",
        "CPT codes", "Test", "When", "starting date according to Evicore", "Latest date according to Evicore",
        "Frequency", "Actual tracking", "Is there a gap?", "Gap details"
    ]
    for c in col_order:
        if c not in out_df.columns:
            out_df[c] = ""
    return out_df[col_order]


def main() -> None:
    ap = argparse.ArgumentParser(description="Join latest Step4 output with encounters and normalized obesity rules.")
    ap.add_argument("--account-name", required=True, help="Storage account name (example: genafuncapp)")
    ap.add_argument("--encounters-container", default="input", help="Container holding EncounterData.xlsx")
    ap.add_argument("--encounters-blob", default="EncounterData.xlsx", help="Blob name inside encounters-container")
    ap.add_argument("--rules-container", default="input", help="Container holding obesity_rules_norm.csv")
    ap.add_argument("--rules-blob", default="obesity_rules_norm.csv", help="Blob name inside rules-container")
    ap.add_argument("--step4-container", default="output", help="Container holding step4_*_Care_Tracking_output.csv")
    ap.add_argument("--step4-prefix", default="", help="Prefix inside step4-container (usually empty)")
    ap.add_argument("--output-container", default="output", help="Container to upload join output")
    ap.add_argument("--output-prefix", default="", help="Prefix inside output-container (optional)")
    ap.add_argument("--local-out", default="", help="Local output path (optional)")
    args = ap.parse_args()

    account_url = f"https://{args.account_name}.blob.core.windows.net"
    bsc = _blob_service_client(account_url)

    step4_blob = pick_latest_step4(bsc, args.step4_container, args.step4_prefix)
    log(f"[JOIN] Latest Step4: container={args.step4_container} blob={step4_blob}")

    step4_raw = _download_bytes(bsc, BlobLoc(args.step4_container, step4_blob))
    rules_raw = _download_bytes(bsc, BlobLoc(args.rules_container, args.rules_blob))
    enc_raw = _download_bytes(bsc, BlobLoc(args.encounters_container, args.encounters_blob))

    step4_df = load_step4_df(step4_raw)
    rules_df = load_rules_df(rules_raw)
    enc_df = load_encounters_df(enc_raw)

    out_df = build_output(step4_df, rules_df, enc_df)
    if out_df.empty:
        log("[JOIN] No output rows (no obesity patients found or no matching rules)." )
        return

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    local_path = Path(args.local_out) if args.local_out else Path(f"step4_obesity_transactions_{ts}.csv")
    out_df.to_csv(local_path, index=False)
    log(f"[JOIN] Wrote local: {local_path} ({len(out_df)} rows)")

    out_prefix = (args.output_prefix or "").lstrip("/")
    if out_prefix and not out_prefix.endswith("/"):
        out_prefix += "/"
    out_blob = f"{out_prefix}step4_obesity_transactions_{ts}.csv"

    _upload_bytes(bsc, BlobLoc(args.output_container, out_blob), local_path.read_bytes())
    log(f"[JOIN] Uploaded: container={args.output_container} blob={out_blob}")


if __name__ == "__main__":
    main()
