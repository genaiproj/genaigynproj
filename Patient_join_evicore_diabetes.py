#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
import re
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient


GRACE_WEEKS = 3
_CPT_RE = re.compile(r"\b(\d{5})\b")
_STEP4_TS_RE = re.compile(r"step4_(\d{8})_(\d{6})", re.IGNORECASE)


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


def _parse_week(v) -> Optional[float]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none"):
        return None
    m = re.search(r"(-?\d+(\.\d+)?)", s)
    return float(m.group(1)) if m else None


def _compute_window(preg_start: pd.Timestamp, start_week: Optional[float], end_week: Optional[float]) -> Tuple[pd.Timestamp, pd.Timestamp]:
    win_start = preg_start + timedelta(weeks=float(start_week)) if start_week is not None and not pd.isna(start_week) else preg_start
    win_end = preg_start + timedelta(weeks=float(end_week)) if end_week is not None and not pd.isna(end_week) else preg_start + timedelta(weeks=40)
    return win_start, win_end


def _parse_rule_cpts(raw) -> List[str]:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return []
    codes = _CPT_RE.findall(str(raw))
    return sorted(set(codes)) if codes else []


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
    prefix = (prefix or "").lstrip("/")  # never allow leading slash
    if prefix and not prefix.endswith("/"):
        prefix += "/"

    cc = bsc.get_container_client(container)
    candidates = []

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
        raise FileNotFoundError(f"No Step4 Care Tracking output found in container '{container}' under prefix '{prefix or '(root)'}'.")

    candidates.sort(key=lambda x: (x[0] or datetime.min.replace(tzinfo=timezone.utc), x[1] or datetime.min.replace(tzinfo=timezone.utc)))
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

    required = ["condition_id", "test_name", "cpt_codes", "start_week", "end_week", "frequency_type", "frequency_value", "stop_condition"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"diabetes_rules_norm.csv missing required columns: {missing}")

    df = df.copy()
    df["start_week"] = df["start_week"].apply(_parse_week)
    df["end_week"] = df["end_week"].apply(_parse_week)

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


def build_output(step4: pd.DataFrame, rules: pd.DataFrame, encounters: pd.DataFrame) -> pd.DataFrame:
    enc_by_pid = build_enc_index(encounters)

    # Prefer the generic category tags if they exist.
    # Fallback to pregnancy_complications for backward compatibility.
    step4 = step4.copy()
    # Updated 2026-01-29: always resolve pregnancy_complications column for output,
    # even when we filter diabetes patients using the generic category tags.
    comp_col = "pregnancy_complications" if "pregnancy_complications" in step4.columns else None
    if not comp_col:
        step4["pregnancy_complications"] = ""
        comp_col = "pregnancy_complications"

    # Prefer the generic category tags if they exist, fallback to pregnancy_complications text.
    cat_col = "category" if "category" in step4.columns else None
    if cat_col:
        step4[cat_col] = step4[cat_col].fillna("").astype(str)
        diabetic = step4[step4[cat_col].str.contains(r"DIABETES\|", case=False, na=False)].copy()
    else:
        step4[comp_col] = step4[comp_col].fillna("").astype(str)
        diabetic = step4[step4[comp_col].str.contains(r"diab", case=False, na=False)].copy()

    if diabetic.empty:
        return pd.DataFrame()

    results: List[dict] = []

    for _, p in diabetic.iterrows():
        pid = _normalize_patient_id(p.get("patient_id", ""))
        preg_start = _to_dt(p.get("start_of_pregnancy_date"))
        if not pid or pd.isna(preg_start):
            continue

        enc_p = enc_by_pid.get(pid, pd.DataFrame(columns=["enc_patient_id", "enc_cpt", "enc_dos"]))

        for _, r in rules.iterrows():
            cpt_list = _parse_rule_cpts(r.get("cpt_codes"))
            win_start, win_end = _compute_window(preg_start, r.get("start_week"), r.get("end_week"))

            grace_start = win_start - timedelta(weeks=GRACE_WEEKS)
            grace_end = win_end + timedelta(weeks=GRACE_WEEKS)

            hits = pd.DataFrame(columns=["enc_cpt", "enc_dos"])
            if not enc_p.empty and cpt_list:
                mask = (
                    enc_p["enc_cpt"].isin(cpt_list)
                    & (enc_p["enc_dos"] >= grace_start)
                    & (enc_p["enc_dos"] <= grace_end)
                )
                hits = enc_p.loc[mask, ["enc_cpt", "enc_dos"]].dropna().copy()

                hits["dos_key"] = hits["enc_dos"].dt.strftime("%Y-%m-%d")
                hits = hits.drop_duplicates(subset=["enc_cpt", "dos_key"]).drop(columns=["dos_key"])

            freq = int(len(hits))
            gap = "Yes" if freq == 0 else "No"

            tracking = ""
            if freq > 0:
                tracking = " | ".join(f"{row.enc_cpt}:{_fmt_mdy(row.enc_dos)}" for row in hits.sort_values("enc_dos").itertuples(index=False))

            results.append({
                "patient_id": pid,
                "first_name": p.get("first_name", ""),
                "last_name": p.get("last_name", ""),
                "date_of_birth": _fmt_mdy(_to_dt(p.get("date_of_birth"))),
                "start_of_pregnancy_date": _fmt_mdy(preg_start),
                "reason_for_pregnancy_date": str(p.get("reason_for_pregnancy_date", "")).strip(),
                "payer": p.get("payer", ""),
                "current_gestational_age": p.get("current_gestational_age", ""),
                "pregnancy_complications": p.get(comp_col, "") if comp_col else p.get("pregnancy_complications", ""),
                "is_pregnancy_completed": p.get("is_pregnancy_completed", ""),
                "reason_for_pregnancy_completion": p.get("reason_for_pregnancy_completion", ""),
                "maternal_age_at_start_of_pregnancy": p.get("maternal_age_at_start_of_pregnancy", ""),
                "BMI_at_pregnancy_start": p.get("BMI_at_pregnancy_start", ""),
                "BMI_DOS": _fmt_mdy(_to_dt(p.get("BMI_DOS"))),
                "BMI_search": p.get("BMI_search", ""),
                "CPT codes": str(r.get("cpt_codes", "")).strip(),
                "Test": str(r.get("test_name", "")).strip(),
                "When": str(r.get("stop_condition", "")).strip(),
                "starting date according to Evicore": _fmt_mdy(win_start),
                "Latest date according to Evicore": _fmt_mdy(win_end),
                "Frequency": freq,
                "Actual tracking": tracking,
                "Is there a gap?": gap,
            })

    out_df = pd.DataFrame(results)
    col_order = [
        "patient_id", "first_name", "last_name", "date_of_birth", "start_of_pregnancy_date",
        "reason_for_pregnancy_date", "payer", "current_gestational_age", "pregnancy_complications",
        "is_pregnancy_completed", "reason_for_pregnancy_completion", "maternal_age_at_start_of_pregnancy",
        "BMI_at_pregnancy_start", "BMI_DOS", "BMI_search", "CPT codes", "Test", "When",
        "starting date according to Evicore", "Latest date according to Evicore", "Frequency",
        "Actual tracking", "Is there a gap?"
    ]
    for c in col_order:
        if c not in out_df.columns:
            out_df[c] = ""
    return out_df[col_order]


def main() -> None:
    ap = argparse.ArgumentParser(description="Join latest Step4 BMICare tracking output with encounters + normalized diabetes rules.")
    ap.add_argument("--account-name", required=True, help="Storage account name (e.g., genafuncapp)")
    ap.add_argument("--encounters-container", default="input", help="Container holding EncounterData.xlsx")
    ap.add_argument("--encounters-blob", default="EncounterData.xlsx", help="Blob name inside encounters-container")
    ap.add_argument("--rules-container", default="input", help="Container holding diabetes_rules_norm.csv")
    ap.add_argument("--rules-blob", default="diabetes_rules_norm.csv", help="Blob name inside rules-container")
    ap.add_argument("--step4-container", default="output", help="Container holding step4_*_Care_Tracking_output.csv")
    ap.add_argument("--step4-prefix", default="", help="Prefix inside step4-container (usually empty)")
    ap.add_argument("--output-container", default="output", help="Container to upload join output")
    ap.add_argument("--output-prefix", default="", help="Prefix inside output-container (optional)")
    ap.add_argument("--local-out", default="", help="Local output path (default: ./step4_diabetes_transactions_<timestamp>.csv)")
    args = ap.parse_args()

    account_url = f"https://{args.account_name}.blob.core.windows.net"
    bsc = _blob_service_client(account_url)

    step4_blob = pick_latest_step4(bsc, args.step4_container, args.step4_prefix)
    print(f"[JOIN] Latest Step4: container={args.step4_container} blob={step4_blob}")

    step4_raw = _download_bytes(bsc, BlobLoc(args.step4_container, step4_blob))
    rules_raw = _download_bytes(bsc, BlobLoc(args.rules_container, args.rules_blob))
    enc_raw = _download_bytes(bsc, BlobLoc(args.encounters_container, args.encounters_blob))

    step4_df = load_step4_df(step4_raw)
    rules_df = load_rules_df(rules_raw)
    enc_df = load_encounters_df(enc_raw)

    out_df = build_output(step4_df, rules_df, enc_df)
    if out_df.empty:
        print("[JOIN] No output rows (no diabetic patients or no matching rules).")
        return

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    local_path = Path(args.local_out) if args.local_out else Path(f"step4_diabetes_transactions_{ts}.csv")
    out_df.to_csv(local_path, index=False)
    print(f"[JOIN] Wrote local: {local_path} ({len(out_df)} rows)")

    out_prefix = (args.output_prefix or "").lstrip("/")
    if out_prefix and not out_prefix.endswith("/"):
        out_prefix += "/"
    out_blob = f"{out_prefix}step4_diabetes_transactions_{ts}.csv"

    _upload_bytes(bsc, BlobLoc(args.output_container, out_blob), local_path.read_bytes())
    print(f"[JOIN] Uploaded: container={args.output_container} blob={out_blob}")


if __name__ == "__main__":
    main()