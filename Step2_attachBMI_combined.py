#!/usr/bin/env python3
"""
Step2_attachBMI_combined.py

What this does
1) Downloads Step1 output CSV and bmi.csv from Azure Blob Storage
2) Attaches BMI to each patient (exact pregnancy start date match if possible, otherwise closest BMI date)
3) Adds obesity program tags into category / category_reason
4) Preserves pregnancy_complications so it never becomes blank

Update history
- 2026-01-28: obesity tagging added (category model)
- 2026-02-06: fix pregnancy_complications getting blank after Step2
"""

import os
import sys
import argparse
from datetime import datetime, timezone
from typing import Dict, List, Optional, Tuple

from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient, ContentSettings
from azure.core.exceptions import ResourceNotFoundError


def get_bsc(storage_account_url: str) -> BlobServiceClient:
    return BlobServiceClient(account_url=storage_account_url, credential=DefaultAzureCredential())


def download_blob(bsc: BlobServiceClient, container: str, blob_name: str, local_path: str) -> None:
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    print(f"[INFO] Step2: Downloading {container}/{blob_name} -> {local_path}")
    try:
        data = bsc.get_blob_client(container, blob_name).download_blob().readall()
    except ResourceNotFoundError:
        print(f"[ERROR] Step2: Blob NOT found: container='{container}', blob='{blob_name}'")
        sys.exit(1)

    with open(local_path, "wb") as f:
        f.write(data)


def upload_file(bsc: BlobServiceClient, container: str, local_path: str, blob_name: str) -> None:
    print(f"[INFO] Step2: Uploading {local_path} -> {container}/{blob_name}")
    with open(local_path, "rb") as f:
        bsc.get_blob_client(container, blob_name).upload_blob(
            f,
            overwrite=True,
            content_settings=ContentSettings(content_type="text/csv"),
        )


def find_newest_csv(folder: str) -> str:
    files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".csv")]
    if not files:
        raise FileNotFoundError(f"No CSV files found in {folder}")
    return max(files, key=os.path.getmtime)


def clean_patient_id(x) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    if not s:
        return ""
    if "." in s:
        s = s.split(".", 1)[0]
    return s.strip()


def extract_date_part(x) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    if not s:
        return ""
    if ":" in s:
        s = s.split(":", 1)[0]
    return s.strip()


def split_semicolon_tags(s: str) -> List[str]:
    if s is None:
        return []
    text = str(s).strip()
    if not text or text.lower() == "none":
        return []
    return [p.strip() for p in text.split(";") if p.strip()]


def join_semicolon_tags(tags: List[str]) -> str:
    out: List[str] = []
    seen = set()
    for t in tags:
        t = str(t).strip()
        if not t:
            continue
        key = t.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(t)
    return "; ".join(out)


def obesity_tags_from_bmi(bmi_val) -> Tuple[List[str], List[str]]:
    import pandas as pd

    if bmi_val is None or (isinstance(bmi_val, float) and pd.isna(bmi_val)):
        return [], []

    try:
        bmi = float(bmi_val)
    except Exception:
        return [], []

    if bmi < 30:
        return [], []

    tags = ["OBESITY|OBESITY_BASE"]
    reasons = [f"OBESITY|OBESITY_BASE: BMI_at_pregnancy_start={bmi:.1f} (>=30)"]

    if bmi < 35:
        tags.append("OBESITY|OBESITY_CLASS_I")
        reasons.append(f"OBESITY|OBESITY_CLASS_I: BMI_at_pregnancy_start={bmi:.1f} (30-34.9)")
    elif bmi < 40:
        tags.append("OBESITY|OBESITY_CLASS_II")
        reasons.append(f"OBESITY|OBESITY_CLASS_II: BMI_at_pregnancy_start={bmi:.1f} (35-39.9)")
    else:
        tags.append("OBESITY|OBESITY_CLASS_III")
        reasons.append(f"OBESITY|OBESITY_CLASS_III: BMI_at_pregnancy_start={bmi:.1f} (>=40)")

    return tags, reasons


def normalize_bmi_columns(bmi_df):
    cols = {c: str(c).strip().lower() for c in bmi_df.columns}

    def pick(*candidates: str) -> Optional[str]:
        for cand in candidates:
            for orig, low in cols.items():
                if low == cand:
                    return orig
        return None

    patient_col = pick("patientid", "patient_id", "patient id")
    date_col = pick("cln enc date", "cln_enc_date", "encounter_date", "dos", "enc date")
    bmi_col = pick("enc bmi", "enc_bmi", "bmi", "enc bmi (kg/m2)")

    if patient_col is None or date_col is None or bmi_col is None:
        raise KeyError(
            "BMI file must contain patient id, encounter date, and BMI columns. "
            f"Found patient_col={patient_col}, date_col={date_col}, bmi_col={bmi_col}."
        )

    return patient_col, date_col, bmi_col


def run_step2_local(step1_csv_path: str, bmi_csv_path: str, output_dir: str) -> str:
    import pandas as pd

    os.makedirs(output_dir, exist_ok=True)

    print(f"[INFO] Step2-local: reading Step1 data from {step1_csv_path}")
    step1_df = pd.read_csv(step1_csv_path, dtype=str, keep_default_na=False, encoding="latin1")

    print(f"[INFO] Step2-local: reading BMI data from {bmi_csv_path}")
    bmi_df = pd.read_csv(bmi_csv_path, dtype=str, keep_default_na=False, encoding="latin1")

    if "patient_id" not in step1_df.columns:
        raise KeyError("Expected 'patient_id' column in Step1 output.")
    if "reason_for_pregnancy_date" not in step1_df.columns:
        raise KeyError("Expected 'reason_for_pregnancy_date' column in Step1 output.")

    # Pass-through columns that must not disappear
    if "pregnancy_complications" not in step1_df.columns:
        step1_df["pregnancy_complications"] = "None"
    if "category" not in step1_df.columns:
        step1_df["category"] = ""
    if "category_reason" not in step1_df.columns:
        step1_df["category_reason"] = ""

    patient_col, date_col, bmi_col = normalize_bmi_columns(bmi_df)

    bmi_df = bmi_df.copy()
    bmi_df[patient_col] = bmi_df[patient_col].apply(clean_patient_id)
    bmi_df["_enc_dt"] = pd.to_datetime(bmi_df[date_col], errors="coerce")
    bmi_df["_bmi_num"] = pd.to_numeric(bmi_df[bmi_col], errors="coerce")

    bmi_df = bmi_df.dropna(subset=["_enc_dt", "_bmi_num"]).copy()
    bmi_df.sort_values([patient_col, "_enc_dt"], ascending=[True, True], inplace=True)

    bmi_by_patient: Dict[str, pd.DataFrame] = {pid: g.copy() for pid, g in bmi_df.groupby(patient_col)}

    def find_bmi(patient_id: str, preg_start_raw: str) -> Tuple[Optional[float], str, str]:
        dos_raw = extract_date_part(preg_start_raw)
        if not dos_raw:
            return None, "", f"{patient_id} - missing pregnancy start date"

        target_dt = pd.to_datetime(dos_raw, errors="coerce")
        if pd.isna(target_dt):
            return None, dos_raw, f"{patient_id} - invalid pregnancy start date: {dos_raw}"

        g = bmi_by_patient.get(patient_id)
        if g is None or g.empty:
            return None, dos_raw, f"{patient_id} on {dos_raw} - no BMI rows"

        g2 = g.copy()
        g2["_enc_date"] = g2["_enc_dt"].dt.date
        tgt_date = target_dt.date()

        exact = g2[g2["_enc_date"] == tgt_date]
        if not exact.empty:
            row = exact.sort_values("_enc_dt", ascending=False).iloc[0]
            return float(row["_bmi_num"]), dos_raw, f"{patient_id} on {row[date_col]} (exact match)"

        g2["_diff"] = (g2["_enc_dt"] - target_dt).abs()
        row = g2.sort_values("_diff", ascending=True).iloc[0]
        return float(row["_bmi_num"]), dos_raw, f"{patient_id} on {row[date_col]} (closest to {dos_raw})"

    print("[INFO] Step2-local: computing BMI_at_pregnancy_start / BMI_DOS / BMI_search ...")
    bmi_vals: List[Optional[float]] = []
    bmi_dos: List[str] = []
    bmi_search: List[str] = []

    for _, r in step1_df.iterrows():
        pid = clean_patient_id(r.get("patient_id"))
        v, d, s = find_bmi(pid, r.get("reason_for_pregnancy_date"))
        bmi_vals.append(v)
        bmi_dos.append(d)
        bmi_search.append(s)

    step1_df["BMI_at_pregnancy_start"] = bmi_vals
    step1_df["BMI_DOS"] = bmi_dos
    step1_df["BMI_search"] = bmi_search
    step1_df["BMI_at_pregnancy_start_num"] = pd.to_numeric(step1_df["BMI_at_pregnancy_start"], errors="coerce")

    # Update obesity program tags and keep pregnancy_complications non-empty
    new_cats: List[str] = []
    new_reasons: List[str] = []
    new_comps: List[str] = []

    for _, r in step1_df.iterrows():
        bmi_num = r.get("BMI_at_pregnancy_start_num")
        ob_tags, ob_reasons = obesity_tags_from_bmi(bmi_num)

        cat = split_semicolon_tags(r.get("category"))
        cat = [t for t in cat if not str(t).upper().startswith("OBESITY|")]
        cat.extend(ob_tags)
        new_cats.append(join_semicolon_tags(cat))

        cr = split_semicolon_tags(r.get("category_reason"))
        cr = [t for t in cr if not str(t).upper().startswith("OBESITY|")]
        cr.extend(ob_reasons)
        new_reasons.append(join_semicolon_tags(cr))

        comp_raw = str(r.get("pregnancy_complications") or "").strip()
        if not comp_raw or comp_raw.lower() in ("nan", "null"):
            comp_raw = "None"

        comps: List[str] = []
        if comp_raw.lower() != "none":
            text = comp_raw.replace(";", ",")
            comps = [p.strip() for p in text.split(",") if p.strip()]

        if ob_tags and not any("obesity" in c.lower() for c in comps):
            comps.append("Obesity Complication")

        new_comps.append(", ".join(comps) if comps else "None")

    step1_df["category"] = new_cats
    step1_df["category_reason"] = new_reasons
    step1_df["pregnancy_complications"] = new_comps

    out_path = os.path.join(output_dir, "AttachedBMI_output.csv")
    step1_df.to_csv(out_path, index=False, encoding="utf-8")
    print(f"[INFO] Step2-local: wrote {out_path}")
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Step 2: Attach BMI to current pregnant patients. "
            "Downloads Step1 output + bmi.csv from Azure, runs local logic, "
            "and uploads the resulting CSV to Azure."
        )
    )

    parser.add_argument("--storage-account-url", required=True)
    parser.add_argument("--step1-container", required=True)
    parser.add_argument("--step1-blob", required=True)
    parser.add_argument("--bmi-container", default="input")
    parser.add_argument("--bmi-blob", default="bmi.csv")
    parser.add_argument("--output-container", required=True)

    args = parser.parse_args()

    bsc = get_bsc(args.storage_account_url)

    run_ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    workdir = os.path.join("/tmp", f"step2_run_{run_ts}")
    os.makedirs(workdir, exist_ok=True)

    step1_local = os.path.join(workdir, "Step1_output.csv")
    bmi_local = os.path.join(workdir, "bmi.csv")

    download_blob(bsc, args.step1_container, args.step1_blob, step1_local)
    download_blob(bsc, args.bmi_container, args.bmi_blob, bmi_local)

    output_dir = os.path.join(workdir, "Step2_Output")
    final_csv = run_step2_local(step1_local, bmi_local, output_dir)

    out_ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_blob_name = f"step2_{out_ts}_{os.path.basename(final_csv)}"

    upload_file(bsc, args.output_container, final_csv, out_blob_name)
    print(f"[INFO] Step2 final output blob: {args.output_container}/{out_blob_name}")


if __name__ == "__main__":
    main()
