#!/usr/bin/env python3

import os
import sys
import argparse
from datetime import datetime, timezone

from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient, ContentSettings
from azure.core.exceptions import ResourceNotFoundError


def get_bsc(storage_account_url: str) -> BlobServiceClient:
    cred = DefaultAzureCredential()
    return BlobServiceClient(account_url=storage_account_url, credential=cred)


def download_blob(bsc, container: str, blob_name: str, local_path: str):
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    print(f"[INFO] Step2: Downloading {container}/{blob_name} -> {local_path}")
    try:
        data = (
            bsc.get_blob_client(container, blob_name)
            .download_blob()
            .readall()
        )
    except ResourceNotFoundError:
        print(
            f"[ERROR] Step2: Blob NOT found: container='{container}', blob='{blob_name}'"
        )
        sys.exit(1)

    with open(local_path, "wb") as f:
        f.write(data)


def upload_file(bsc, container: str, local_path: str, blob_name: str):
    print(f"[INFO] Step2: Uploading {local_path} -> {container}/{blob_name}")
    with open(local_path, "rb") as f:
        bsc.get_blob_client(container, blob_name).upload_blob(
            f,
            overwrite=True,
            content_settings=ContentSettings(content_type="text/csv"),
        )


def find_newest_csv(folder: str) -> str:
    files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(".csv")
    ]
    if not files:
        print(f"[ERROR] Step2: No CSV files found in {folder}")
        sys.exit(1)
    return max(files, key=os.path.getmtime)


def clean_patient_id(patient_id):
    if patient_id is None:
        return ""

    s = str(patient_id).strip()

    if "." in s:
        s = s.split(".", 1)[0]

    if s.isdigit():
        s = str(int(s))

    return s


def convert_date_format(date_str, target_format="%m/%d/%Y"):
    from datetime import datetime as dt
    import pandas as pd

    if date_str is None or str(date_str).strip() == "":
        return None

    try:
        text = str(date_str).strip()
        input_formats = ["%Y-%m-%d", "%m/%d/%Y", "%Y/%m/%d"]

        for fmt in input_formats:
            try:
                parsed_date = dt.strptime(text, fmt)
                return parsed_date.strftime(target_format)
            except ValueError:
                continue

        parsed = pd.to_datetime(text, errors="coerce")
        if not pd.isna(parsed):
            return parsed.strftime(target_format)

        print(f"[WARN] Step2-local: could not convert date {date_str}")
        return None
    except Exception as e:
        print(f"[WARN] Step2-local: could not convert date {date_str}: {e}")
        return None


def extract_date_from_reason(reason_date_str):
    if reason_date_str is None:
        return None

    text = str(reason_date_str).strip()
    if not text:
        return None

    if ":" in text:
        text = text.split(":", 1)[0]

    return text.strip()


def clean_bmi_data(bmi_df):
    import pandas as pd

    bmi_df = bmi_df.copy()
    bmi_df["patientid"] = bmi_df["patientid"].apply(clean_patient_id)
    bmi_df["cln_enc_date_dt"] = pd.to_datetime(
        bmi_df["cln enc date"], errors="coerce"
    )

    bmi_df_filtered = bmi_df.dropna(subset=["cln_enc_date_dt", "enc BMI"])

    bmi_df_sorted = bmi_df_filtered.sort_values(
        ["patientid", "cln_enc_date_dt"],
        ascending=[True, False],
    )

    bmi_df_cleaned = bmi_df_sorted.drop_duplicates(subset=["patientid"])

    return bmi_df_cleaned


def run_step2_local(step1_csv_path: str, bmi_csv_path: str, output_dir: str):
    import pandas as pd

    os.makedirs(output_dir, exist_ok=True)

    print(f"[INFO] Step2-local: reading Step1 data from {step1_csv_path}")
    step1_df = pd.read_csv(step1_csv_path, encoding="latin1")

    print(f"[INFO] Step2-local: reading BMI data from {bmi_csv_path}")
    bmi_df = pd.read_csv(bmi_csv_path, encoding="latin1")

    if "patientid" not in bmi_df.columns and "patient_id" in bmi_df.columns:
        bmi_df = bmi_df.rename(columns={"patient_id": "patientid"})
    bmi_df["patientid"] = bmi_df["patientid"].apply(clean_patient_id)

    if "patient_id" not in step1_df.columns:
        raise KeyError("Expected 'patient_id' column in Step1 output.")

    if "patientid" not in bmi_df.columns or "cln enc date" not in bmi_df.columns:
        raise KeyError("Expected 'patientid' and 'cln enc date' columns in BMI file.")

    if "reason_for_pregnancy_date" not in step1_df.columns:
        raise KeyError("Expected 'reason_for_pregnancy_date' column in Step1 output.")

    bmi_df_cleaned = clean_bmi_data(bmi_df)

    def get_bmi_for_patient(row):
        patient_id = clean_patient_id(row["patient_id"])

        dos_raw = extract_date_from_reason(row.get("reason_for_pregnancy_date"))
        if dos_raw is None or str(dos_raw).strip() == "":
            return None, None, "No date extracted"

        dos_formatted = convert_date_format(dos_raw)
        if dos_formatted is None:
            bmi_search = f"{patient_id} - Date formatting error for {dos_raw}"
            return None, dos_raw, bmi_search

        target_dt = pd.to_datetime(dos_raw, errors="coerce")

        matching_bmi = bmi_df_cleaned[
            (bmi_df_cleaned["patientid"] == patient_id)
            & (bmi_df_cleaned["cln_enc_date_dt"] == target_dt)
        ]

        if matching_bmi.empty:
            patient_bmis = bmi_df[bmi_df["patientid"] == patient_id]

            if not patient_bmis.empty:
                patient_bmis = patient_bmis.copy()
                patient_bmis["dos_dt"] = pd.to_datetime(
                    patient_bmis["cln enc date"], errors="coerce"
                )
                target_date = pd.to_datetime(dos_raw, errors="coerce")

                if not pd.isna(target_date):
                    patient_bmis["date_diff"] = (
                        patient_bmis["dos_dt"] - target_date
                    ).abs()
                    closest_bmi = patient_bmis.loc[
                        patient_bmis["date_diff"].idxmin()
                    ]

                    bmi_val = closest_bmi["enc BMI"]
                    bmi_date = closest_bmi["cln enc date"]
                    bmi_search = (
                        f"{patient_id} on {bmi_date} (closest to {dos_formatted})"
                    )
                    return bmi_val, dos_raw, bmi_search

        if not matching_bmi.empty:
            bmi_val = matching_bmi["enc BMI"].iloc[0]
            bmi_search = f"{patient_id} on {dos_formatted}"
            return bmi_val, dos_raw, bmi_search

        bmi_search = f"{patient_id} on {dos_formatted} - No BMI found"
        return None, dos_raw, bmi_search

    print(
        "[INFO] Step2-local: computing "
        "BMI_at_pregnancy_start / BMI_DOS / BMI_search ..."
    )
    bmi_results = step1_df.apply(get_bmi_for_patient, axis=1, result_type="expand")

    step1_df["BMI_at_pregnancy_start"] = bmi_results[0]
    # Updated 2026-01-28: Obesity tagging and category model
    # - Compute BMI_at_pregnancy_start_num
    # - Update pregnancy_complications with OBESITY_BASE / OBESITY_CLASS_* (so it matches rule condition_id)
    # - Update the generic category + category_reason columns
    import pandas as pd

    step1_df["BMI_at_pregnancy_start_num"] = pd.to_numeric(
        step1_df["BMI_at_pregnancy_start"], errors="coerce"
    )

    def _split_tags(s: str):
        if s is None:
            return []
        s = str(s).strip()
        if not s or s.lower() == "none":
            return []
        parts = [p.strip() for p in s.split(";")]
        return [p for p in parts if p]

    def _join_tags(tags):
        out = []
        seen = set()
        for t in tags:
            if not t:
                continue
            key = str(t).strip().lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(str(t).strip())
        return "; ".join(out)

    def _obesity_tags_from_bmi(bmi_val):
        if pd.isna(bmi_val):
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

    def _update_category_and_complications(row):
        bmi_val = row.get("BMI_at_pregnancy_start_num")
        new_ob_tags, new_ob_reasons = _obesity_tags_from_bmi(bmi_val)

        # category
        cat = _split_tags(row.get("category"))
        cat = [t for t in cat if not t.upper().startswith("OBESITY|")]
        cat.extend(new_ob_tags)

        # category_reason
        cat_r = _split_tags(row.get("category_reason"))
        cat_r = [t for t in cat_r if not t.upper().startswith("OBESITY|")]
        cat_r.extend(new_ob_reasons)

        # pregnancy_complications
        comp = row.get("pregnancy_complications")
        comp_s = "" if comp is None else str(comp).strip()
        if comp_s.lower() == "none":
            comp_s = ""
        parts = [p.strip() for p in comp_s.replace(";", ",").split(",") if p.strip()]

        remove_set = {
            "obesity complication",
            "basic_obesity",
            "class1_obesity",
            "class2_obesity",
            "class3_obesity",
            "obesity_base",
            "obesity_class_i",
            "obesity_class_ii",
            "obesity_class_iii",
        }
        cleaned = [p for p in parts if p.strip().lower() not in remove_set]

        # Add standardized obesity labels to match condition_id
        for t in new_ob_tags:
            label = t.split("|", 1)[1] if "|" in t else t
            cleaned.append(label)

        return _join_tags(cat), _join_tags(cat_r), "; ".join(_join_tags(cleaned).split("; ")) if cleaned else ""

    # Ensure the generic category columns exist even if Step1 did not create them.
    if "category" not in step1_df.columns:
        step1_df["category"] = ""
    if "category_reason" not in step1_df.columns:
        step1_df["category_reason"] = ""

    if "pregnancy_complications" not in step1_df.columns:
        step1_df["pregnancy_complications"] = ""

    upd = step1_df.apply(_update_category_and_complications, axis=1, result_type="expand")
    step1_df["category"] = upd[0]
    step1_df["category_reason"] = upd[1]
    step1_df["pregnancy_complications"] = upd[2]

    step1_df["BMI_DOS"] = bmi_results[1]
    step1_df["BMI_search"] = bmi_results[2]

    out_path = os.path.join(output_dir, "AttachedBMI_output.csv")
    step1_df.to_csv(out_path, index=False)
    print(f"[INFO] Step2-local: wrote {out_path}")


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Step 2: Attach BMI to current pregnant patients. "
            "Downloads Step1 output + bmi.csv from Azure, runs local "
            "business logic, and uploads the resulting CSV to Azure."
        )
    )

    parser.add_argument(
        "--storage-account-url",
        required=True,
        help="e.g. https://genafuncapp.blob.core.windows.net",
    )
    parser.add_argument(
        "--step1-container",
        required=True,
        help="Container where Step1 final CSV lives (e.g. 'intermediate')",
    )
    parser.add_argument(
        "--step1-blob",
        required=True,
        help="Blob name of Step1 output CSV (e.g. 'Current_Pregnant_Patients_....csv')",
    )
    parser.add_argument(
        "--bmi-container",
        default="input",
        help="Container where bmi.csv lives (default: 'input')",
    )
    parser.add_argument(
        "--bmi-blob",
        default="bmi.csv",
        help="Blob name of BMI file (default: 'bmi.csv')",
    )
    parser.add_argument(
        "--output-container",
        required=True,
        help="Container to upload Step2 CSV into (e.g. 'intermediate')",
    )

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
    os.makedirs(output_dir, exist_ok=True)

    run_step2_local(step1_local, bmi_local, output_dir)

    final_csv = find_newest_csv(output_dir)

    out_ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_blob_name = f"step2_{out_ts}_{os.path.basename(final_csv)}"

    upload_file(bsc, args.output_container, final_csv, out_blob_name)

    print(f"[INFO] Step2 final output blob: {args.output_container}/{out_blob_name}")


if __name__ == "__main__":
    main()
