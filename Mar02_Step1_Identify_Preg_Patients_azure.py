#!/usr/bin/env python3
"""
Step 1: Identify current pregnant patients (Azure).

Update 2026-01-28
- Introduced a generic category model used by downstream rule joins.
- Replaced diabetes_category and diabetes_category_reason with:
  - category
  - category_reason
- Standardized FGR flagging to a single field (has_fgr).

Key outputs
- pregnancy_complications: readable list (existing behavior)
- category: semi-colon separated tags like DIABETES|..., OBESITY|..., FGR|YES
- category_reason: short reasons for category tags
- has_fgr: Y or N

Note
category is intended as a stable interface so the same Step1 to Step4 pipeline can
feed both diabetes and obesity rule joins.
"""

import os
import sys
import re
import argparse
import pandas as pd
from datetime import datetime, timedelta

from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient, ContentSettings
from azure.core.exceptions import ResourceNotFoundError


def get_bsc(storage_account_url: str) -> BlobServiceClient:
    cred = DefaultAzureCredential()
    return BlobServiceClient(account_url=storage_account_url, credential=cred)


def download_blob(bsc, container: str, blob_name: str, local_path: str):
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    print(f"[INFO] Step1: Downloading {container}/{blob_name} -> {local_path}")
    try:
        data = bsc.get_blob_client(container, blob_name).download_blob().readall()
    except ResourceNotFoundError:
        print(f"[ERROR] Step1: Blob NOT found: container='{container}', blob='{blob_name}'")
        sys.exit(1)

    with open(local_path, "wb") as f:
        f.write(data)


def upload_file(bsc, container: str, local_path: str, blob_name: str):
    print(f"[INFO] Step1: Uploading {local_path} -> {container}/{blob_name}")
    with open(local_path, "rb") as f:
        bsc.get_blob_client(container, blob_name).upload_blob(
            f,
            overwrite=True,
            content_settings=ContentSettings(content_type="text/csv"),
        )


def find_newest_csv(folder: str) -> str:
    files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".csv")]
    if not files:
        print(f"[ERROR] Step1: No CSV files found in {folder}")
        sys.exit(1)
    return max(files, key=os.path.getmtime)


def _normalize_icd(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip().upper()
    if s.endswith(".0") and s[:-2].replace(".", "").isdigit():
        s = s[:-2]
    return s


def _flatten_dx(patient_group: pd.DataFrame) -> tuple[set[str], str]:
    code_cols = [c for c in patient_group.columns if re.match(r"(?i)^DiagICD_10_\d+$", str(c))]
    desc_cols = [c for c in patient_group.columns if re.match(r"(?i)^DiagICD_10_Desc\d+$", str(c))]

    if code_cols:
        code_cols.sort(key=lambda x: int(re.findall(r"(\d+)$", str(x))[0]))
    if desc_cols:
        desc_cols.sort(key=lambda x: int(re.findall(r"(\d+)$", str(x))[0]))

    codes: set[str] = set()
    descs: list[str] = []

    for _, row in patient_group.iterrows():
        for c in code_cols:
            v = _normalize_icd(row.get(c))
            if v:
                codes.add(v)
        for c in desc_cols:
            v = row.get(c)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            s = str(v).strip()
            if s:
                descs.append(s)

    return codes, " ".join(descs).lower()


def classify_diabetes(dx_codes: set[str], dx_desc_text: str) -> tuple[str | None, str]:
    """Business-driven diabetes classification (high-risk conditions).

    Outputs (when detected):
      - Pre_existing diabetes
      - Gestational diabetes in pregnancy, insulin controlled
      - Gestational diabetes, oral medication
      - Gestational diabetes, diet controlled

    If diabetes is present but does not match these buckets, this function returns
    None so downstream logic can keep the original (less granular) labels.
    """

    t = (dx_desc_text or "").lower()

    # 1) Exact phrase requested by business
    if "unsp pre-existing diabetes in pregnancy" in t:
        return "Pre_existing diabetes", "Matched Dx description: 'Unsp pre-existing diabetes in pregnancy'"

    # 2) Pre-existing diabetes in pregnancy (ICD O24.0–O24.3 and common diabetes codes)
    pre_o24_prefixes = ("O24.0", "O24.1", "O24.2", "O24.3", "O24.8", "O24.9")
    if any(code.startswith(pre) for code in dx_codes for pre in pre_o24_prefixes):
        return "Pre_existing diabetes", "ICD O24.0–O24.3/O24.8/O24.9"

    if any(code.startswith(("E08", "E09", "E10", "E11", "E13")) for code in dx_codes):
        return "Pre_existing diabetes", "ICD E08/E09/E10/E11/E13"

    if any(k in t for k in ["pre-existing diabetes", "pre existing diabetes", "pregestational", "type 1 diabetes", "type 2 diabetes", "t1dm", "t2dm"]):
        return "Pre_existing diabetes", "Dx description indicates pre-existing/pregestational diabetes"

    # 3) Gestational diabetes: control-specific subtypes (best-effort)
    # Diet controlled
    if any(code.startswith("O24.410") for code in dx_codes) or ("gestational" in t and "diet controlled" in t):
        return "Gestational diabetes, diet controlled", "Diet controlled (O24.410 or keyword)"

    # Insulin controlled
    if any(code.startswith("O24.414") for code in dx_codes) or (("insulin controlled" in t or "insulin-controlled" in t or "insulin requiring" in t) and ("gestational" in t or "diabetes in pregnancy" in t)):
        return "Gestational diabetes in pregnancy, insulin controlled", "Insulin controlled (O24.414 or keyword)"

    # Oral medication controlled
    if any(code.startswith("O24.415") for code in dx_codes) or ("controlled by oral" in t) or ("oral hypogly" in t):
        return "Gestational diabetes, oral medication", "Oral medication (O24.415 or keyword)"

    # Diabetes is mentioned, but not one of the supported buckets.
    if "diabetes" in t:
        # Avoid pure screening encounters.
        if re.search(r"\b(encounter\s+for\s+)?screen(?:ing)?\s+for\b.*\bdiabetes\b", t, re.IGNORECASE):
            return None, ""
        return None, "Diabetes mentioned but not mapped to the 3 requested buckets"

    return None, ""


def has_fgr(dx_codes: set[str], dx_desc_text: str) -> bool:
    # FGR / poor fetal growth
    if any(code.startswith(("O36.59", "O36.51", "O36.61")) for code in dx_codes):
        return True
    kw = ["fetal growth restriction", "poor fetal growth", "poor fetl grth", "susp placntl insuff", "iugr", "intrauterine growth restriction", "placental insufficiency"]
    return any(k in dx_desc_text for k in kw)


class PregnantPatientsIdentifier:
    def __init__(self, base_path: str):
        self.base_path = base_path
        self.output_dir = os.path.join(base_path, "Current_Pregnant_Patients")
        os.makedirs(self.output_dir, exist_ok=True)

    def load_encounters(self, filepath: str) -> pd.DataFrame:
        encounters_df = pd.read_excel(filepath, parse_dates=["DOS", "DOB"])

        encounters_df = encounters_df.rename(
            columns={
                "patientid": "patient_id",
                "DOS": "encounter_date",
                "patient lastname": "last_name",
                "patient firstname": "first_name",
            }
        )

        required = ["patient_id", "encounter_date", "first_name", "last_name", "DOB"]
        for col in required:
            if col not in encounters_df.columns:
                raise ValueError(f"Missing required column: {col}")

        return encounters_df

    def extract_gestational_age(self, description) -> int | None:
        weeks_pattern = r"(\d+)\s*(?:week|wk|weeks)"
        desc_str = str(description).lower()
        match = re.search(weeks_pattern, desc_str)
        return int(match.group(1)) if match else None

    def determine_pregnancy_details(self, patient_group: pd.DataFrame):
        diag_desc_columns = ["DiagICD_10_Desc1", "DiagICD_10_Desc2", "DiagICD_10_Desc3", "DiagICD_10_Desc4"]

        pregnancy_encounters = patient_group[
            patient_group[diag_desc_columns].apply(
                lambda row: any(self.extract_gestational_age(str(desc)) is not None for desc in row if pd.notna(desc)),
                axis=1,
            )
        ]

        if pregnancy_encounters.empty:
            return None

        sorted_encounters = pregnancy_encounters.sort_values("encounter_date")

        for _, encounter in sorted_encounters.iterrows():
            for col in diag_desc_columns:
                desc = encounter.get(col)
                weeks = self.extract_gestational_age(desc)
                if weeks is not None:
                    start_date = encounter["encounter_date"] - timedelta(weeks=weeks)
                    return start_date, str(desc), encounter["encounter_date"]

        return None

    def calculate_gestational_age(self, start_date, current_date=None) -> float:
        current_date = current_date or datetime.now()
        return round((current_date - start_date).days / 7, 1)

    def check_pregnancy_completion(self, patient_group: pd.DataFrame, start_date):
        completion_terms = [
            "delivery", "cesarean", "c-section", "vaginal delivery", "obstetric care", "postpartum care",
            "abortion", "miscarriage", "spontaneous abortion", "therapeutic abortion", "termination of pregnancy",
            "dilation and curettage", "d&c", "placental removal", "obstetrical care",
        ]

        post_start = patient_group[patient_group["encounter_date"] > start_date]
        completion = post_start[post_start["CPTDescription"].apply(lambda x: any(term in str(x).lower() for term in completion_terms))]

        if not completion.empty:
            first = completion.iloc[0]
            return True, f"{first['encounter_date'].strftime('%Y-%m-%d')}: {first['CPTDescription']}"

        # If beyond 42 weeks, look for postpartum follow-ups
        ga_weeks = (datetime.now() - start_date).days / 7
        if ga_weeks > 42:
            desc_cols = ["DiagICD_10_Desc1", "DiagICD_10_Desc2", "DiagICD_10_Desc3", "DiagICD_10_Desc4"]
            postpartum_terms = ["postpartum follow up", "postpartum checkup", "postpartum visit", "postpartum care"]
            postpartum = post_start[post_start[desc_cols].apply(
                lambda row: any(any(term in str(desc).lower() for term in postpartum_terms) for desc in row if pd.notna(desc)),
                axis=1
            )]
            if not postpartum.empty:
                first_post = postpartum.iloc[0]
                return True, f"Pregnancy completed (>42 weeks) - Postpartum follow-up on {first_post['encounter_date'].strftime('%Y-%m-%d')}"

        return False, None

    def get_latest_payer(self, patient_group: pd.DataFrame) -> str:
        sorted_encounters = patient_group.sort_values("encounter_date")
        payer_col = "payer" if "payer" in sorted_encounters.columns else "Payer"
        vals = sorted_encounters[payer_col].dropna()
        return vals.iloc[-1] if not vals.empty else "Unknown"

    def identify_complications(self, patient_group: pd.DataFrame):
        codes, desc_text = _flatten_dx(patient_group)
        diabetes_cat, diabetes_reason = classify_diabetes(codes, desc_text)
        fgr_flag = has_fgr(codes, desc_text)

        # Keep the original coarse-grained diabetes labels as a fallback.
        # If we successfully mapped diabetes into one of the requested granular buckets,
        # we will skip these generic patterns to avoid double-counting.
        screening_diabetes_regex = r"\b(encounter\s+for\s+)?screen(?:ing)?\s+for\b.*\bdiabetes\b"

        complication_patterns = [
            (r"obesity", "Obesity Complication"),
            (r"hypertension", "Hypertension in Pregnancy"),
            (r"anemia", "Anemia in Pregnancy"),
            (r"asthma", "Asthma in Pregnancy"),
            (r"thyroid", "Thyroid Disorder in Pregnancy"),
            (r"preeclampsia", "Preeclampsia"),
            (r"gestational hypertension", "Gestational Hypertension"),
            (r"placenta", "Placental Complications"),
            (r"advanced maternal age", "Advanced Maternal Age"),
            (r"multiple gestation", "Multiple Gestation"),
            (r"bleeding", "Antepartum Hemorrhage"),
            (r"infection", "Pregnancy-related Infection"),
            (r"thrombosis", "Thrombosis Risk"),
            (r"screening for oth suspected endocrine disorder", "Endocrine Disorder Screening"),
            (r"diseases of the circ sys comp pregnancy", "Circulatory System Complications"),
            (r"gestational diabetes", "Gestational Diabetes"),
            (r"diabetes", "Diabetes in Pregnancy"),
        ]

        desc_cols = ["DiagICD_10_Desc1", "DiagICD_10_Desc2", "DiagICD_10_Desc3", "DiagICD_10_Desc4"]

        complications = []
        for pattern, label in complication_patterns:
            if diabetes_cat and label in ("Gestational Diabetes", "Diabetes in Pregnancy"):
                continue
            comp_encounters = patient_group[patient_group[desc_cols].apply(
                lambda row: any(
                    (
                        (not re.search(screening_diabetes_regex, str(desc), re.IGNORECASE))
                        if label == "Diabetes in Pregnancy"
                        else True
                    )
                    and re.search(pattern, str(desc), re.IGNORECASE)
                    for desc in row
                    if pd.notna(desc)
                ),
                axis=1,
            )]
            if not comp_encounters.empty:
                complications.append(label)

        if diabetes_cat:
            complications.append(diabetes_cat)

        complications = sorted(set(complications))
        return complications, diabetes_cat, diabetes_reason, fgr_flag

    def process_encounters(self, filepath: str):
        encounters_df = self.load_encounters(filepath)

        ten_months_ago = datetime.now() - timedelta(days=300)
        recent = encounters_df[encounters_df["encounter_date"] >= ten_months_ago]

        pregnant_patients = []

        for patient_id, patient_group in recent.groupby("patient_id"):
            start_info = self.determine_pregnancy_details(patient_group)
            if start_info is None:
                continue

            start_date, preg_desc, preg_dos = start_info
            patient_info = patient_group.iloc[0]

            is_completed, completion_reason = self.check_pregnancy_completion(patient_group, start_date)
            if is_completed:
                continue

            current_ga = self.calculate_gestational_age(start_date)
            if current_ga > 43:
                continue

            maternal_age_at_start = preg_dos.year - patient_info["DOB"].year
            complications, diabetes_cat, diabetes_reason, fgr_flag = self.identify_complications(patient_group)
            # Updated 2026-01-28: Build generic category tags (diabetes, obesity signal, FGR).
            category_tags = []
            category_reasons = []

            if diabetes_cat:
                category_tags.append(f"DIABETES|{diabetes_cat}")
                if diabetes_reason:
                    category_reasons.append(f"DIABETES|{diabetes_cat}: {diabetes_reason}")
                else:
                    category_reasons.append(f"DIABETES|{diabetes_cat}: matched diabetes logic")

            if any(str(c).strip().lower() == "obesity complication" for c in (complications or [])):
                category_tags.append("OBESITY|DX_SIGNAL")
                category_reasons.append("OBESITY|DX_SIGNAL: diagnosis text indicates obesity")

            if fgr_flag:
                category_tags.append("FGR|YES")
                category_reasons.append("FGR|YES: matched FGR logic")

            def _dedupe_keep_order(items):
                seen = set()
                out = []
                for x in items:
                    if not x:
                        continue
                    k = str(x).strip().lower()
                    if k in seen:
                        continue
                    seen.add(k)
                    out.append(str(x).strip())
                return out

            category = "; ".join(_dedupe_keep_order(category_tags))
            category_reason = "; ".join(_dedupe_keep_order(category_reasons))
            has_fgr = "Y" if fgr_flag else "N"
            payer = self.get_latest_payer(patient_group)

            pregnant_patients.append({
                "patient_id": patient_id,
                "first_name": patient_info["first_name"],
                "last_name": patient_info["last_name"],
                "date_of_birth": patient_info["DOB"],
                "start_of_pregnancy_date": start_date,
                "reason_for_pregnancy_date": f"{preg_dos.strftime('%Y-%m-%d')}: {preg_desc}",
                "payer": payer,
                "current_gestational_age": current_ga,
                "pregnancy_complications": ", ".join(complications) if complications else "None",
                "category": category,
                "category_reason": category_reason,
                "has_fgr": has_fgr,
                "is_pregnancy_completed": is_completed,
                "reason_for_pregnancy_completion": completion_reason or "",
                "maternal_age_at_start_of_pregnancy": maternal_age_at_start,
            })

        df = pd.DataFrame(pregnant_patients)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_file = os.path.join(self.output_dir, f"Current_Pregnant_Patients_{timestamp}.csv")

        if df.empty:
            print("No currently pregnant patients found.")
            return None

        df.to_csv(out_file, index=False)
        print(f"Current pregnant patients saved to {out_file}")
        return df


def main():
    parser = argparse.ArgumentParser(description="Step 1: Identify current pregnant patients (generic category model).")
    parser.add_argument("--storage-account-url", required=True)
    parser.add_argument("--encounters-container", required=True)
    parser.add_argument("--encounters-blob", required=True)
    parser.add_argument("--output-container", required=True)
    args = parser.parse_args()

    bsc = get_bsc(args.storage_account_url)

    run_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    workdir = os.path.join("/tmp", f"step1_run_{run_ts}")
    os.makedirs(workdir, exist_ok=True)

    local_enc = os.path.join(workdir, args.encounters_blob)
    download_blob(bsc, args.encounters_container, args.encounters_blob, local_enc)

    identifier = PregnantPatientsIdentifier(workdir)
    pregnant_patients = identifier.process_encounters(local_enc)

    if pregnant_patients is None or pregnant_patients.empty:
        print("[INFO] Step1: No currently pregnant patients found. Nothing to upload.")
        return

    final_csv = find_newest_csv(identifier.output_dir)
    final_name = os.path.basename(final_csv)

    upload_file(bsc, args.output_container, final_csv, final_name)

    print(f"[INFO] Step1: Final output blob: {args.output_container}/{final_name}")


if __name__ == "__main__":
    main()