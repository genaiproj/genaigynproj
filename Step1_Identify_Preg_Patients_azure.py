#!/usr/bin/env python3

import os
import sys
import re
import argparse
import pandas as pd
from datetime import datetime, timedelta

from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient, ContentSettings
from azure.core.exceptions import ResourceNotFoundError


class PregnantPatientsIdentifier:
    def __init__(self, base_path):
        self.base_path = base_path
        self.output_dir = os.path.join(base_path, "Current_Pregnant_Patients")
        os.makedirs(self.output_dir, exist_ok=True)

    def load_encounters(self, filepath):
        try:
            encounters_df = pd.read_excel(filepath, parse_dates=["DOS", "DOB"])
            encounters_df = encounters_df.rename(
                columns={
                    "patientid": "patient_id",
                    "DOS": "encounter_date",
                    "patient lastname": "last_name",
                    "patient firstname": "first_name",
                }
            )

            required_columns = [
                "patient_id",
                "encounter_date",
                "first_name",
                "last_name",
                "DOB",
            ]
            for col in required_columns:
                if col not in encounters_df.columns:
                    raise ValueError(f"Missing required column: {col}")

            return encounters_df
        except Exception as e:
            print(f"Error loading encounters data: {e}")
            return None

    def extract_gestational_age(self, description):
        weeks_pattern = r"(\d+)\s*(?:week|wk|weeks)"
        desc_str = str(description).lower()
        match = re.search(weeks_pattern, desc_str)
        if match:
            return int(match.group(1))
        return None

    def determine_pregnancy_details(self, patient_group):
        diag_desc_columns = [
            "DiagICD_10_Desc1",
            "DiagICD_10_Desc2",
            "DiagICD_10_Desc3",
            "DiagICD_10_Desc4",
        ]

        pregnancy_encounters = patient_group[
            patient_group[diag_desc_columns].apply(
                lambda row: any(
                    self.extract_gestational_age(str(desc)) is not None
                    for desc in row
                    if pd.notna(desc)
                ),
                axis=1,
            )
        ]

        if pregnancy_encounters.empty:
            return None

        sorted_encounters = pregnancy_encounters.sort_values("encounter_date")

        for _, encounter in sorted_encounters.iterrows():
            for col in diag_desc_columns:
                desc = encounter[col]
                weeks = self.extract_gestational_age(desc)
                if weeks is not None:
                    start_date = encounter["encounter_date"] - timedelta(weeks=weeks)
                    return start_date, desc, encounter["encounter_date"]

        return None

    def calculate_gestational_age(self, start_date, current_date=None):
        if current_date is None:
            current_date = datetime.now()
        gestational_age_weeks = (current_date - start_date).days / 7
        return round(gestational_age_weeks, 1)

    def identify_complications(self, encounters_df):
        complication_patterns = [
            (r"obesity", "Obesity Complication"),
            (r"diabetes", "Diabetes in Pregnancy"),
            (r"hypertension", "Hypertension in Pregnancy"),
            (r"anemia", "Anemia in Pregnancy"),
            (r"asthma", "Asthma in Pregnancy"),
            (r"thyroid", "Thyroid Disorder in Pregnancy"),
            (r"preeclampsia", "Preeclampsia"),
            (r"gestational hypertension", "Gestational Hypertension"),
            (r"gestational diabetes", "Gestational Diabetes"),
            (r"placenta", "Placental Complications"),
            (r"advanced maternal age", "Advanced Maternal Age"),
            (r"multiple gestation", "Multiple Gestation"),
            (r"bleeding", "Antepartum Hemorrhage"),
            (r"infection", "Pregnancy-related Infection"),
            (r"thrombosis", "Thrombosis Risk"),
            (
                r"screening for oth suspected endocrine disorder",
                "Endocrine Disorder Screening",
            ),
            (
                r"diseases of the circ sys comp pregnancy",
                "Circulatory System Complications",
            ),
        ]

        description_fields = [
            "DiagICD_10_Desc1",
            "DiagICD_10_Desc2",
            "DiagICD_10_Desc3",
            "DiagICD_10_Desc4",
        ]

        complications = []
        for pattern, description in complication_patterns:
            comp_encounters = encounters_df[
                encounters_df[description_fields].apply(
                    lambda row: any(
                        re.search(pattern, str(desc), re.IGNORECASE)
                        for desc in row
                        if pd.notna(desc)
                    ),
                    axis=1,
                )
            ]
            if not comp_encounters.empty:
                complications.append(description)

        return list(set(complications))

    def check_pregnancy_completion(self, patient_group, start_date):
        completion_terms = [
            "delivery",
            "cesarean",
            "c-section",
            "vaginal delivery",
            "obstetric care",
            "postpartum care",
            "abortion",
            "miscarriage",
            "spontaneous abortion",
            "therapeutic abortion",
            "termination of pregnancy",
            "dilation and curettage",
            "d&c",
            "placental removal",
            "obstetrical care",
        ]

        post_start_encounters = patient_group[
            patient_group["encounter_date"] > start_date
        ]

        completion_encounters = post_start_encounters[
            post_start_encounters["CPTDescription"].apply(
                lambda x: any(term in str(x).lower() for term in completion_terms)
            )
        ]

        if not completion_encounters.empty:
            first_completion = completion_encounters.iloc[0]
            return (
                True,
                f"{first_completion['encounter_date'].strftime('%Y-%m-%d')}: "
                f"{first_completion['CPTDescription']}",
            )

        current_date = datetime.now()
        gestational_age_weeks = (current_date - start_date).days / 7

        if gestational_age_weeks > 42:
            diag_desc_columns = [
                "DiagICD_10_Desc1",
                "DiagICD_10_Desc2",
                "DiagICD_10_Desc3",
                "DiagICD_10_Desc4",
            ]

            postpartum_terms = [
                "postpartum follow up",
                "postpartum checkup",
                "postpartum visit",
                "postpartum care",
            ]

            postpartum_encounters = post_start_encounters[
                post_start_encounters[diag_desc_columns].apply(
                    lambda row: any(
                        any(term in str(desc).lower() for term in postpartum_terms)
                        for desc in row
                        if pd.notna(desc)
                    ),
                    axis=1,
                )
            ]

            if not postpartum_encounters.empty:
                first_postpartum = postpartum_encounters.iloc[0]
                return (
                    True,
                    "Pregnancy completed (>42 weeks) - Postpartum follow-up on "
                    f"{first_postpartum['encounter_date'].strftime('%Y-%m-%d')}",
                )

        return False, None

    def get_latest_payer(self, patient_group):
        sorted_encounters = patient_group.sort_values("encounter_date")
        payer_column = "payer" if "payer" in sorted_encounters.columns else "Payer"
        latest_payer = (
            sorted_encounters[payer_column].dropna().iloc[-1]
            if not sorted_encounters[payer_column].dropna().empty
            else "Unknown"
        )
        return latest_payer

    def process_encounters(self, filepath):
        encounters_df = self.load_encounters(filepath)
        if encounters_df is None:
            print("Could not load encounters data.")
            return None

        ten_months_ago = datetime.now() - timedelta(days=300)
        recent_encounters = encounters_df[
            encounters_df["encounter_date"] >= ten_months_ago
        ]

        pregnant_patients = []

        for patient_id, patient_group in recent_encounters.groupby("patient_id"):
            pregnancy_start_info = self.determine_pregnancy_details(patient_group)
            if pregnancy_start_info is None:
                continue

            pregnancy_start_date, preg_desc, preg_dos = pregnancy_start_info
            patient_info = patient_group.iloc[0]

            is_completed, completion_reason = self.check_pregnancy_completion(
                patient_group, pregnancy_start_date
            )
            if is_completed:
                continue

            maternal_age_at_start = preg_dos.year - patient_info["DOB"].year
            complications = self.identify_complications(patient_group)
            current_gestational_age = self.calculate_gestational_age(
                pregnancy_start_date
            )

            if current_gestational_age > 43:
                continue

            payer = self.get_latest_payer(patient_group)

            patient_record = {
                "patient_id": patient_id,
                "first_name": patient_info["first_name"],
                "last_name": patient_info["last_name"],
                "date_of_birth": patient_info["DOB"],
                "start_of_pregnancy_date": pregnancy_start_date,
                "reason_for_pregnancy_date": f"{preg_dos.strftime('%Y-%m-%d')}: {preg_desc}",
                "payer": payer,
                "current_gestational_age": current_gestational_age,
                "pregnancy_complications": ", ".join(complications)
                if complications
                else "None",
                "is_pregnancy_completed": is_completed,
                "reason_for_pregnancy_completion": completion_reason,
                "maternal_age_at_start_of_pregnancy": maternal_age_at_start,
            }

            pregnant_patients.append(patient_record)

        pregnant_patients_df = pd.DataFrame(pregnant_patients)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(
            self.output_dir, f"Current_Pregnant_Patients_{timestamp}.csv"
        )

        if not pregnant_patients_df.empty:
            pregnant_patients_df.to_csv(output_file, index=False)
            print(f"Current pregnant patients saved to {output_file}")
            return pregnant_patients_df
        else:
            print("No currently pregnant patients found.")
            return None


def get_bsc(storage_account_url: str) -> BlobServiceClient:
    cred = DefaultAzureCredential()
    return BlobServiceClient(account_url=storage_account_url, credential=cred)


def download_blob(bsc, container: str, blob_name: str, local_path: str):
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    print(f"[INFO] Step1: Downloading {container}/{blob_name} -> {local_path}")
    try:
        data = (
            bsc.get_blob_client(container, blob_name)
            .download_blob()
            .readall()
        )
    except ResourceNotFoundError:
        print(
            f"[ERROR] Step1: Blob NOT found: container='{container}', blob='{blob_name}'"
        )
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
    files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(".csv")
    ]
    if not files:
        print(f"[ERROR] Step1: No CSV files found in {folder}")
        sys.exit(1)
    return max(files, key=os.path.getmtime)


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Download EncounterData.xlsx from Azure, "
            "identify current pregnant patients, and upload "
            "Current_Pregnant_Patients_*.csv back to Azure."
        )
    )
    parser.add_argument(
        "--storage-account-url",
        required=True,
        help="e.g. https://genafuncapp.blob.core.windows.net",
    )
    parser.add_argument(
        "--encounters-container",
        required=True,
        help="Container where EncounterData.xlsx lives (e.g. 'input')",
    )
    parser.add_argument(
        "--encounters-blob",
        required=True,
        help="Blob name of encounters file (e.g. 'EncounterData.xlsx')",
    )
    parser.add_argument(
        "--output-container",
        required=True,
        help="Container to upload Current_Pregnant_Patients_*.csv (e.g. 'intermediate')",
    )

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
