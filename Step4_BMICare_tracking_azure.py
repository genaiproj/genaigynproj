#!/usr/bin/env python3

# Update 13-04-2026 for add sorting and rename output file as per mail from satarupa dated 13-04-2026

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
    print(f"[INFO] Step4: Downloading {container}/{blob_name} -> {local_path}")
    try:
        data = (
            bsc.get_blob_client(container, blob_name)
            .download_blob()
            .readall()
        )
    except ResourceNotFoundError:
        print(
            f"[ERROR] Step4: Blob NOT found: container='{container}', blob='{blob_name}'"
        )
        sys.exit(1)

    with open(local_path, "wb") as f:
        f.write(data)


def upload_file(bsc, container: str, local_path: str, blob_name: str):
    print(f"[INFO] Step4: Uploading {local_path} -> {container}/{blob_name}")
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
        print(f"[ERROR] Step4: No CSV files found in {folder}")
        sys.exit(1)
    return max(files, key=os.path.getmtime)


def run_step4_local(step3_csv_path: str, output_dir: str):
    import pandas as pd

    os.makedirs(output_dir, exist_ok=True)
    step3_df = pd.read_csv(step3_csv_path, encoding="latin1")
    care_df = step3_df

    # out_path = os.path.join(output_dir, "BMICare_Tracking_output.csv")
    #added 13-04-2026
    out_ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(output_dir, f"Current_Pregnant_Patients_{out_ts}.csv")
    care_df.to_csv(out_path, index=False)
    
    
def main():
    parser = argparse.ArgumentParser(
        description=(
            "Step 4: BMI care tracking. "
            "Downloads Step3 output from Azure, runs local care-tracking logic, "
            "and uploads the resulting CSV back to Azure."
        )
    )
    parser.add_argument(
        "--storage-account-url",
        required=True,
        help="e.g. https://genafuncapp.blob.core.windows.net",
    )
    parser.add_argument(
        "--step3-container",
        default="intermediate",
        help="Container where Step3 output CSV lives (default: intermediate)",
    )
    parser.add_argument(
        "--step3-blob",
        required=True,
        help=(
            "Blob name of Step3 output CSV "
            "(e.g. 'step3_YYYYMMDD_HHMMSS_SummarizedPatients_output.csv')"
        ),
    )
    parser.add_argument(
        "--output-container",
        default="output",
        help="Container to upload Step4 output CSV (default: output)",
    )
    args = parser.parse_args()

    bsc = get_bsc(args.storage_account_url)

    run_ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    workdir = os.path.join("/tmp", f"step4_run_{run_ts}")
    os.makedirs(workdir, exist_ok=True)

    step3_local = os.path.join(workdir, "Step3_output.csv")
    download_blob(bsc, args.step3_container, args.step3_blob, step3_local)

    output_dir = os.path.join(workdir, "Step4_Output")
    os.makedirs(output_dir, exist_ok=True)

    run_step4_local(step3_local, output_dir)

    final_csv = find_newest_csv(output_dir)

    #added on 14-04-2026
    out_ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    #out_blob_name = f"step4_{out_ts}_{os.path.basename(final_csv)}"
    base = os.path.splitext(os.path.basename(final_csv))[0]
    out_blob_name = f"WHF_{base}.csv"

    upload_file(bsc, args.output_container, final_csv, out_blob_name)

    print(f"[INFO] Step4 final output blob: {args.output_container}/{out_blob_name}")


if __name__ == "__main__":
    main()
