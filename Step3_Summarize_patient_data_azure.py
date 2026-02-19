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
    print(f"[INFO] Step3: Downloading {container}/{blob_name} -> {local_path}")
    try:
        data = (
            bsc.get_blob_client(container, blob_name)
            .download_blob()
            .readall()
        )
    except ResourceNotFoundError:
        print(
            f"[ERROR] Step3: Blob NOT found: container='{container}', blob='{blob_name}'"
        )
        sys.exit(1)

    with open(local_path, "wb") as f:
        f.write(data)


def upload_file(bsc, container: str, local_path: str, blob_name: str):
    print(f"[INFO] Step3: Uploading {local_path} -> {container}/{blob_name}")
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
        print(f"[ERROR] Step3: No CSV files found in {folder}")
        sys.exit(1)
    return max(files, key=os.path.getmtime)


def run_step3_local(step2_csv_path: str, output_dir: str):
    import pandas as pd

    os.makedirs(output_dir, exist_ok=True)

    print(f"[INFO] Step3-local: reading Step2 data from {step2_csv_path}")
    step2_df = pd.read_csv(step2_csv_path, encoding="latin1")

    summary_df = step2_df.copy()

    out_path = os.path.join(output_dir, "SummarizedPatients_output.csv")
    summary_df.to_csv(out_path, index=False)
    print(f"[INFO] Step3-local: wrote {out_path}")


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Step 3: Summarize patient data. "
            "Downloads Step2 output from Azure, runs local summary logic, "
            "and uploads the resulting CSV back to Azure."
        )
    )
    parser.add_argument(
        "--storage-account-url",
        required=True,
        help="e.g. https://genafuncapp.blob.core.windows.net",
    )
    parser.add_argument(
        "--step2-container",
        default="intermediate",
        help="Container where Step2 output CSV lives (default: intermediate)",
    )
    parser.add_argument(
        "--step2-blob",
        required=True,
        help="Blob name of Step2 output CSV",
    )
    parser.add_argument(
        "--output-container",
        default="intermediate",
        help="Container to upload Step3 output CSV (default: intermediate)",
    )

    args = parser.parse_args()
    bsc = get_bsc(args.storage_account_url)

    run_ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    workdir = os.path.join("/tmp", f"step3_run_{run_ts}")
    os.makedirs(workdir, exist_ok=True)

    step2_local = os.path.join(workdir, "Step2_output.csv")
    download_blob(bsc, args.step2_container, args.step2_blob, step2_local)

    output_dir = os.path.join(workdir, "Step3_Output")
    os.makedirs(output_dir, exist_ok=True)

    run_step3_local(step2_local, output_dir)

    final_csv = find_newest_csv(output_dir)

    out_ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    out_blob_name = f"step3_{out_ts}_{os.path.basename(final_csv)}"

    upload_file(bsc, args.output_container, final_csv, out_blob_name)

    print(f"[INFO] Step3 final output blob: {args.output_container}/{out_blob_name}")


if __name__ == "__main__":
    main()
