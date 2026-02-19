#!/usr/bin/env python3

import os
import re
import sys
import subprocess
from datetime import datetime, timezone

from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient
from azure.core.exceptions import ResourceNotFoundError

STORAGE_ACCOUNT_URL = "https://genafuncapp.blob.core.windows.net"

SCRIPT_CONTAINER = "scripts"
INPUT_CONTAINER = "input"
INTERMEDIATE_CONTAINER = "intermediate"
FINAL_CONTAINER = "output"

SCRIPT_BLOBS = {
    "Step1_Identify_Preg_Patients_azure.py": "Step1_Identify_Preg_Patients_azure.py",
    "Step2_attachBMI_combined.py": "Step2_attachBMI_combined.py",
    "Step3_Summarize_patient_data_azure.py": "Step3_Summarize_patient_data_azure.py",
    "Step4_BMICare_tracking_azure.py": "Step4_BMICare_tracking_azure.py",
}

ENCOUNTER_BLOB = "EncounterData.xlsx"
BMI_BLOB = "bmi.csv"


def get_bsc(storage_account_url: str) -> BlobServiceClient:
    cred = DefaultAzureCredential()
    return BlobServiceClient(account_url=storage_account_url, credential=cred)


def download_blob(bsc, container: str, blob_name: str, local_path: str):
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    print(f"[ORCH] Downloading {container}/{blob_name} -> {local_path}")
    try:
        data = (
            bsc.get_blob_client(container, blob_name)
            .download_blob()
            .readall()
        )
    except ResourceNotFoundError:
        raise FileNotFoundError(
            f"Blob not found: container='{container}', blob='{blob_name}'"
        )

    with open(local_path, "wb") as f:
        f.write(data)


def run_step(script_path, step_args):
    if not os.path.isfile(script_path):
        raise FileNotFoundError(f"Script not found locally: {script_path}")

    cmd = ["python3", script_path] + step_args

    print(f"\n[ORCH] Running: {' '.join(cmd)}")
    completed = subprocess.run(cmd, capture_output=True, text=True)

    if completed.stdout:
        print(completed.stdout)
    if completed.stderr:
        print(completed.stderr, file=sys.stderr)

    if completed.returncode != 0:
        raise RuntimeError(
            f"Step script {os.path.basename(script_path)} failed with exit code "
            f"{completed.returncode}"
        )

    output = completed.stdout
    last_path = None
    for line in output.splitlines():
        m = re.search(r"final output blob:\s*(\S+)", line, re.IGNORECASE)
        if m:
            last_path = m.group(1)

    if not last_path:
        raise ValueError(
            f"Could not find 'final output blob:' line in output of "
            f"{os.path.basename(script_path)}"
        )

    if "/" not in last_path:
        raise ValueError(
            f"Unexpected blob path format from {os.path.basename(script_path)}: {last_path}"
        )

    container, blobname = last_path.split("/", 1)
    print(
        f"[ORCH] {os.path.basename(script_path)} produced blob: "
        f"container={container}, blob={blobname}"
    )
    return container, blobname


def main():
    print("[ORCH] Pregnancy BMI pipeline orchestration starting...")
    print(f"[ORCH] Storage account: {STORAGE_ACCOUNT_URL}")
    print(f"[ORCH] Script container: {SCRIPT_CONTAINER}")

    bsc = get_bsc(STORAGE_ACCOUNT_URL)

    run_ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    scripts_dir = os.path.join("/tmp", f"preg_bmi_scripts_{run_ts}")
    os.makedirs(scripts_dir, exist_ok=True)

    local_scripts = {}
    for local_name, blob_name in SCRIPT_BLOBS.items():
        local_path = os.path.join(scripts_dir, local_name)
        download_blob(bsc, SCRIPT_CONTAINER, blob_name, local_path)
        local_scripts[local_name] = local_path

    step1_script = local_scripts["Step1_Identify_Preg_Patients_azure.py"]
    step2_script = local_scripts["Step2_attachBMI_combined.py"]
    step3_script = local_scripts["Step3_Summarize_patient_data_azure.py"]
    step4_script = local_scripts["Step4_BMICare_tracking_azure.py"]

    step1_args = [
        "--storage-account-url", STORAGE_ACCOUNT_URL,
        "--encounters-container", INPUT_CONTAINER,
        "--encounters-blob", ENCOUNTER_BLOB,
        "--output-container", INTERMEDIATE_CONTAINER,
    ]
    step1_container, step1_blob = run_step(step1_script, step1_args)

    step2_args = [
        "--storage-account-url", STORAGE_ACCOUNT_URL,
        "--step1-container", step1_container,
        "--step1-blob", step1_blob,
        "--bmi-container", INPUT_CONTAINER,
        "--bmi-blob", BMI_BLOB,
        "--output-container", INTERMEDIATE_CONTAINER,
    ]
    step2_container, step2_blob = run_step(step2_script, step2_args)

    step3_args = [
        "--storage-account-url", STORAGE_ACCOUNT_URL,
        "--step2-container", step2_container,
        "--step2-blob", step2_blob,
        "--output-container", INTERMEDIATE_CONTAINER,
    ]
    step3_container, step3_blob = run_step(step3_script, step3_args)

    step4_args = [
        "--storage-account-url", STORAGE_ACCOUNT_URL,
        "--step3-container", step3_container,
        "--step3-blob", step3_blob,
        "--output-container", FINAL_CONTAINER,
    ]
    final_container, final_blob = run_step(step4_script, step4_args)

    print("\n[ORCH] Pipeline completed successfully.")
    print(f"[ORCH] Final Step 4 output blob: {final_container}/{final_blob}")
    print(
        f"[ORCH] You can download it from container '{FINAL_CONTAINER}' in "
        f"account '{STORAGE_ACCOUNT_URL}'."
    )
    print(f"[ORCH] Temporary scripts directory: {scripts_dir}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[ORCH] ERROR: {e}", file=sys.stderr)
        sys.exit(1)
