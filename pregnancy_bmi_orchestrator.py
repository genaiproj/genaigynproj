#!/usr/bin/env python3
# pregnancy_bmi_orchestrator.py
# Updated: 2026-04-12 for adding step7 for hidden pregnant patient
#
# Runs Step1 -> Step4 for BMI care tracking, then runs Step5 delivery tracking report, then run care gap trackking script and then run hidden pregnancy.
# After that, archives output blobs older than 4 weeks into output/archive_file/.

# Updated: 2026-04-04
# Add step6 (Care gap tracking report)
# Add step7 (Hidden Pregnancy tracking report)

import os
import re
import sys
import subprocess
from datetime import datetime, timedelta, timezone
from typing import Optional, Tuple

from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient
from azure.core.exceptions import ResourceNotFoundError


STORAGE_ACCOUNT_URL = "https://genafuncapp.blob.core.windows.net"

SCRIPT_CONTAINER = "scripts"
INPUT_CONTAINER = "input"
INTERMEDIATE_CONTAINER = "intermediate"
FINAL_CONTAINER = "output"

ARCHIVE_PREFIX = "archive_file/"
ARCHIVE_DAYS = 28

ENCOUNTER_BLOB = "EncounterData.xlsx"
BMI_BLOB = "bmi.csv"

DELIVERY_GRACE_DAYS = 7

# Put all Steps file in the scripts container using this name.
SCRIPT_BLOBS = {
    "Step1_Identify_Preg_Patients_azure.py": "Step1_Identify_Preg_Patients_azure.py",
    "Step2_attachBMI_combined.py": "Step2_attachBMI_combined.py",
    "Step3_Summarize_patient_data_azure.py": "Step3_Summarize_patient_data_azure.py",
    "Step4_BMICare_tracking_azure.py": "Step4_BMICare_tracking_azure.py",
    "Step5_Delivery_Tracking_Report.py": "Step5_Delivery_Tracking_Report.py",
    "Step6_Patient_Caregap_Tracking_Report.py": "Step6_Patient_Caregap_Tracking_Report.py",
    "Step7_Hideen_Pregnant_Patient_Report.py": "Step7_Hideen_Pregnant_Patient_Report.py",
}


def get_bsc(storage_account_url: str) -> BlobServiceClient:
    return BlobServiceClient(account_url=storage_account_url, credential=DefaultAzureCredential())


def download_blob(bsc: BlobServiceClient, container: str, blob_name: str, local_path: str) -> None:
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    print(f"[ORCH] Downloading {container}/{blob_name} -> {local_path}")
    try:
        data = bsc.get_blob_client(container, blob_name).download_blob().readall()
    except ResourceNotFoundError:
        raise FileNotFoundError(f"Blob not found: container='{container}', blob='{blob_name}'")

    with open(local_path, "wb") as f:
        f.write(data)


def run_step(script_path: str, step_args: list[str]) -> Tuple[str, str]:
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
        raise RuntimeError(f"Step script {os.path.basename(script_path)} failed with exit code {completed.returncode}")

    container, blob = parse_output_location(completed.stdout)
    if not container or not blob:
        raise ValueError(
            f"Could not find output location in logs of {os.path.basename(script_path)}. "
            "Expected a line like 'final output blob: <container>/<blob>' or 'Uploaded: container=<c> blob=<b>'."
        )

    print(f"[ORCH] {os.path.basename(script_path)} produced blob: container={container}, blob={blob}")
    return container, blob


def parse_output_location(stdout: str) -> Tuple[Optional[str], Optional[str]]:
    # Step1-4 scripts print: "Final output blob: <container>/<blob>"
    # Step5 prints: "Uploaded: container=<container> blob=<blob>"
    # Step6 prints: "Uploaded: container=<container> blob=<blob>"
    # Step7 prints: "Uploaded: container=<container> blob=<blob>"
    last_container = None
    last_blob = None

    for line in (stdout or "").splitlines():
        m1 = re.search(r"final output blob:\s*(\S+)", line, re.IGNORECASE)
        if m1:
            path = m1.group(1).strip()
            if "/" in path:
                c, b = path.split("/", 1)
                last_container, last_blob = c, b
                continue

        m2 = re.search(r"uploaded:\s*container=(\S+)\s+blob=(\S+)", line, re.IGNORECASE)
        if m2:
            last_container, last_blob = m2.group(1).strip(), m2.group(2).strip()
            continue

    return last_container, last_blob


def archive_old_output_files(
    bsc: BlobServiceClient,
    container: str,
    archive_prefix: str,
    older_than_days: int,
) -> None:
    cc = bsc.get_container_client(container)
    cutoff = datetime.now(timezone.utc) - timedelta(days=older_than_days)

    to_archive = []
    for blob in cc.list_blobs():
        name = blob.name
        if not name:
            continue
        if name.startswith(archive_prefix):
            continue
        if name.endswith("/"):
            continue
        lm = blob.last_modified
        if lm is None:
            continue
        if lm.tzinfo is None:
            lm = lm.replace(tzinfo=timezone.utc)
        if lm < cutoff:
            to_archive.append((name, lm))

    if not to_archive:
        print("[ORCH] Archive step: nothing older than 4 weeks.")
        return

    to_archive.sort(key=lambda x: x[1])
    print(f"[ORCH] Archive step: {len(to_archive)} blobs older than {older_than_days} days will be moved to {archive_prefix}")

    for name, lm in to_archive:
        dest_name = archive_prefix + name.lstrip("/")
        src = bsc.get_blob_client(container=container, blob=name)
        dst = bsc.get_blob_client(container=container, blob=dest_name)

        src_url = src.url
        copy_props = dst.start_copy_from_url(src_url)
        copy_id = copy_props.get("copy_id")

        # Poll for completion (server-side copy)
        for _ in range(60):
            props = dst.get_blob_properties()
            status = props.copy.status if props.copy else None
            if status in ("success", "failed", "aborted"):
                break
            time.sleep(2)

        props = dst.get_blob_properties()
        status = props.copy.status if props.copy else None
        if status != "success":
            print(f"[ORCH] Archive warning: copy failed for {name} -> {dest_name} (status={status}, copy_id={copy_id})")
            continue

        src.delete_blob()
        print(f"[ORCH] Archived: {name} -> {dest_name}")


def main() -> None:
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
    step5_script = local_scripts["Step5_Delivery_Tracking_Report.py"]
    step6_script = local_scripts["Step6_Patient_Caregap_Tracking_Report.py"]
    step7_script = local_scripts["Step7_Hideen_Pregnant_Patient_Report.py"]

    step1_container, step1_blob = run_step(step1_script, [
        "--storage-account-url", STORAGE_ACCOUNT_URL,
        "--encounters-container", INPUT_CONTAINER,
        "--encounters-blob", ENCOUNTER_BLOB,
        "--output-container", INTERMEDIATE_CONTAINER,
    ])

    step2_container, step2_blob = run_step(step2_script, [
        "--storage-account-url", STORAGE_ACCOUNT_URL,
        "--step1-container", step1_container,
        "--step1-blob", step1_blob,
        "--bmi-container", INPUT_CONTAINER,
        "--bmi-blob", BMI_BLOB,
        "--output-container", INTERMEDIATE_CONTAINER,
    ])

    step3_container, step3_blob = run_step(step3_script, [
        "--storage-account-url", STORAGE_ACCOUNT_URL,
        "--step2-container", step2_container,
        "--step2-blob", step2_blob,
        "--output-container", INTERMEDIATE_CONTAINER,
    ])

    final_container, final_blob = run_step(step4_script, [
        "--storage-account-url", STORAGE_ACCOUNT_URL,
        "--step3-container", step3_container,
        "--step3-blob", step3_blob,
        "--output-container", FINAL_CONTAINER,
    ])

    print(f"\n[ORCH] Step4 completed: {final_container}/{final_blob}")

    # Step5: delivery tracking report (compares latest two Step4 outputs)
    step5_container, step5_blob = run_step(step5_script, [
        "--account-name", "genafuncapp",
        "--step4-container", FINAL_CONTAINER,
        "--step4-prefix", "",
        "--encounters-container", INPUT_CONTAINER,
        "--encounters-blob", ENCOUNTER_BLOB,
        "--output-container", FINAL_CONTAINER,
        "--output-prefix", "",
        "--grace-days", str(DELIVERY_GRACE_DAYS),
    ])
    
    print(f"\n[ORCH] Step5 completed: {step5_container}/{step5_blob}")

    
    # Step6: caregap tracking report (consider latest Step4 output)
    step6_container, step6_blob = run_step(step6_script, [
        "--account-name", "genafuncapp",
        "--step4-container", FINAL_CONTAINER,
        "--step4-prefix", "",
        #"--encounters-container", INPUT_CONTAINER,
        #"--encounters-blob", ENCOUNTER_BLOB,
        "--output-container", FINAL_CONTAINER,
        "--output-prefix", "",
        #"--grace-days", str(DELIVERY_GRACE_DAYS),
    ])

    print(f"\n[ORCH] Step6 completed: {step6_container}/{step6_blob}")
    
    # Step7: hidden pregnant patient report (consider latest encounter input file and Step4 output file)
    step7_container, step7_blob = run_step(step7_script, [
        "--account-name", "genafuncapp",
        "--step4-container", FINAL_CONTAINER,
        "--step4-prefix", "",
        "--encounters-container", INPUT_CONTAINER,
        "--encounters-blob", ENCOUNTER_BLOB,
        "--output-container", FINAL_CONTAINER,
        #"--output-prefix", "",
        #"--grace-days", str(DELIVERY_GRACE_DAYS),
    ])
    
    print(f"\n[ORCH] Step7 completed: {step7_container}/{step7_blob}")

    # Archive older output files so the main output folder stays tidy
    archive_old_output_files(
        bsc=bsc,
        container=FINAL_CONTAINER,
        archive_prefix=ARCHIVE_PREFIX,
        older_than_days=ARCHIVE_DAYS,
    )

    print("\n[ORCH] Pipeline completed successfully.")
    
    print(f"[ORCH] Temporary scripts directory: {scripts_dir}")


if __name__ == "__main__":
    try:
        import time
        main()
    except Exception as e:
        print(f"[ORCH] ERROR: {e}", file=sys.stderr)
        sys.exit(1)
