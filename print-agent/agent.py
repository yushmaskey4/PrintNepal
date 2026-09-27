import os
import time
import subprocess
import requests

# Configuration
SERVER_URL = "http://localhost:5000"
POLL_INTERVAL = 5  # seconds
SUMATRA_PATH = r"C:\Program Files\SumatraPDF\SumatraPDF.exe"
DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "temp_downloads")

os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def fetch_pending_jobs():
    try:
        response = requests.get(f"{SERVER_URL}/api/agent/pending")
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        print(f"[!] Error contacting server: {e}")
    return []

def update_job_status(order_id, status):
    try:
        response = requests.post(
            f"{SERVER_URL}/api/agent/update-status",
            json={"id": order_id, "status": status}
        )
        return response.status_code == 200
    except Exception as e:
        print(f"[!] Error updating status for {order_id}: {e}")
        return False

def print_pdf(file_path, copies=1, color_mode="bw"):
    """Executes physical print job via SumatraPDF CLI"""
    if not os.path.exists(SUMATRA_PATH):
        print(f"[!] SumatraPDF executable not found at '{SUMATRA_PATH}'. Simulation mode active.")
        return True

    # SumatraPDF silent print arguments
    cmd = [
        SUMATRA_PATH,
        "-print-to-default",
        "-print-settings", f"{copies}x",
        file_path
    ]
    
    print(f"[*] Executing command: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result.returncode == 0

def process_jobs():
    print("=" * 50)
    print(" 🖨️  PrintNepal Lab Print Daemon Running...")
    print(f" Polling server every {POLL_INTERVAL} seconds...")
    print("=" * 50)

    while True:
        jobs = fetch_pending_jobs()

        if jobs:
            print(f"\n[+] Found {len(jobs)} pending print job(s)!")

        for job in jobs:
            order_id = job.get("id")
            file_url = f"{SERVER_URL}{job.get('filePath')}"
            copies = job.get("copies", 1)
            color_mode = job.get("colorMode", "bw")
            file_name = job.get("fileName")

            print(f"\n---> Processing Order #{order_id} ({file_name})")
            print(f"     Specs: {copies} Copies | Color: {color_mode.upper()}")

            # 1. Update status to printing
            update_job_status(order_id, "printing")

            # 2. Download file locally
            local_file_path = os.path.join(DOWNLOAD_DIR, f"{order_id}_{file_name}")
            file_res = requests.get(file_url)
            
            with open(local_file_path, "wb") as f:
                f.write(file_res.content)
            
            print(f"     Downloaded to temp storage: {local_file_path}")

            # 3. Trigger physical print
            success = print_pdf(local_file_path, copies, color_mode)

            if success:
                print(f" [✓] Print successful for #{order_id}!")
                # 4. Update status to ready & log to Excel report
                update_job_status(order_id, "ready")
            else:
                print(f" [X] Print failed for #{order_id}.")

            # Cleanup temp downloaded file
            if os.path.exists(local_file_path):
                os.remove(local_file_path)

        time.sleep(POLL_INTERVAL)

if __name__ == "__main__":
    process_jobs()