import time
import requests
import os
import subprocess

API_URL = "http://localhost:5000/api"
# Path to SumatraPDF executable on Lab PC (download from sumatrapdfreader.org)
SUMATRA_PATH = r"C:\Program Files\SumatraPDF\SumatraPDF.exe"

def check_and_print():
    try:
        response = requests.get(f"{API_URL}/agent/pending")
        if response.status_code == 200:
            jobs = response.json()
            for job in jobs:
                print(f"[PRINTNEPAL AGENT] Processing Order: {job['id']} for {job['customerName']}")

                # 1. Update status to 'printing'
                requests.post(f"{API_URL}/agent/update-status", json={"id": job['id'], "status": "printing"})

                # 2. Download PDF file
                file_url = f"http://localhost:5000{job['filePath']}"
                file_data = requests.get(file_url).content
                local_file = f"temp_{job['id']}.pdf"

                with open(local_file, "wb") as f:
                    f.write(file_data)

                # 3. Print via SumatraPDF (if installed)
                color_mode = "color" if job['colorMode'] == 'color' else "monochrome"
                duplex_mode = "duplexlong" if job['duplex'] else "simplex"
                
                if os.path.exists(SUMATRA_PATH):
                    cmd = [
                        SUMATRA_PATH,
                        "-print-to-default",
                        "-print-settings", f"{job['copies']}x,{color_mode},{duplex_mode}",
                        local_file
                    ]
                    subprocess.run(cmd, check=True)
                else:
                    print(f"[DEMO MODE] Simulated printing for {local_file}")
                    time.sleep(3) # Simulate printing delay

                # 4. Clean up temp file
                if os.path.exists(local_file):
                    os.remove(local_file)

                # 5. Update status to 'ready'
                requests.post(f"{API_URL}/agent/update-status", json={"id": job['id'], "status": "ready"})
                print(f"[PRINTNEPAL AGENT] Finished Order: {job['id']}")

    except Exception as e:
        print("[PRINTNEPAL AGENT] Error checking queue:", e)

if __name__ == "__main__":
    print("=== PrintNepal Local Lab Agent Started ===")
    while True:
        check_and_print()
        time.sleep(3)