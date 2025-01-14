from pathlib import Path
import subprocess
import logging

# Adjust paths to match deployment environment
BASE_DIR = Path(__file__).resolve().parent

def handler(request):
    try:
        logging.info("Running report automation script.")
        result = subprocess.run(
            ["python3", str(BASE_DIR / "report_automation/scripts/report_automation.py")],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        return {
            "statusCode": 200,
            "body": "Script executed successfully.",
            "output": result.stdout,
        }
    except subprocess.CalledProcessError as e:
        logging.error(f"Error executing script: {e.stderr}")
        return {
            "statusCode": 500,
            "body": "Script execution failed.",
            "error": e.stderr,
        }
