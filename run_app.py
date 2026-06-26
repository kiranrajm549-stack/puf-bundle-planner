import sys
import subprocess
from pathlib import Path

def main():
    base_dir = Path(__file__).resolve().parent
    app_file = base_dir / "bundle_plan_app.py"
    subprocess.Popen([
        sys.executable, "-m", "streamlit", "run", str(app_file),
        "--server.headless=true",
        "--global.developmentMode=false"
    ])

if __name__ == "__main__":
    main()
