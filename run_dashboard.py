"""Start the Access Review Dashboard Streamlit server."""

import subprocess
import sys
from pathlib import Path

if __name__ == "__main__":
    dashboard = Path(__file__).parent / "dashboard.py"
    subprocess.run(
        [sys.executable, "-m", "streamlit", "run", str(dashboard)],
        check=True,
    )
