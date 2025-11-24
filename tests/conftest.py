import sys
from pathlib import Path

# Ensure src package root is importable when running tests without installing
ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))
