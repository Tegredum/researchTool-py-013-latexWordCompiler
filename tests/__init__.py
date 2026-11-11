# author: Codex
"""Allow source-checkout tests without installing the package globally."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
