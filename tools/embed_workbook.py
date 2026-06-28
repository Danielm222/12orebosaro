"""Aggiorna il fallback Excel usato dalle pagine aperte direttamente da disco."""

from base64 import b64encode
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = ROOT / "12HBosaro_manager.xlsx"
OUTPUT = ROOT / "assets" / "tournament-workbook.js"


def main() -> None:
    encoded = b64encode(WORKBOOK.read_bytes()).decode("ascii")
    OUTPUT.write_text(
        f"window.TOURNAMENT_WORKBOOK_BASE64 = '{encoded}';\n",
        encoding="utf-8",
    )
    print(f"Aggiornato {OUTPUT.relative_to(ROOT)} da {WORKBOOK.name}")


if __name__ == "__main__":
    main()
