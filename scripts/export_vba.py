"""Export VBA from the workbook into src/ so code changes show up in Git diffs.

Usage:
    pip install oletools
    python scripts/export_vba.py [path/to/workbook.xlsm]
"""
import re
import sys
from pathlib import Path

from oletools.olevba import VBA_Parser

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "workbook" / "XYZ_TrendWear_Order_Payment_Reports.xlsm"
SRC = ROOT / "src"


def clean(code: str) -> str:
    # Drop the hidden "Attribute VB_..." header lines and empty event stubs
    lines = [l for l in code.splitlines() if not l.startswith("Attribute VB_")]
    code = "\n".join(lines).strip() + "\n"
    return re.sub(r"Private Sub \w+\(\)\s*\n\s*End Sub\s*\n", "", code)


def main() -> None:
    parser = VBA_Parser(str(WORKBOOK))
    written = []
    for _, _, name, code in parser.extract_macros():
        body = clean(code)
        if not body.strip():
            continue  # empty sheet / ThisWorkbook modules
        if name.endswith(".frm"):
            out = SRC / "forms" / name.replace(".frm", ".code.vba")
        elif name.endswith(".cls"):
            out = SRC / "classes" / name
        else:
            out = SRC / "modules" / name
            body = f'Attribute VB_Name = "{Path(name).stem}"\n' + body
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(body, encoding="utf-8", newline="\r\n")
        written.append(out.relative_to(ROOT))
    parser.close()
    print("Exported:", *written, sep="\n  ")


if __name__ == "__main__":
    main()
