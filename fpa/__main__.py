"""
python -m fpa build     # write the workbook (formulas only, no cached values yet)
python -m fpa export    # write the Power BI star schema from a recalculated workbook

openpyxl writes formulas without results. Excel calculates them on open: open
output/FY2026_Operating_Plan.xlsx, save it, then run `export`. The repository
ships the workbook already recalculated, with the export alongside it.
"""
import sys
from pathlib import Path

from .powerbi import export
from .workbook import build

BOOK = Path(__file__).resolve().parent.parent / "output" / "FY2026_Operating_Plan.xlsx"

if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "build"
    BOOK.parent.mkdir(exist_ok=True)
    if cmd == "build":
        build(BOOK)
        print(f"wrote {BOOK.name}. Open and save it in Excel to calculate, then run: python -m fpa export")
    elif cmd == "export":
        export(BOOK)
        print("wrote powerbi/")
