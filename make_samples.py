"""Generate synthetic sample data for the reconcile demo.

All names are obviously fake ("Sample Co 001"). No randomness: every
difference is planted on purpose so the expected report is known.

System A is a clean export. System B is the same ledger re-keyed by hand:
  * cosmetic noise that must NOT count as a difference:
    extra spaces, different case, "Limited" vs "Ltd", "1,250.00" vs 1250,
    "03/02/2026" (d/m/y) vs 2026-02-03
  * real differences: 3 amounts, 2 statuses, 1 date
  * 2 rows only in A, 2 rows only in B
  * 2 rows whose id was mistyped in B (fuzzy name match should pair them)
  * 1 decoy: "Sample Co 071" exists only in B and must NOT fuzzy-match
    "Sample Co 017" (numbers in names have to agree)
Outputs (in samples/): system_a.xlsx, system_b.xlsx, system_a.csv,
system_b.csv and two_sheets.xlsx (sheets "System A" and "System B").
"""
from datetime import date
from pathlib import Path
import sys

import pandas as pd

HERE = Path(__file__).resolve().parent
STATUSES = ("Paid", "Open", "Overdue")


def system_a() -> pd.DataFrame:
    rows = []
    for n in range(1, 31):
        rows.append({
            "id": 1000 + n,
            "name": f"Sample Co {n:03d} Limited",
            "amount": round(100 + n * 37.25, 2),
            "invoice_date": date(2026, 1 + n % 6, 1 + (n * 3) % 27),
            "status": STATUSES[n % 3],
        })
    return pd.DataFrame(rows)


def system_b(a: pd.DataFrame) -> pd.DataFrame:
    b = a.copy().astype(object)
    # ---- cosmetic noise (should be normalised away)
    for i in range(len(b)):
        n = i + 1
        if n % 4 == 0:
            b.at[i, "name"] = f"  sample co {n:03d}   ltd "
        if n % 5 == 0:
            b.at[i, "amount"] = f"{a.at[i, 'amount']:,.2f}"          # "1,250.00" as text
        if n % 3 == 0:
            d = a.at[i, "invoice_date"]
            b.at[i, "invoice_date"] = d.strftime("%d/%m/%Y")          # d/m/y text
        if n % 7 == 0:
            b.at[i, "status"] = str(a.at[i, "status"]).upper() + " "
    # ---- real differences
    idx = {int(r): i for i, r in enumerate(b["id"])}
    b.at[idx[1003], "amount"] = 250.00            # was 211.75
    b.at[idx[1011], "amount"] = 520.00            # was 509.75
    b.at[idx[1022], "amount"] = 920.00            # was 919.50
    b.at[idx[1008], "status"] = "Paid"            # was Overdue
    b.at[idx[1013], "status"] = "Overdue"         # was Open
    b.at[idx[1019], "invoice_date"] = "2026-02-25"  # was 2026-02-04
    # ---- mistyped ids in B (names still match closely)
    b.at[idx[1017], "id"] = 1071                  # typo; name "Sample Co 017 Limited"
    b.at[idx[1026], "id"] = "1O26"                # letter O instead of zero
    b.at[idx[1026], "name"] = "Sample Co. 026 Ltd"
    # ---- rows only in A
    b = b.drop(index=[idx[1005], idx[1029]])
    # ---- rows only in B (one is the decoy)
    extra = pd.DataFrame([
        {"id": 1101, "name": "Sample Co 071 Limited", "amount": 300.00,
         "invoice_date": "2026-03-10", "status": "Open"},
        {"id": 1102, "name": "Sample Co 099 Limited", "amount": 150.00,
         "invoice_date": "2026-03-11", "status": "Paid"},
    ])
    return pd.concat([b, extra], ignore_index=True)


def build(out_dir: Path) -> dict[str, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    a = system_a()
    b = system_b(a)
    paths = {
        "a_xlsx": out_dir / "system_a.xlsx", "b_xlsx": out_dir / "system_b.xlsx",
        "a_csv": out_dir / "system_a.csv", "b_csv": out_dir / "system_b.csv",
        "two_sheets": out_dir / "two_sheets.xlsx",
    }
    a.to_excel(paths["a_xlsx"], index=False, sheet_name="System A")
    b.to_excel(paths["b_xlsx"], index=False, sheet_name="System B")
    a.to_csv(paths["a_csv"], index=False)
    b.to_csv(paths["b_csv"], index=False)
    with pd.ExcelWriter(paths["two_sheets"]) as xw:
        a.to_excel(xw, index=False, sheet_name="System A")
        b.to_excel(xw, index=False, sheet_name="System B")
    return paths


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "samples"
    for k, v in build(target).items():
        print(k, v.name)
