import json
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "inputs" / "teams.xlsx"
OUTPUT = ROOT / "certificate_data.json"


def clean(value):
    if value is None:
        return ""
    return str(value).replace("\r", "").replace("\n", "").strip()


def display_award(value):
    value = clean(value)
    if value.startswith("校") and not value.startswith("校级"):
        return "校级" + value[1:]
    return value


wb = openpyxl.load_workbook(INPUT, data_only=True)
ws = wb["Sheet1"]

records = []
for row in ws.iter_rows(min_row=2, values_only=True):
    award, work_title, team_name, members, teacher = row[:5]
    if not any(clean(v) for v in (award, work_title, team_name, members, teacher)):
        continue
    records.append(
        {
            "award": clean(award),
            "award_display": display_award(award),
            "work_title": clean(work_title),
            "team_name": clean(team_name),
            "members": clean(members),
            "teacher": clean(teacher),
        }
    )

payload = {
    "issuer": "徐州医科大学",
    "event": "“正大杯”第十六届校级选拔赛",
    "date": "2026年6月23日",
    "records": records,
}

OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"Wrote {len(records)} records to {OUTPUT}")
