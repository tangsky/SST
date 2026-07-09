import datetime
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUT = ROOT / "assets" / "revenue-data.js"
NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def find_workbook():
    candidates = [p for p in DATA_DIR.glob("*.xlsx") if not p.name.startswith("~$")]
    preferred = [p for p in candidates if p.stem == "mock_data"]
    if preferred:
        return preferred[0]
    if len(candidates) == 1:
        return candidates[0]
    raise SystemExit(f"Expected one workbook or data/mock_data.xlsx, got: {[p.name for p in candidates]}")


def read_xlsx(path):
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", NS):
                shared.append("".join(t.text or "" for t in si.findall(".//a:t", NS)))

        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        def col_index(cell_ref):
            n = 0
            for ch in re.match(r"([A-Z]+)", cell_ref).group(1):
                n = n * 26 + ord(ch) - 64
            return n - 1

        def cell_value(cell):
            value = cell.find("a:v", NS)
            cell_type = cell.attrib.get("t")
            if value is None:
                inline = cell.find("a:is/a:t", NS)
                return inline.text if inline is not None else ""
            text = value.text or ""
            if cell_type == "s" and text.isdigit():
                idx = int(text)
                return shared[idx] if idx < len(shared) else text
            return text

        sheets = {}
        for sheet in workbook.findall("a:sheets/a:sheet", NS):
            rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = relmap[rel_id]
            sheet_path = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
            root = ET.fromstring(z.read(sheet_path))
            rows = []
            for row in root.findall("a:sheetData/a:row", NS):
                values = []
                for cell in row.findall("a:c", NS):
                    idx = col_index(cell.attrib["r"])
                    while len(values) <= idx:
                        values.append("")
                    values[idx] = cell_value(cell)
                rows.append(values)
            sheets[sheet.attrib["name"].strip().lower()] = rows
        return sheets


def excel_date(value):
    return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(value)))


def to_float(value):
    return float(value or 0)


def to_int(value):
    return int(round(float(value or 0)))


def build_payload(path):
    sheets = read_xlsx(path)
    missing = {"revenue", "budget", "product"} - set(sheets)
    if missing:
        raise SystemExit(f"Missing required sheet(s): {sorted(missing)}")

    records = []
    for row in sheets["revenue"][1:]:
        row += [""] * 9
        if not row[0]:
            continue
        records.append(
            {
                "d": str(excel_date(row[0])),
                "m": int(float(row[1])),
                "storeCode": row[2],
                "storeName": row[3],
                "biz": row[4],
                "channel": row[5],
                "type": row[6],
                "orders": to_int(row[7]),
                "revenue": round(to_float(row[8]), 2),
            }
        )

    budgets = {}
    for row in sheets["budget"][1:]:
        row += [""] * 3
        if not row[1]:
            continue
        nums = re.findall(r"\d+", str(row[1]))
        if not nums:
            continue
        key = int(nums[0])
        if key < 100000:
            key = 202600 + key
        budgets[str(key)] = round(to_float(row[2]), 2)

    products = []
    for row in sheets["product"][1:]:
        row += [""] * 9
        if not row[0]:
            continue
        date = excel_date(row[0])
        products.append(
            {
                "d": str(date),
                "m": date.year * 100 + date.month,
                "storeCode": row[1],
                "storeName": row[2],
                "type": row[3],
                "code": row[4],
                "name": row[5],
                "category": row[6] or "未分类",
                "qty": to_int(row[7]),
                "income": round(to_float(row[8]), 2),
            }
        )

    type_counts = Counter(item["type"] for item in records)
    normal = type_counts.most_common(1)[0][0] if type_counts else "常规订单"
    return {
        "sourceFile": str(path.relative_to(ROOT)).replace("\\", "/"),
        "normalOrderType": normal,
        "records": records,
        "budgets": budgets,
        "products": products,
    }


def main():
    path = find_workbook()
    payload = build_payload(path)
    OUT.write_text(
        "window.SST_REVENUE_DATA = " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";\n",
        encoding="utf-8",
    )
    months = sorted({item["m"] for item in payload["records"]})
    product_months = sorted({item["m"] for item in payload["products"]})
    print(
        json.dumps(
            {
                "source": payload["sourceFile"],
                "output": str(OUT.relative_to(ROOT)),
                "records": len(payload["records"]),
                "products": len(payload["products"]),
                "months": months,
                "productMonths": product_months,
                "normalOrderType": payload["normalOrderType"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()