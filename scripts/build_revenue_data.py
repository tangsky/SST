import datetime
import hashlib
import math
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


def excel_date(value, sheet, row_number):
    try:
        serial = float(value)
    except (TypeError, ValueError):
        raise SystemExit(f"{sheet} row {row_number}: 日期不是有效的 Excel 日期值: {value!r}")
    if not math.isfinite(serial):
        raise SystemExit(f"{sheet} row {row_number}: 日期不是有限数值")
    result = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(serial))
    if not 2000 <= result.year <= 2100:
        raise SystemExit(f"{sheet} row {row_number}: 日期超出支持范围: {result}")
    return result


def excel_time(value, sheet, row_number):
    if value is None or str(value).strip() == "":
        return "00:00", 0
    try:
        raw = float(value)
    except (TypeError, ValueError):
        text = str(value).strip()
        match = re.search(r"(\d{1,2}):(\d{2})", text)
        if not match:
            raise SystemExit(f"{sheet} row {row_number}: invalid half-hour time: {value!r}")
        hour, minute = int(match.group(1)), int(match.group(2))
    else:
        if raw >= 1:
            raw = raw % 1
        minutes = int(round(raw * 24 * 60)) % (24 * 60)
        hour, minute = divmod(minutes, 60)
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        raise SystemExit(f"{sheet} row {row_number}: half-hour out of 00:00-23:59 range: {value!r}")
    slot = hour * 60 + minute
    return f"{hour:02d}:{minute:02d}", slot


def _number(value, sheet, row_number, field):
    if value is None or str(value).strip() == "":
        return 0.0
    try:
        number = float(value)
    except (TypeError, ValueError):
        raise SystemExit(f"{sheet} row {row_number}: {field} 不是有效数字: {value!r}")
    if not math.isfinite(number):
        raise SystemExit(f"{sheet} row {row_number}: {field} 不是有限数值")
    return number


def to_float(value, sheet, row_number, field):
    return _number(value, sheet, row_number, field)


def to_int(value, sheet, row_number, field):
    number = _number(value, sheet, row_number, field)
    if number < 0:
        raise SystemExit(f"{sheet} row {row_number}: {field} 不能为负数")
    return int(round(number))


def to_signed_int(value, sheet, row_number, field):
    number = _number(value, sheet, row_number, field)
    return int(round(number))


def parse_month_key(value, default_year, sheet, row_number):
    nums = [int(item) for item in re.findall(r"\d+", str(value))]
    if not nums:
        raise SystemExit(f"{sheet} row {row_number}: 月份不是有效值: {value!r}")
    if len(nums) >= 2 and 1900 <= nums[0] <= 2100:
        year, month = nums[0], nums[1]
    elif 100000 <= nums[0] <= 999999:
        year, month = nums[0] // 100, nums[0] % 100
    elif 1 <= nums[0] <= 12:
        year, month = default_year, nums[0]
    else:
        raise SystemExit(f"{sheet} row {row_number}: 月份超出支持范围: {value!r}")
    if not 1 <= month <= 12:
        raise SystemExit(f"{sheet} row {row_number}: 月份超出 1-12 范围: {value!r}")
    return year * 100 + month

def build_payload(path):
    sheets = read_xlsx(path)
    missing = {"revenue", "budget", "product"} - set(sheets)
    if missing:
        raise SystemExit(f"Missing required sheet(s): {sorted(missing)}")

    records = []
    for row_number, raw_row in enumerate(sheets["revenue"][1:], start=2):
        row = list(raw_row) + [""] * 9
        if not str(row[0]).strip():
            continue
        date = excel_date(row[0], "revenue", row_number)
        month = to_int(row[1], "revenue", row_number, "月份")
        expected_month = date.year * 100 + date.month
        if month != expected_month:
            raise SystemExit(f"revenue row {row_number}: 月份 {month} 与日期 {date} 不一致")
        records.append(
            {
                "d": str(date),
                "m": month,
                "storeCode": str(row[2]).strip(),
                "storeName": str(row[3]).strip(),
                "biz": str(row[4]).strip(),
                "channel": str(row[5]).strip(),
                "type": str(row[6]).strip(),
                "orders": to_int(row[7], "revenue", row_number, "订单数"),
                "revenue": round(to_float(row[8], "revenue", row_number, "营收"), 2),
            }
        )

    default_year = max((int(item["d"][:4]) for item in records), default=datetime.date.today().year)
    budgets = {}
    for row_number, raw_row in enumerate(sheets["budget"][1:], start=2):
        row = list(raw_row) + [""] * 3
        if not str(row[1]).strip():
            continue
        key = parse_month_key(row[1], default_year, "budget", row_number)
        budgets[str(key)] = round(to_float(row[2], "budget", row_number, "月预算"), 2)

    products = []
    for row_number, raw_row in enumerate(sheets["product"][1:], start=2):
        row = list(raw_row) + [""] * 9
        if not str(row[0]).strip():
            continue
        date = excel_date(row[0], "product", row_number)
        products.append(
            {
                "d": str(date),
                "m": date.year * 100 + date.month,
                "storeCode": str(row[1]).strip(),
                "storeName": str(row[2]).strip(),
                "type": str(row[3]).strip(),
                "code": str(row[4]).strip(),
                "name": str(row[5]).strip(),
                "category": str(row[6]).strip() or "未分类",
                "qty": to_signed_int(row[7], "product", row_number, "销量"),
                "income": round(to_float(row[8], "product", row_number, "收入"), 2),
            }
        )

    daily_records = []
    for row_number, raw_row in enumerate(sheets.get("daily_revenue", [])[1:], start=2):
        row = list(raw_row) + [""] * 9
        if not str(row[0]).strip():
            continue
        date = excel_date(row[0], "daily_revenue", row_number)
        time_label, slot = excel_time(row[1], "daily_revenue", row_number)
        daily_records.append(
            {
                "d": str(date),
                "m": date.year * 100 + date.month,
                "time": time_label,
                "slot": slot,
                "period": str(row[2]).strip(),
                "storeCode": str(row[3]).strip(),
                "storeName": str(row[4]).strip(),
                "biz": str(row[5]).strip(),
                "channel": str(row[6]).strip(),
                "orders": to_int(row[7], "daily_revenue", row_number, "orders"),
                "revenue": round(to_float(row[8], "daily_revenue", row_number, "revenue"), 2),
            }
        )

    daily_products = []
    for row_number, raw_row in enumerate(sheets.get("daily_product", [])[1:], start=2):
        row = list(raw_row) + [""] * 10
        if not str(row[0]).strip():
            continue
        date = excel_date(row[0], "daily_product", row_number)
        time_label, slot = excel_time(row[3], "daily_product", row_number)
        daily_products.append(
            {
                "d": str(date),
                "m": date.year * 100 + date.month,
                "storeCode": str(row[1]).strip(),
                "storeName": str(row[2]).strip(),
                "time": time_label,
                "slot": slot,
                "period": str(row[4]).strip(),
                "code": str(row[5]).strip(),
                "name": str(row[6]).strip(),
                "category": str(row[7]).strip() or "orders",
                "qty": to_signed_int(row[8], "daily_product", row_number, "orders"),
                "income": round(to_float(row[9], "daily_product", row_number, "revenue"), 2),
            }
        )

    type_counts = Counter(item["type"] for item in records if item["type"])
    normal = type_counts.most_common(1)[0][0] if type_counts else "常规订单"
    months = sorted({item["m"] for item in records})
    max_date = max((item["d"] for item in records), default=None)
    return {
        "sourceFile": str(path.relative_to(ROOT)).replace("\\", "/"),
        "normalOrderType": normal,
        "availableMonths": months,
        "maxDataDate": max_date,
        "records": records,
        "budgets": budgets,
        "products": products,
        "availableDailyDates": sorted({item["d"] for item in daily_records}),
        "maxDailyDate": max((item["d"] for item in daily_records), default=None),
        "dailyRecords": daily_records,
        "dailyProducts": daily_products,
    }


def main():
    path = find_workbook()
    payload = build_payload(path)
    OUT.write_text(
        "window.SST_REVENUE_DATA = " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";\n",
        encoding="utf-8",
    )
    version = hashlib.sha1(OUT.read_bytes()).hexdigest()[:12]
    for page_name in ("annual.html", "monthly.html", "custom.html", "daily.html"):
        page_path = ROOT / page_name
        page_text = page_path.read_text(encoding="utf-8")
        updated_page, replacements = re.subn(
            r'(assets/revenue-data\.js\?v=)[^"\']+',
            rf"\g<1>{version}",
            page_text,
            count=1,
        )
        if replacements != 1:
            raise SystemExit(f"Could not find the revenue-data.js cache version in {page_name}")
        if updated_page != page_text:
            page_path.write_text(updated_page, encoding="utf-8")
    months = payload["availableMonths"]
    print(
        json.dumps(
            {
                "source": payload["sourceFile"],
                "output": str(OUT.relative_to(ROOT)),
                "records": len(payload["records"]),
                "products": len(payload["products"]),
                "dailyRecords": len(payload["dailyRecords"]),
                "dailyProducts": len(payload["dailyProducts"]),
                "dailyDates": payload["availableDailyDates"],
                "months": months,
                "maxDataDate": payload["maxDataDate"],
                "normalOrderType": payload["normalOrderType"],
                "cacheVersion": version,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()