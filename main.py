"""RESTful API: สถิติจำนวนผู้เข้าชมสวนสัตว์ ปีงบประมาณ 2569 (องค์การสวนสัตว์แห่งประเทศไทย)
ข้อมูลจริงจาก Data.go.th (resource id 0ecc96cf-d914-42a1-b184-dbbd013c3c9c)
อ่านจากไฟล์ทรัพยากรจริง data/zoo-visitors2569.xlsx (ไม่มีข้อมูลสมมติ)
"""
from functools import lru_cache
from pathlib import Path

import openpyxl
from fastapi import FastAPI, HTTPException, Query

XLSX_PATH = Path(__file__).parent / "data" / "zoo-visitors2569.xlsx"

SOURCE = {
    "publisher": "องค์การสวนสัตว์แห่งประเทศไทย ในพระบรมราชูปถัมภ์",
    "dataset": "zoo-visitors2569",
    "title": "รายงานสถิติจำนวนผู้เข้าชมสวนสัตว์ ปีงบประมาณ 2569",
    "portal": "Data.go.th",
    "resource_id": "0ecc96cf-d914-42a1-b184-dbbd013c3c9c",
    "file": XLSX_PATH.name,
}
FISCAL_YEAR = 2569

app = FastAPI(
    title="Zoo Visitors API (FY2569)",
    description="สถิติผู้เข้าชมสวนสัตว์ ปีงบประมาณ 2569 จาก Data.go.th ขององค์การสวนสัตว์แห่งประเทศไทย",
    version="1.0.0",
)


@lru_cache
def load():
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True, read_only=True)
    rows = list(wb["statistics"].iter_rows(values_only=True))
    months = [m for m in rows[0][3:] if m]
    records = []
    for r in rows[1:]:
        if r[0] is None:
            continue
        monthly = {m: int(v or 0) for m, v in zip(months, r[3:3 + len(months)])}
        records.append({
            "category": r[0],
            "zoo": r[1],
            "fiscal_year": r[2],
            "monthly": monthly,
            "total": sum(monthly.values()),
        })
    dictionary = {}
    if "Data Dictionary" in wb.sheetnames:
        for k, v in list(wb["Data Dictionary"].iter_rows(values_only=True))[1:]:
            if k:
                dictionary[k] = v
    return months, records, dictionary


def envelope(data, **extra):
    return {"source": SOURCE, "fiscal_year": FISCAL_YEAR, **extra, "data": data}


@app.get("/")
def root():
    return {"message": "Zoo Visitors API FY2569", "docs": "/docs", "endpoint": "/api/zoo-visitors"}


@app.get("/api/zoo-visitors")
def zoo_visitors(
    zoo: str | None = Query(None, description="กรองชื่อสวนสัตว์ (ตรงบางส่วน) เช่น เชียงใหม่"),
    category: str | None = Query(None, description="กรองประเภทผู้เข้าชม (ตรงบางส่วน) เช่น ผู้ใหญ่"),
):
    months, records, dictionary = load()
    out = [r for r in records
           if (not zoo or zoo in r["zoo"]) and (not category or category in r["category"])]
    if not out:
        raise HTTPException(404, "ไม่พบข้อมูลตามเงื่อนไขที่ระบุ")
    return envelope(out, months=months, count=len(out), category_descriptions=dictionary)


@app.get("/api/zoo-visitors/zoos")
def zoos():
    _, records, _ = load()
    names = list(dict.fromkeys(r["zoo"] for r in records))
    return envelope(names, count=len(names))


@app.get("/api/zoo-visitors/summary")
def summary():
    months, records, _ = load()
    by_zoo: dict[str, int] = {}
    for r in records:
        by_zoo[r["zoo"]] = by_zoo.get(r["zoo"], 0) + r["total"]
    by_month = {m: sum(r["monthly"][m] for r in records) for m in months}
    return envelope({
        "grand_total": sum(by_zoo.values()),
        "total_by_zoo": by_zoo,
        "total_by_month": by_month,
    })
