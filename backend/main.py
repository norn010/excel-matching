# -*- coding: utf-8 -*-
"""
API สำหรับเทียบข้อมูล Excel 2 ไฟล์
- เทียบ 7 คอลัมน์: ชื่อลูกค้า, ไฟแนนซ์, Moder Code, เลขถัง, ราคาขาย, COM F/N, COM
- ใช้ "เลขถัง" เป็น key หลัก (match_key_col = 3)
"""

from __future__ import annotations

import shutil
import tempfile
import time
import re
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from excel_reader import (
    COMPARE_COLUMN_LABELS,
    detect_compare_layout,
    get_sheet_names,
    read_esg_rows,
    read_tax_sheet_rows,
)


def _safe_unlink(path: str | None) -> None:
    if not path:
        return
    p = Path(path)
    for _ in range(3):
        try:
            p.unlink(missing_ok=True)
            return
        except PermissionError:
            time.sleep(0.1)
        except Exception:
            return


def _norm(v: str) -> str:
    return (v or "").strip()


def _norm_tank(v: str) -> str:
    # normalize เลขถัง: ตัดช่องว่าง/ขีด/อักขระพิเศษ และ upper-case
    return re.sub(r"[^A-Za-z0-9]", "", _norm(v)).upper()


def _cell_match(esg_val: str, tax_val: str, col_index: int, case_sensitive: bool = True) -> tuple[bool, bool]:
    """
    return (is_match, is_compared)
    - ESG ว่าง: ไม่เอามาคิดผลรวม
    - Moder Code: strict compare แบบ normalize whitespace + lowercase
    - ตัวเลข: compare เชิงตัวเลขได้
    """
    a = _norm(esg_val)
    b = _norm(tax_val)
    if a == "":
        return True, False

    if col_index == 2:
        la = " ".join(a.lower().split())
        lb = " ".join(b.lower().split())
        return la == lb, True

    # เลขถัง: เทียบแบบ normalize เสมอ เพื่อให้ match ได้แม้รูปแบบต่างกัน
    if col_index == 3:
        return _norm_tank(a) == _norm_tank(b), True

    try:
        na = float(a.replace(",", ""))
        nb = float(b.replace(",", ""))
        return abs(na - nb) < 1e-9, True
    except Exception:
        pass

    if not case_sensitive:
        return a.lower() == b.lower(), True
    return a == b, True


def _match_rows(
    esg_rows: list[list[str]],
    tax_rows: list[list[str]],
    match_key_col: int = 3,
    case_sensitive: bool = True,
    compare_indexes: list[int] | None = None,
) -> list[dict]:
    if compare_indexes is None:
        compare_indexes = [3, 4, 5, 6]  # เลขถัง, ราคาขาย, COM F/N, COM
    # lookup โดย key หลัก + key รอง (เลขถัง, ชื่อลูกค้า, ไฟแนนซ์)
    max_cols = max([len(r) for r in tax_rows], default=0)
    lookups: list[dict] = [{} for _ in range(max_cols)]

    def k(ci: int, v: str) -> str:
        if ci == 3:
            return _norm_tank(v)
        x = _norm(v)
        return x if case_sensitive else x.lower()

    for row in tax_rows:
        for ci in range(min(len(row), max_cols)):
            key = k(ci, row[ci])
            if not key:
                continue
            lookups[ci].setdefault(key, []).append(row)

    # ถ้าใช้เลขถังเป็น key หลัก ให้ strict เฉพาะ key นี้
    # (ลดโอกาสไปแมตช์คนอื่นจากคีย์รองเวลาเทสแก้ข้อมูล)
    fallback_keys = [] if match_key_col == 3 else [3, 0, 1]  # tank -> customer -> finance
    results: list[dict] = []

    for idx, esg_row in enumerate(esg_rows):
        candidates = []
        seen = set()

        def add_by_col(ci: int):
            if ci < 0 or ci >= len(esg_row) or ci >= len(lookups):
                return
            key = k(ci, esg_row[ci])
            if not key:
                return
            for cand in lookups[ci].get(key, []):
                cid = id(cand)
                if cid in seen:
                    continue
                seen.add(cid)
                candidates.append(cand)

        add_by_col(match_key_col)
        for fk in fallback_keys:
            if fk != match_key_col:
                add_by_col(fk)

        if not candidates:
            compared = [(i in compare_indexes) and bool(_norm(esg_row[i])) for i in range(len(esg_row))]
            results.append(
                {
                    "row": idx + 1,
                    "values_esg": esg_row,
                    "values_tax": [""] * len(esg_row),
                    # ไม่มีแถวอ้างอิง => คอลัมน์ที่ ESG มีค่าถือว่า "ไม่ตรง"
                    "cells_match": [False if compared[i] else True for i in range(len(esg_row))],
                    "cells_compared": compared,
                    "all_match": False,
                    "found": False,
                }
            )
            continue

        def score(cand: list[str]) -> tuple[int, int]:
            vals = list(cand) + [""] * (len(esg_row) - len(cand))
            vals = vals[: len(esg_row)]
            m, c = 0, 0
            for i in compare_indexes:
                if i >= len(esg_row):
                    continue
                ok, compared = _cell_match(esg_row[i], vals[i], i, case_sensitive=case_sensitive)
                if compared:
                    c += 1
                    if ok:
                        m += 1
            return m, c

        best = max(candidates, key=score)
        tax_vals = list(best) + [""] * (len(esg_row) - len(best))
        tax_vals = tax_vals[: len(esg_row)]

        cells_match, cells_compared = [], []
        for i in range(len(esg_row)):
            if i in compare_indexes:
                ok, compared = _cell_match(esg_row[i], tax_vals[i], i, case_sensitive=case_sensitive)
            else:
                ok, compared = True, False
            cells_match.append(ok)
            cells_compared.append(compared)
        compared_idx = [i for i, x in enumerate(cells_compared) if x]
        all_match = bool(compared_idx) and all(cells_match[i] for i in compared_idx)

        results.append(
            {
                "row": idx + 1,
                "values_esg": esg_row,
                "values_tax": tax_vals,
                "cells_match": cells_match,
                "cells_compared": cells_compared,
                "all_match": all_match,
                "found": True,
            }
        )
    return results


app = FastAPI(title="Excel Matching API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TAX_SHEET_NAME = "ตารางไมวัน"

ROOT_DIR = Path(__file__).resolve().parent.parent
FRONTEND_DIR = ROOT_DIR / "frontend"
INDEX_HTML = FRONTEND_DIR / "index.html"


@app.get("/")
def root():
    if INDEX_HTML.exists():
        return RedirectResponse("/app/")
    return {"message": "Excel Matching API"}


@app.get("/app")
@app.get("/app/")
def serve_app():
    if not INDEX_HTML.exists():
        raise HTTPException(404, "Frontend not found")
    return FileResponse(INDEX_HTML)


if FRONTEND_DIR.exists():
    app.mount("/app", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="app")


@app.get("/sheets")
async def list_sheets(file: UploadFile = File(...)):
    suffix = Path(file.filename or "").suffix.lower()
    if suffix not in (".xls", ".xlsx", ".xlsm"):
        raise HTTPException(400, "รองรับเฉพาะ .xls/.xlsx")
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            shutil.copyfileobj(file.file, tmp)
            tmp_path = tmp.name
        return {"sheet_names": get_sheet_names(tmp_path)}
    finally:
        _safe_unlink(tmp_path)


@app.post("/match-columns")
async def match_columns(
    esg_file: UploadFile = File(...),
    tax_file: UploadFile = File(...),
    esg_sheet_name: str = Query("", description="ชื่อชีต ESG (ว่าง = auto)"),
    sheet_name: str = Query(TAX_SHEET_NAME, description="ชื่อชีตภาษีขาย"),
    esg_cols: str = Query("", description="คอลัมน์ ESG 7 ตัว (csv)"),
    tax_cols: str = Query("", description="คอลัมน์ TAX 7 ตัว (csv)"),
    match_key_col: int = Query(3, description="0=ชื่อลูกค้า, 1=ไฟแนนซ์, 3=เลขถัง"),
    case_sensitive: bool = Query(True),
):
    started_at = time.perf_counter()
    esg_suffix = Path(esg_file.filename or "").suffix.lower()
    tax_suffix = Path(tax_file.filename or "").suffix.lower()
    if esg_suffix not in (".xls", ".xlsx", ".xlsm") or tax_suffix not in (".xls", ".xlsx", ".xlsm"):
        raise HTTPException(400, "ทั้งสองไฟล์ต้องเป็น .xls/.xlsx")

    def parse_cols(raw: str) -> list[int]:
        raw = (raw or "").strip()
        if not raw:
            return []
        return [int(x.strip()) for x in raw.split(",") if x.strip()]

    esg_col_indices = parse_cols(esg_cols)
    tax_col_indices = parse_cols(tax_cols)
    if esg_col_indices and len(esg_col_indices) != 7:
        raise HTTPException(400, "esg_cols ต้องมี 7 ตัว")
    if tax_col_indices and len(tax_col_indices) != 7:
        raise HTTPException(400, "tax_cols ต้องมี 7 ตัว")

    tmp_esg = None
    tmp_tax = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=esg_suffix) as tmp:
            shutil.copyfileobj(esg_file.file, tmp)
            tmp_esg = tmp.name
        with tempfile.NamedTemporaryFile(delete=False, suffix=tax_suffix) as tmp:
            shutil.copyfileobj(tax_file.file, tmp)
            tmp_tax = tmp.name

        # auto detect เมื่อไม่ส่ง cols
        if not esg_col_indices:
            esg_col_indices, esg_start_row, detected_esg_sheet, _ = detect_compare_layout(
                tmp_esg,
                sheet_name=(esg_sheet_name.strip() or None),
                fallback_cols=[2, 3, 5, 6, 8, 14, 15],
            )
        else:
            esg_start_row = 6
            detected_esg_sheet = (esg_sheet_name.strip() or None)

        if not tax_col_indices:
            tax_col_indices, tax_start_row, detected_tax_sheet, tax_scores = detect_compare_layout(
                tmp_tax,
                sheet_name=(sheet_name.strip() or None),
                fallback_cols=[1, 2, 4, 5, 7, 13, 14],
            )
            # ถ้าอ่านไฟล์ที่โครงคนละแบบมากๆ ให้แจ้งแทนการเทียบมั่ว
            # ตอนนี้เทียบเฉพาะ: เลขถัง, ราคาขาย, COM F/N, COM
            required = ["tank", "price", "com_fn", "com"]
            missing = [f for f in required if tax_scores.get(f, 0) <= 0]
            if missing:
                raise HTTPException(
                    400,
                    f"ไม่พบหัวคอลัมน์สำคัญในไฟล์ภาษีขาย: {missing}. กรุณาระบุ tax_cols เอง",
                )
        else:
            tax_start_row = 0
            detected_tax_sheet = (sheet_name.strip() or None)

        esg_rows = read_esg_rows(
            tmp_esg,
            start_row=esg_start_row,
            col_indices=esg_col_indices,
            sheet_name=detected_esg_sheet,
        )
        # ธุรกรรมจริง: ต้องมี customer + tank + key
        esg_rows = [
            r
            for r in esg_rows
            if len(r) > 3
            and _norm(r[0])
            and _norm(r[3])
            and _norm(r[3]) != "-"
            and (match_key_col < len(r) and _norm(r[match_key_col]))
        ]

        tax_rows = read_tax_sheet_rows(
            tmp_tax,
            sheet_name=detected_tax_sheet or sheet_name,
            col_indices=tax_col_indices,
            start_row=tax_start_row,
        )

        # เทียบเฉพาะ: เลขถัง, ราคาขาย, COM F/N, COM
        results = _match_rows(
            esg_rows,
            tax_rows,
            match_key_col=match_key_col,
            case_sensitive=case_sensitive,
            compare_indexes=[3, 4, 5, 6],
        )

        # ถ้ารอบแรกหาแถวไม่เจอเกือบทั้งหมด ให้ auto-recover ด้วยการ detect layout ใหม่
        found_count = sum(1 for x in results if x.get("found"))
        if found_count == 0:
            auto_tax_cols, auto_tax_start, auto_tax_sheet, auto_tax_scores = detect_compare_layout(
                tmp_tax,
                sheet_name=None,
                fallback_cols=[1, 2, 4, 5, 7, 13, 14],
            )
            required = ["tank", "price", "com_fn", "com"]
            missing = [f for f in required if auto_tax_scores.get(f, 0) <= 0]
            if not missing:
                retry_tax_rows = read_tax_sheet_rows(
                    tmp_tax,
                    sheet_name=auto_tax_sheet,
                    col_indices=auto_tax_cols,
                    start_row=auto_tax_start,
                )
                retry_results = _match_rows(
                    esg_rows,
                    retry_tax_rows,
                    match_key_col=match_key_col,
                    case_sensitive=case_sensitive,
                    compare_indexes=[3, 4, 5, 6],
                )
                retry_found = sum(1 for x in retry_results if x.get("found"))
                if retry_found > found_count:
                    results = retry_results
                    tax_rows = retry_tax_rows
                    tax_col_indices = auto_tax_cols
                    tax_start_row = auto_tax_start
                    detected_tax_sheet = auto_tax_sheet

        for r in results:
            mismatches = []
            compared = r.get("cells_compared", [True] * len(r.get("cells_match", [])))
            for i, ok in enumerate(r.get("cells_match", [])):
                if i < len(compared) and not compared[i]:
                    continue
                if not ok:
                    mismatches.append(
                        {
                            "column_label": COMPARE_COLUMN_LABELS[i] if i < len(COMPARE_COLUMN_LABELS) else f"คอลัมน์{i}",
                            "esg_value": r["values_esg"][i] if i < len(r["values_esg"]) else "",
                            "tax_value": r["values_tax"][i] if i < len(r["values_tax"]) else "",
                        }
                    )
            r["mismatches"] = mismatches

        all_match_count = sum(1 for x in results if x.get("all_match"))
        elapsed_ms = round((time.perf_counter() - started_at) * 1000, 2)
        return JSONResponse(
            {
                "column_labels": COMPARE_COLUMN_LABELS,
                "match_key_label": COMPARE_COLUMN_LABELS[match_key_col] if match_key_col < len(COMPARE_COLUMN_LABELS) else "คอลัมน์จับคู่",
                "elapsed_ms": elapsed_ms,
                "detected": {
                    "esg_sheet": detected_esg_sheet,
                    "tax_sheet": detected_tax_sheet,
                    "esg_cols": esg_col_indices,
                    "tax_cols": tax_col_indices,
                },
                "summary": {
                    "total_rows": len(results),
                    "all_match": all_match_count,
                    "has_mismatch": len(results) - all_match_count,
                },
                "results": results,
            }
        )
    finally:
        _safe_unlink(tmp_esg)
        _safe_unlink(tmp_tax)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
