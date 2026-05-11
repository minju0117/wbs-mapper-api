"""
WBS 매퍼 FastAPI 서버
POST /map-wbs  — download WBS 업로드 → 원본 템플릿과 매핑 → 완성된 엑셀 반환
템플릿 파일은 서버에 original_wbs.xlsx 로 번들되어 있음
"""

import copy
import io
import os
import shutil
import tempfile

import openpyxl
from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from wbs_mapper import (
    COL_GUBUN,
    TEMPLATE_DATA_START_ROW,
    fill_template,
    load_download_rows,
    unmerge_ab_data_area,
)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "original_wbs.xlsx")

app = FastAPI(title="WBS Mapper API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)


@app.get("/")
def health():
    return {"status": "ok"}


@app.post("/map-wbs")
async def map_wbs(
    download: UploadFile = File(..., description="Lovable에서 다운받은 WBS xlsx"),
    sheet: str = Query(default="V1.4", description="템플릿 시트 이름"),
):
    if not download.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail=f"{download.filename}: xlsx 파일만 허용됩니다.")

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail="서버에 원본 템플릿 파일이 없습니다.")

    with tempfile.TemporaryDirectory() as tmpdir:
        dl_path  = os.path.join(tmpdir, "download.xlsx")
        tpl_path = os.path.join(tmpdir, "template.xlsx")
        out_path = os.path.join(tmpdir, "output.xlsx")

        with open(dl_path, "wb") as f:
            f.write(await download.read())

        shutil.copy2(TEMPLATE_PATH, tpl_path)

        try:
            dl_rows = load_download_rows(dl_path)
        except Exception as e:
            raise HTTPException(status_code=422, detail=f"다운로드 WBS 파싱 오류: {e}")

        if not dl_rows:
            raise HTTPException(status_code=422, detail="다운로드 WBS에 데이터가 없습니다.")

        # 템플릿에서 구분별 fill 색상 수집
        wb_orig = openpyxl.load_workbook(tpl_path)
        sheet_name = sheet if sheet in wb_orig.sheetnames else wb_orig.sheetnames[0]
        ws_orig = wb_orig[sheet_name]
        gubun_fills = {}
        for row in range(TEMPLATE_DATA_START_ROW, ws_orig.max_row + 1):
            val = ws_orig.cell(row, COL_GUBUN).value
            if val and val not in gubun_fills:
                gubun_fills[val] = copy.copy(ws_orig.cell(row, COL_GUBUN).fill)

        shutil.copy2(tpl_path, out_path)
        wb = openpyxl.load_workbook(out_path)

        # 히든 시트 삭제 (대상 시트 제외)
        for name in [n for n in wb.sheetnames if wb[n].sheet_state in ("hidden", "veryHidden") and n != sheet_name]:
            del wb[name]

        ws = wb[sheet_name if sheet_name in wb.sheetnames else wb.sheetnames[0]]

        unmerge_ab_data_area(ws, TEMPLATE_DATA_START_ROW)
        fill_template(ws, dl_rows, TEMPLATE_DATA_START_ROW, gubun_fills)

        last_data_row = TEMPLATE_DATA_START_ROW + len(dl_rows) - 1
        if ws.max_row > last_data_row:
            ws.delete_rows(last_data_row + 1, ws.max_row - last_data_row)

        wb.save(out_path)

        with open(out_path, "rb") as f:
            content = f.read()

    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="output_WBS.xlsx"'},
    )
