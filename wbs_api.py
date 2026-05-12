"""
WBS 매퍼 FastAPI 서버
POST /map-wbs  — download WBS 업로드 → 원본 템플릿과 매핑 → 완성된 엑셀 반환
템플릿 파일은 서버에 original_wbs.xlsx 로 번들되어 있음
"""

import copy
import io
import json
import os
import shutil
import tempfile
from typing import List, Optional

import anthropic
import openpyxl
from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from wbs_mapper import (
    COL_GUBUN,
    TEMPLATE_DATA_START_ROW,
    copy_previous_version_sheets,
    extract_header_info,
    fill_template,
    load_download_rows,
    unmerge_ab_data_area,
    update_holiday_sheet,
)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "original_wbs.xlsx")

app = FastAPI(title="WBS Mapper API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)


class MenuInfo(BaseModel):
    name: str
    phase: str = "1차"
    pages: int = 1


class ScheduleRequest(BaseModel):
    project: str
    client: str
    start_date: str
    first_launch_date: Optional[str] = None
    end_date: str
    multilingual: bool = False
    conditions: Optional[str] = None
    menus: List[MenuInfo] = []


SCHEDULE_SYSTEM_PROMPT = """당신은 한국 웹 에이전시의 WBS 일정 전문가입니다.
주어진 프로젝트 정보를 바탕으로 현실적이고 논리적인 WBS 일정을 JSON으로 생성합니다.

## 담당 구분 (반드시 아래 셋 중 하나만 사용)
- 더위버: 내부 단독 작업 (기획, 설계, 개발, 디자인)
- 고객사: 클라이언트 단독 작업 (콘텐츠 제공, 자료 전달)
- 더위버/고객사: 협업 필요 작업 (킥오프, 요구사항 정의, 검수, 최종 확인)

## 업무 의존관계 (절대 어기지 말 것)
- 착수(킥오프 → 요구사항 정의 → WBS 수립 → IA 설계)는 순차 진행
- 화면설계는 IA 설계 완료 후 시작
- 디자인은 화면설계 완료 후 시작
- 개발은 디자인 완료 후 시작 (퍼블리싱 포함)
- 검수/QA는 개발 완료 후 시작
- 콘텐츠 수급(고객사 담당)은 착수 단계와 병행 가능
- 같은 담당자가 동시에 두 업무를 진행하지 않도록 조정

## 기간 산정 기준
- 페이지 수, 복잡도, 다국어 여부를 반영하여 현실적으로 산정
- 다국어 지원 시 개발/퍼블 기간 1.5배
- 1차 오픈일이 있으면 1차 대상 메뉴를 우선 배치

## 출력 형식
JSON 배열만 출력하세요. 다른 텍스트 없이:
[
  {
    "구분": "착수",
    "업무": "기획",
    "세부항목": "킥오프",
    "담당": "더위버/고객사",
    "시작일": "YYYY-MM-DD",
    "종료일": "YYYY-MM-DD",
    "기간": 1,
    "계획": 0,
    "진척": 0
  }
]"""


@app.get("/")
def health():
    return {"status": "ok"}


@app.post("/generate-schedule")
async def generate_schedule(request: ScheduleRequest):
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY 환경변수가 설정되지 않았습니다.")

    menu_lines = "\n".join(
        f"  - {m.name}: {m.phase} 오픈, {m.pages}페이지"
        for m in request.menus
    )

    user_message = f"""다음 웹 프로젝트의 WBS 일정을 생성해주세요.

프로젝트명: {request.project}
고객사: {request.client}
시작일: {request.start_date}
1차 오픈일: {request.first_launch_date or '없음'}
종료일: {request.end_date}
다국어 지원: {'예' if request.multilingual else '아니오'}
기타 조건: {request.conditions or '없음'}

메뉴 구성:
{menu_lines}

위 정보를 바탕으로 업무 의존관계를 지키는 현실적인 WBS 일정을 JSON으로 생성해주세요."""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4096,
            system=SCHEDULE_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_message}],
        )
        raw = message.content[0].text.strip()

        # JSON 블록 추출 (```json ... ``` 감싸진 경우 대비)
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.strip()

        tasks = json.loads(raw)
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=500, detail=f"Claude 응답 파싱 오류: {e}")
    except anthropic.APIError as e:
        raise HTTPException(status_code=502, detail=f"Claude API 오류: {e}")

    return {"tasks": tasks}


@app.post("/map-wbs")
async def map_wbs(
    download: UploadFile = File(..., description="Lovable에서 다운받은 WBS xlsx"),
    previous: UploadFile | None = File(None, description="이전 출력 WBS (버전 히스토리용, 선택)"),
    version: str | None = Query(None, description="WBS 버전 (예: v2)"),
    project: str | None = Query(None, description="프로젝트명"),
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

        prev_path = None
        if previous:
            prev_path = os.path.join(tmpdir, "previous.xlsx")
            with open(prev_path, "wb") as f:
                f.write(await previous.read())

        shutil.copy2(TEMPLATE_PATH, tpl_path)

        try:
            dl_rows = load_download_rows(dl_path)
        except Exception as e:
            raise HTTPException(status_code=422, detail=f"다운로드 WBS 파싱 오류: {e}")

        # query param 미전달 시 다운로드 WBS 상단에서 자동 추출
        if not project or not version:
            extracted_project, extracted_version = extract_header_info(dl_path)
            if not project:
                project = extracted_project
            if not version:
                version = extracted_version

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

        # ① 히든 시트 전부 삭제 (공휴일만 보존)
        for name in [
            n for n in wb.sheetnames
            if wb[n].sheet_state in ("hidden", "veryHidden") and n != "공휴일"
        ]:
            del wb[name]

        # ② 이전 출력의 버전 시트를 hidden으로 복사 (삭제 이후에 추가)
        if prev_path:
            copy_previous_version_sheets(wb, prev_path, version)

        # ③ 작업 시트 선택 및 시트명 버전으로 변경
        ws = wb[sheet_name if sheet_name in wb.sheetnames else wb.sheetnames[0]]
        if version:
            ws.title = version

        # ④ 프로젝트명(A1), Last update(A2) 갱신
        from datetime import datetime as _dt
        if project:
            ws.cell(row=1, column=1).value = project
        ws.cell(row=2, column=1).value = f"▥ Last update : {_dt.today().strftime('%Y-%m-%d')}"

        # 공휴일 시트 자동 업데이트
        from datetime import datetime as _dt2
        dates = [r[k] for r in dl_rows for k in ("시작일", "종료일") if isinstance(r.get(k), _dt2)]
        if dates:
            sy, ey = min(d.year for d in dates), max(d.year for d in dates)
        else:
            sy = ey = _dt2.today().year
        holiday_ref = update_holiday_sheet(wb, sy, ey)

        unmerge_ab_data_area(ws, TEMPLATE_DATA_START_ROW)
        fill_template(ws, dl_rows, TEMPLATE_DATA_START_ROW, gubun_fills, holiday_ref)

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
