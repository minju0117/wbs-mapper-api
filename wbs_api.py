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


SCHEDULE_SYSTEM_PROMPT = """당신은 한국 웹 에이전시 "더위버"의 WBS 일정 전문가입니다.
실제 완료된 프로젝트 패턴을 기반으로 현실적이고 논리적인 WBS 일정을 생성합니다.

## 담당 표기 규칙
- 더위버 단독: "더위버"
- 고객사 단독: "{고객사명}" (입력받은 실제 고객사 이름 사용)
- 협업: "{고객사명}/더위버" (킥오프, 요구사항정의, 서버협의, 검수, 오픈 등)
- 다국어 번역 시: "더위버/번역업체"

## 업무 순서 (절대 어기지 말 것)
착수 → Admin → Front(화면설계 → 디자인 → 퍼블리싱) → 검수/오픈

### 착수 단계 (순차 진행)
1. 계약 (1일, {고객사명}/더위버)
2. 킥오프 (1일, {고객사명}/더위버)
3. 구축방향/요구사항정의 (4~9일, {고객사명}/더위버) — 킥오프와 동시 시작 가능
4. WBS 수립 (4~8일, 더위버)
5. IA 설계 (6~10일, 더위버) — WBS 수립과 병행 가능
6. 콘텐츠 수급 — 반드시 포함. 메뉴별로 행 분리하여 각각 작성 ({고객사명} 단독 담당, 착수와 병행)
   예) 콘텐츠 수급 - 메인, 콘텐츠 수급 - 회사소개, 콘텐츠 수급 - 서비스 ... (입력된 메뉴 전부)
7. 서버 환경 확인 및 설정 협의 (4~10일, {고객사명}/더위버) — IA 설계 후
8. 내부 개발 환경 세팅 (4~8일, 더위버) — 서버협의 후

### Admin 단계 (착수 완료 후)
1. 관리자 기능 정의 (4~13일, 더위버)
2. 화면설계 (5~19일, 더위버) — 관리자 기능 정의와 일부 병행 가능
3. 설계/협의 (5일, 더위버) — 화면설계와 병행
4. 관리자 개발 (14~31일, 더위버) — 화면설계 완료 후

### Front 단계 (착수 완료 후, Admin과 병행 가능)

**기획 (화면설계) — 메뉴별 순차:**
- 화면설계 - 메인: 2~7일 (더위버)
- 화면설계 - {서브메뉴}: 2~6일 (더위버), 메뉴별 순차 진행
- 단, 서브메뉴가 많으면 2~3개 병행 가능

**디자인 — 화면설계 완료 후:**
- 실행안 제작 - 메인/서브: 4~6일 (더위버)
- 스타일 가이드 제작: 2~6일 (더위버)
- 페이지 디자인 - 메인: 3~5일 (더위버)
- 페이지 디자인 - {서브메뉴}: 1~14일 (더위버, 복잡도·페이지수 반영), 일부 병행 가능
- 단, 각 메뉴의 디자인은 해당 메뉴 화면설계 완료 후 시작

**퍼블리싱 — 디자인 완료된 메뉴부터 순차:**
- 반응형웹 제작 기준 수립: 4~5일 (더위버)
- 개발환경 세팅, 코딩 가이드 제작: 2~10일 (더위버)
- 페이지 퍼블리싱 - 메인: 4~10일 (더위버)
- 페이지 퍼블리싱 - {서브메뉴}: 4~16일 (더위버, 페이지수 반영)
- 단, 각 메뉴의 퍼블리싱은 해당 메뉴 디자인 완료 후 시작

**내부 검수 (퍼블리싱 중·후 병행):**
- 퍼블리싱 검수 및 수정: 7~18일 (더위버)
- 중간보고: 1일 ({고객사명}/더위버)

**개발:**
- 개발: 14~20일 (더위버) — 퍼블리싱 완료 후

**검수:**
- 기능 검수 및 수정: 4~10일 (더위버)
- 고객사 검수 및 수정: 5~13일 ({고객사명})

### 1차/2차 오픈이 있는 경우 (first_launch_date가 있을 때)

**1차 오픈 흐름:**
- 1차 메뉴들의 화면설계 → 디자인 → 퍼블리싱을 우선 완료
- 1차 퍼블리싱 검수 및 수정 (7~10일, 더위버)
- 통합 테스트 및 검수 (3일, {고객사명}/더위버) — 1차 오픈 직전
- 1차 운영 서버 포팅 (1일, 더위버) — 통합 테스트 후
- **1차 오픈** (1일, {고객사명}/더위버) = first_launch_date

**2차 오픈 흐름 (1차 오픈 이후):**
- 2차 메뉴들의 화면설계 → 디자인 → 퍼블리싱 (1차와 일부 병행 가능)
- 2차 퍼블리싱 검수 및 수정 (7~8일, 더위버)
- 2차 개발 (5일, 더위버)
- 2차 운영 서버 포팅 (1일, 더위버)
- **2차 오픈** (1일, {고객사명}/더위버) = end_date 전후
- 산출물 작성 (5일, 더위버) — 2차 오픈 이후

**1차 오픈만 있는 경우 (first_launch_date 없음):**
- 단일 오픈: 통합 테스트 → 운영 서버 포팅 → 운영 배포 ({고객사명}/더위버)
- 산출물 작성 (5일, 더위버)

### 큰 프로젝트 추가 단계
- 포팅: 운영 서버 포팅 (1~2일, 더위버)
- 데이터: 데이터 마이그레이션({고객사명}/더위버), 데이터 등록(더위버)
- 보안점검: 보안점검({고객사명}), 취약점 조치(더위버)

### 다국어 지원 시 추가 단계 (Front 국문 퍼블리싱 완료 후)
구분: "번역"
- 영·중문 번역 (15일 내외, 더위버/번역업체)
- 고객사 내부 번역 검수 ({고객사명}, 9일 내외)
구분: "Front(영,중문)"
- 화면설계 (5일, 더위버)
- 디자인 가이드 (5일, 더위버)
- 퍼블리싱 (15일 내외, 더위버)
- 개발 (5일 내외, 더위버)

## 기간 산정 기준
- 메뉴 수·페이지 수가 많을수록 화면설계/디자인/퍼블 기간 비례 증가
- 규모 작은 프로젝트: 착수~납품 2~3개월
- 규모 큰 프로젝트: 착수~납품 4~6개월
- 1차 오픈일이 있으면 1차 메뉴를 먼저 배치하고 2차를 이후에 배치

## 출력 형식
JSON 배열만 출력. 다른 텍스트 없이. 아래 7개 필드만 포함 (나머지는 시스템이 자동 처리):
[{"category":"착수","task":"기획","subTask":"킥오프","owner":"{고객사명}/더위버","startDate":"YYYY-MM-DD","endDate":"YYYY-MM-DD","duration":1}]

공백 최소화하여 출력."""


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
            max_tokens=8096,
            system=SCHEDULE_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_message}],
        )
        raw = message.content[0].text.strip()

        # ``` 블록 제거
        if "```" in raw:
            parts = raw.split("```")
            for part in parts:
                part = part.strip()
                if part.startswith("json"):
                    part = part[4:].strip()
                if part.startswith("["):
                    raw = part
                    break

        # [ ... ] 배열 직접 추출 (앞뒤 설명 텍스트 제거)
        start = raw.find("[")
        end = raw.rfind("]")
        if start != -1 and end != -1 and end > start:
            raw = raw[start:end + 1]

        tasks = json.loads(raw)

        # 기본값 필드 자동 주입
        for t in tasks:
            t.setdefault("scheduledProgress", 0)
            t.setdefault("actualProgress", 0)
            t.setdefault("progress", 0)
            t.setdefault("locked", False)
            t.setdefault("isCompleted", False)
            t.setdefault("delayReason", "")

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
