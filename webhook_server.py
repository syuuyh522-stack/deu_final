#!/usr/bin/env python3
"""Tally 웹훅 수신 → Google Sheets 자동 업데이트 서버"""

import json
import os
from collections import defaultdict
from datetime import datetime

import gspread
import uvicorn
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from google.oauth2.service_account import Credentials

# ─── 설정 (여기만 수정) ───────────────────────────────────────
SPREADSHEET_ID = "1tqihK7NxqEEqohQqFlDsjNMXz531K1cpCtwRvni2uk8"
CREDS_FILE = os.environ.get("GOOGLE_CREDENTIALS", "credentials.json")
DATA_FILE = "scores_backup.json"  # 서버 재시작 시 점수 복원용

# Tally 폼 ID
FORM_PEER_EVAL = "VLNBql"     # 팀원 평가표
FORM_TEAM_EVAL = "q4zGOO"     # 팀별 평가표
FORM_PRESENTATION = "RGjov9"  # 사전 발표 평가표

# Tally 폼 필드 label 키워드 (실제 폼 필드명에 맞게 수정)
# 처음 응답이 오면 GET /logs 에서 실제 필드명 확인 가능
PEER_FIELD_EVALUATEE = "평가 대상"   # 팀원 평가표: 피평가자 이름 필드
PEER_FIELD_SCORE = "점수"            # 팀원 평가표: 점수 필드 (1~5)

TEAM_FIELD_TARGET = "평가 조"        # 팀별 평가표: 평가 대상 조번호 필드
TEAM_FIELD_EVALUATOR = "본인 조"     # 팀별 평가표: 평가자 조번호 필드 (자기 평가 제외용)
TEAM_FIELD_SCORE = "점수"            # 팀별 평가표: 점수 필드 (1~5)
# ─────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

STUDENTS = [
    (1, "박채운"), (1, "천하영"), (1, "함준우"),
    (2, "윤동규"), (2, "오진우"), (2, "정승환"),
    (3, "강윤중"), (3, "이유성"), (3, "장한길"),
    (4, "임소현"), (4, "임수정"), (4, "손나나"),
    (5, "백유민"), (5, "김가희"), (5, "박지환"),
    (6, "윤나영"), (6, "김시현"), (6, "이유진"),
    (7, "이보연"),
    (8, "배연주"), (8, "노윤송"),
    (9, "차아영"), (9, "최고운"),
]

app = FastAPI(title="미인디 점수 자동화 서버")

# 점수 인메모리 저장 (서버 재시작 시 DATA_FILE에서 복원)
peer_scores: dict[str, list[float]] = defaultdict(list)   # {학생이름: [점수들]}
team_scores: dict[int, list[float]] = defaultdict(list)   # {조번호: [점수들]}
raw_logs: list[dict] = []


# ─── 데이터 영속성 ─────────────────────────────────────────────
def save_scores():
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "peer": dict(peer_scores),
            "team": {str(k): v for k, v in team_scores.items()},
        }, f, ensure_ascii=False, indent=2)


def load_scores():
    if not os.path.exists(DATA_FILE):
        return
    with open(DATA_FILE, encoding="utf-8") as f:
        data = json.load(f)
    for name, scores in data.get("peer", {}).items():
        peer_scores[name] = scores
    for team_str, scores in data.get("team", {}).items():
        team_scores[int(team_str)] = scores
    print(f"[복원] 점수 데이터 로드 완료: {DATA_FILE}")


# ─── Google Sheets 업데이트 ────────────────────────────────────
def get_spreadsheet():
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


def push_peer_scores():
    """최종점수 시트 C열 (팀원 평가 평균) 업데이트"""
    ss = get_spreadsheet()
    ws = ss.worksheet("최종점수")
    updates = []
    for i, (_, name) in enumerate(STUDENTS):
        row = 6 + i
        scores = peer_scores.get(name, [])
        avg = round(sum(scores) / len(scores), 2) if scores else ""
        updates.append({"range": f"C{row}", "values": [[avg]]})
    ws.batch_update(updates)
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] 팀원 평가 시트 업데이트 완료")


def push_team_scores():
    """팀점수입력 시트 F열 (조별 평가 평균) 업데이트"""
    ss = get_spreadsheet()
    ws = ss.worksheet("팀점수입력")
    updates = []
    for i in range(1, 10):
        row = 3 + i  # 4~12행
        scores = team_scores.get(i, [])
        avg = round(sum(scores) / len(scores), 2) if scores else ""
        updates.append({"range": f"F{row}", "values": [[avg]]})
    ws.batch_update(updates)
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] 조별 평가 시트 업데이트 완료")


# ─── Tally 필드 파싱 ──────────────────────────────────────────
def find_value(fields: list, keyword: str):
    """필드 label에 keyword가 포함된 첫 번째 필드의 value 반환"""
    for field in fields:
        if keyword.lower() in field.get("label", "").lower():
            val = field.get("value")
            if isinstance(val, list):
                return val[0] if val else None
            return val
    return None


# ─── 웹훅 엔드포인트 ──────────────────────────────────────────
@app.post("/webhook")
async def receive_webhook(request: Request):
    body = await request.json()
    raw_logs.append(body)

    data = body.get("data", {})
    form_id = data.get("formId", "")
    fields = data.get("fields", [])
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"\n[{ts}] 웹훅 수신: formId={form_id}, 필드 {len(fields)}개")

    try:
        if form_id == FORM_PEER_EVAL:
            evaluatee = find_value(fields, PEER_FIELD_EVALUATEE)
            score_raw = find_value(fields, PEER_FIELD_SCORE)
            print(f"  팀원 평가: 대상={evaluatee}, 점수={score_raw}")

            if evaluatee and score_raw is not None:
                peer_scores[evaluatee].append(float(score_raw))
                save_scores()
                push_peer_scores()
            else:
                print("  ⚠ 필드 매핑 실패 → GET /logs 에서 실제 필드명 확인 후 상단 설정 수정")

        elif form_id == FORM_TEAM_EVAL:
            target_raw = find_value(fields, TEAM_FIELD_TARGET)
            evaluator_raw = find_value(fields, TEAM_FIELD_EVALUATOR)
            score_raw = find_value(fields, TEAM_FIELD_SCORE)
            print(f"  팀별 평가: 대상={target_raw}, 평가자={evaluator_raw}, 점수={score_raw}")

            if target_raw and score_raw is not None:
                target = int(target_raw)
                evaluator = int(evaluator_raw) if evaluator_raw else None
                if evaluator == target:
                    print("  → 자기 조 평가 제외")
                else:
                    team_scores[target].append(float(score_raw))
                    save_scores()
                    push_team_scores()
            else:
                print("  ⚠ 필드 매핑 실패 → GET /logs 에서 실제 필드명 확인 후 상단 설정 수정")

        elif form_id == FORM_PRESENTATION:
            print("  사전 발표 평가 수신 (수동 입력 필요)")

    except Exception as e:
        print(f"  오류: {e}")
        return JSONResponse({"status": "error", "detail": str(e)}, status_code=500)

    return JSONResponse({"status": "ok"})


# ─── 유틸 엔드포인트 ──────────────────────────────────────────
@app.get("/")
async def root():
    return JSONResponse({
        "service": "미인디 점수 자동화 서버",
        "endpoints": {
            "POST /webhook": "Tally 웹훅 수신",
            "GET  /scores":  "현재 집계 점수 조회",
            "GET  /logs":    "최근 웹훅 원시 데이터 (필드명 확인용)",
            "POST /resync":  "구글 시트 강제 재동기화",
            "POST /reset":   "점수 초기화",
        },
    })


@app.get("/scores")
async def get_scores():
    """현재 집계된 평균 점수 조회"""
    peer_avg = {n: round(sum(s) / len(s), 2) for n, s in peer_scores.items() if s}
    team_avg = {str(t): round(sum(s) / len(s), 2) for t, s in team_scores.items() if s}
    peer_count = {n: len(s) for n, s in peer_scores.items() if s}
    team_count = {str(t): len(s) for t, s in team_scores.items() if s}
    return JSONResponse({
        "peer_evaluation":  {"averages": peer_avg, "response_count": peer_count},
        "team_evaluation":  {"averages": team_avg, "response_count": team_count},
    })


@app.get("/logs")
async def get_logs(n: int = 3):
    """최근 n개 웹훅 원시 데이터 (Tally 필드명 확인용)"""
    recent = raw_logs[-n:]
    # 필드 label 목록만 추출해서 보기 쉽게
    summary = []
    for log in recent:
        fields = log.get("data", {}).get("fields", [])
        summary.append({
            "formId": log.get("data", {}).get("formId"),
            "submittedAt": log.get("data", {}).get("submittedAt"),
            "fields": [{"label": f.get("label"), "value": f.get("value")} for f in fields],
        })
    return JSONResponse({"recent_logs": summary})


@app.post("/resync")
async def resync():
    """현재 메모리 데이터를 구글 시트에 강제 재동기화"""
    push_peer_scores()
    push_team_scores()
    return JSONResponse({"status": "ok", "message": "구글 시트 업데이트 완료"})


@app.post("/reset")
async def reset():
    """점수 초기화 (주의: 복구 불가)"""
    peer_scores.clear()
    team_scores.clear()
    if os.path.exists(DATA_FILE):
        os.remove(DATA_FILE)
    return JSONResponse({"status": "ok", "message": "점수 초기화 완료"})


# ─── 시작 ─────────────────────────────────────────────────────
@app.on_event("startup")
async def startup():
    load_scores()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    print("=" * 50)
    print("  미인디 점수 자동화 서버")
    print(f"  http://localhost:{port}")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=port)
