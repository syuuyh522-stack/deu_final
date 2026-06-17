#!/usr/bin/env python3
"""
Tally 응답 → Google Sheets 자동 동기화
사용법:
  python sync_tally.py              # 1회 동기화
  python sync_tally.py --watch 30   # 30초마다 자동 반복
  python sync_tally.py --inspect    # Tally 폼 필드명 확인
"""

import argparse
import os
import time
from collections import defaultdict

import gspread
import requests
from google.oauth2.service_account import Credentials

# ─── 설정 ────────────────────────────────────────────────────
TALLY_TOKEN    = os.environ.get("TALLY_TOKEN", "tly-xHfYqWNwxaATbpbg0GHsLlPBP5LKjb1P")
SPREADSHEET_ID = "1tqihK7NxqEEqohQqFlDsjNMXz531K1cpCtwRvni2uk8"
CREDS_FILE     = os.environ.get("GOOGLE_CREDENTIALS", "credentials.json")

FORM_PEER_EVAL = "VLNBql"  # 팀원 평가표
FORM_TEAM_EVAL = "q4zGOO"  # 팀별 평가표

# Tally 폼 질문 제목 키워드 (--inspect 후 실제 값으로 수정)
PEER_FIELD_EVALUATEE = "평가 대상"
PEER_FIELD_SCORE     = "점수"
TEAM_FIELD_TARGET    = "평가 조"
TEAM_FIELD_EVALUATOR = "본인 조"
TEAM_FIELD_SCORE     = "점수"
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

BLUE_HEADER = {"red": 0.851, "green": 0.918, "blue": 0.980}


# ─── Tally API ────────────────────────────────────────────────
def fetch_submissions(form_id: str) -> list[dict]:
    url = f"https://api.tally.so/forms/{form_id}/submissions"
    headers = {"Authorization": f"Bearer {TALLY_TOKEN}"}
    all_submissions = []
    page = 1

    while True:
        resp = requests.get(url, headers=headers, params={"page": page, "limit": 200}, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        all_submissions.extend(data.get("submissions", []))
        if not data.get("hasMore", False):
            break
        page += 1

    return [s for s in all_submissions if s.get("isCompleted")]


def find_value(fields: list, keyword: str):
    for f in fields:
        if keyword.lower() in f.get("label", "").lower():
            val = f.get("value")
            if isinstance(val, list):
                return val[0] if val else None
            return val
    return None


# ─── Google Sheets 연결 ───────────────────────────────────────
def get_spreadsheet():
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


def get_or_create_sheet(ss, title: str):
    try:
        ws = ss.worksheet(title)
        ws.clear()
        return ws
    except gspread.WorksheetNotFound:
        return ss.add_worksheet(title=title, rows=200, cols=20)


# ─── 로우 데이터 저장 ─────────────────────────────────────────
def save_peer_raw(ss, submissions: list[dict]):
    """팀원평가_로우 시트에 원본 응답 저장"""
    ws = get_or_create_sheet(ss, "팀원평가_로우")

    header = ["번호", "제출시간", "피평가자", "점수 (1~5)"]
    rows = [header]
    for i, sub in enumerate(submissions, 1):
        fields     = sub.get("fields", [])
        submitted  = sub.get("createdAt", "")[:19].replace("T", " ")
        evaluatee  = find_value(fields, PEER_FIELD_EVALUATEE) or ""
        score      = find_value(fields, PEER_FIELD_SCORE) or ""
        rows.append([i, submitted, evaluatee, score])

    ws.update(rows, "A1")
    ws.format("A1:D1", {
        "textFormat": {"bold": True},
        "backgroundColor": BLUE_HEADER,
        "horizontalAlignment": "CENTER",
    })
    print(f"  ✓ 팀원평가_로우: {len(submissions)}건 저장")
    return rows[1:]  # header 제외


def save_team_raw(ss, submissions: list[dict]):
    """팀별평가_로우 시트에 원본 응답 저장"""
    ws = get_or_create_sheet(ss, "팀별평가_로우")

    header = ["번호", "제출시간", "평가자 조", "평가 대상 조", "점수 (1~5)", "비고"]
    rows = [header]
    for i, sub in enumerate(submissions, 1):
        fields    = sub.get("fields", [])
        submitted = sub.get("createdAt", "")[:19].replace("T", " ")
        target    = find_value(fields, TEAM_FIELD_TARGET) or ""
        evaluator = find_value(fields, TEAM_FIELD_EVALUATOR) or ""
        score     = find_value(fields, TEAM_FIELD_SCORE) or ""
        note      = "자기평가 제외" if evaluator and target and str(evaluator) == str(target) else ""
        rows.append([i, submitted, evaluator, target, score, note])

    ws.update(rows, "A1")
    ws.format("A1:F1", {
        "textFormat": {"bold": True},
        "backgroundColor": BLUE_HEADER,
        "horizontalAlignment": "CENTER",
    })
    print(f"  ✓ 팀별평가_로우: {len(submissions)}건 저장")
    return rows[1:]


# ─── 집계 및 시트 업데이트 ────────────────────────────────────
def calc_and_update_peer(ss, submissions: list[dict]):
    scores: dict[str, list[float]] = defaultdict(list)
    for sub in submissions:
        fields    = sub.get("fields", [])
        evaluatee = find_value(fields, PEER_FIELD_EVALUATEE)
        score_raw = find_value(fields, PEER_FIELD_SCORE)
        if evaluatee and score_raw is not None:
            try:
                scores[str(evaluatee)].append(float(score_raw))
            except (ValueError, TypeError):
                pass

    peer_avg = {name: round(sum(s) / len(s), 2) for name, s in scores.items() if s}

    ws = ss.worksheet("최종점수")
    updates = []
    for i, (_, name) in enumerate(STUDENTS):
        row = 6 + i
        updates.append({"range": f"C{row}", "values": [[peer_avg.get(name, "")]]})
    ws.batch_update(updates)
    print(f"  ✓ 최종점수 C열 업데이트 ({len(peer_avg)}명 집계)")
    return peer_avg


def calc_and_update_team(ss, submissions: list[dict]):
    scores: dict[int, list[float]] = defaultdict(list)
    for sub in submissions:
        fields    = sub.get("fields", [])
        target_r  = find_value(fields, TEAM_FIELD_TARGET)
        eval_r    = find_value(fields, TEAM_FIELD_EVALUATOR)
        score_raw = find_value(fields, TEAM_FIELD_SCORE)
        if target_r and score_raw is not None:
            try:
                target = int(target_r)
                evaluator = int(eval_r) if eval_r else None
                if evaluator != target:
                    scores[target].append(float(score_raw))
            except (ValueError, TypeError):
                pass

    team_avg = {t: round(sum(s) / len(s), 2) for t, s in scores.items() if s}

    ws = ss.worksheet("팀점수입력")
    updates = []
    for i in range(1, 10):
        updates.append({"range": f"F{3 + i}", "values": [[team_avg.get(i, "")]]})
    ws.batch_update(updates)
    print(f"  ✓ 팀점수입력 F열 업데이트 ({len(team_avg)}개 조 집계)")
    return team_avg


# ─── 필드명 확인 모드 ─────────────────────────────────────────
def inspect_forms():
    for form_id, form_name in [(FORM_PEER_EVAL, "팀원 평가표"), (FORM_TEAM_EVAL, "팀별 평가표")]:
        print(f"\n{'='*50}")
        print(f"폼: {form_name} (ID: {form_id})")
        print("=" * 50)
        submissions = fetch_submissions(form_id)
        if not submissions:
            print("  (완료된 응답 없음)")
            continue
        for f in submissions[0].get("fields", []):
            print(f"  label: {str(f.get('label'))!r:45s} value: {f.get('value')!r}")
        print(f"\n  총 응답 수: {len(submissions)}")


# ─── 메인 ────────────────────────────────────────────────────
def sync_once():
    ts = time.strftime("%H:%M:%S")
    print(f"\n[{ts}] Tally → Google Sheets 동기화 중...")

    ss = get_spreadsheet()

    # 팀원 평가표
    peer_subs = fetch_submissions(FORM_PEER_EVAL)
    print(f"  팀원 평가표: {len(peer_subs)}건 수신")
    save_peer_raw(ss, peer_subs)
    peer_avg = calc_and_update_peer(ss, peer_subs)

    # 팀별 평가표
    team_subs = fetch_submissions(FORM_TEAM_EVAL)
    print(f"  팀별 평가표: {len(team_subs)}건 수신")
    save_team_raw(ss, team_subs)
    team_avg = calc_and_update_team(ss, team_subs)

    print(f"\n  팀원 평균: {peer_avg}")
    print(f"  조별 평균: {team_avg}")
    print("  완료!")


def main():
    parser = argparse.ArgumentParser(description="Tally → Google Sheets 동기화")
    parser.add_argument("--watch", type=int, metavar="초", help="N초마다 자동 반복 동기화")
    parser.add_argument("--inspect", action="store_true", help="Tally 폼 필드명 확인")
    args = parser.parse_args()

    if args.inspect:
        inspect_forms()
        return

    if args.watch:
        print(f"자동 동기화 모드: {args.watch}초마다 실행 (Ctrl+C로 종료)")
        while True:
            try:
                sync_once()
            except Exception as e:
                print(f"  오류: {e}")
            time.sleep(args.watch)
    else:
        sync_once()


if __name__ == "__main__":
    main()
