#!/usr/bin/env python3
"""동의대학교 미인디 기말평가 점수집계표 Google Sheets 자동 생성"""

import os
import sys
import argparse
import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

TEAMS = [
    (1, "1조 (박채운, 천하영, 함준우)"),
    (2, "2조 (윤동규, 오진우, 정승환)"),
    (3, "3조 (강윤중, 이유성, 장한길)"),
    (4, "4조 (임소현, 임수정, 손나나)"),
    (5, "5조 (백유민, 김가희, 박지환)"),
    (6, "6조 (윤나영, 김시현, 이유진)"),
    (7, "7조 (이보연)"),
    (8, "8조 (배연주, 노윤송)"),
    (9, "9조 (차아영, 최고운)"),
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
YELLOW_AUTO = {"red": 0.980, "green": 0.980, "blue": 0.851}
GRAY_BG = {"red": 0.9, "green": 0.9, "blue": 0.9}
WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}


def bold_center(bg=None):
    fmt = {"textFormat": {"bold": True}, "horizontalAlignment": "CENTER"}
    if bg:
        fmt["backgroundColor"] = bg
    return fmt


def setup_team_sheet(ws):
    """시트 1: 팀점수입력"""
    # 제목
    ws.update("A1", [["동의대학교 미인디 기말평가 — 팀점수입력"]])
    ws.format("A1", {"textFormat": {"bold": True, "fontSize": 13}})

    # 3행 헤더
    headers = [
        "조번호",
        "팀 (조)",
        "안지수 교수 점수\n(100점 만점)",
        "홍승윤 강사 점수\n(100점 만점)",
        "교수 점수 평균\n(자동)",
        "조별 평가 평균\n(1~5)",
    ]
    ws.update("A3:F3", [headers])
    ws.format("A3:F3", {**bold_center(BLUE_HEADER), "wrapStrategy": "WRAP"})

    # 4~12행 데이터
    rows = []
    for i, (num, name) in enumerate(TEAMS):
        r = 4 + i
        rows.append([
            num,
            name,
            "",
            "",
            f"=IFERROR(AVERAGE(C{r}:D{r}),\"\")",
            "",
        ])
    ws.update("A4:F12", rows, value_input_option="USER_ENTERED")

    # 스타일
    ws.format("A4:A12", {"horizontalAlignment": "CENTER"})
    ws.format("C4:D12", {"horizontalAlignment": "CENTER"})
    ws.format("E4:E12", {"backgroundColor": YELLOW_AUTO, "horizontalAlignment": "CENTER"})
    ws.format("F4:F12", {"horizontalAlignment": "CENTER"})

    # 열 너비
    requests = [
        col_width(ws, 0, 80),   # A
        col_width(ws, 1, 300),  # B
        col_width(ws, 2, 160),  # C
        col_width(ws, 3, 160),  # D
        col_width(ws, 4, 160),  # E
        col_width(ws, 5, 160),  # F
    ]
    ws.spreadsheet.batch_update({"requests": requests})

    # 3행 높이
    ws.spreadsheet.batch_update({"requests": [row_height(ws, 2, 50)]})

    print("  ✓ 팀점수입력 시트 완료")


def setup_final_sheet(ws):
    """시트 2: 최종점수"""
    # 제목
    ws.update("A1", [["동의대학교 미인디 기말평가 — 최종점수"]])
    ws.format("A1", {"textFormat": {"bold": True, "fontSize": 13}})

    # 가중치 행
    ws.update("A2:E2", [["가중치(%)", "교수", "조별", "팀원", "합계"]])
    ws.format("A2:E2", bold_center(GRAY_BG))

    ws.update("A3:E3", [["비중", 50, 20, 30, "=SUM(B3:D3)"]], value_input_option="USER_ENTERED")
    ws.format("A3:E3", {"horizontalAlignment": "CENTER"})

    # 주의 문구
    ws.update("F2", [["← B3:D3만 수정 (합계 100 유지)"]])
    ws.format("F2", {"textFormat": {"italic": True, "foregroundColor": {"red": 0.5, "green": 0.5, "blue": 0.5}}})

    # 4행 헤더 (수식은 3행 가중치를 참조하므로 $B$3, $C$3, $D$3 사용)
    col_headers = [
        "조번호",
        "이름",
        "팀원 평가 평균\n(1~5)",
        "교수 점수\n(2인 평균)",
        "조별 환산(100)",
        "팀원 환산(100)",
        "최종 점수(100)",
    ]
    ws.update("A5:G5", [col_headers])
    ws.format("A5:G5", {**bold_center(BLUE_HEADER), "wrapStrategy": "WRAP"})

    # 학생 데이터 (6행부터)
    rows = []
    for i, (team_num, name) in enumerate(STUDENTS):
        r = 6 + i
        rows.append([
            team_num,
            name,
            "",  # C: 직접 입력
            f"=VLOOKUP(A{r},팀점수입력!$A$4:$F$12,5,0)",
            f"=VLOOKUP(A{r},팀점수입력!$A$4:$F$12,6,0)/5*100",
            f'=IF(C{r}="","",C{r}/5*100)',
            f'=IF(C{r}="",(D{r}*$B$3+E{r}*$C$3)/($B$3+$C$3),(D{r}*$B$3+E{r}*$C$3+F{r}*$D$3)/SUM($B$3:$D$3))',
        ])

    last_row = 6 + len(STUDENTS) - 1
    ws.update(f"A6:G{last_row}", rows, value_input_option="USER_ENTERED")

    # 스타일
    ws.format(f"A6:B{last_row}", {"horizontalAlignment": "CENTER"})
    ws.format(f"C6:C{last_row}", {"horizontalAlignment": "CENTER"})
    ws.format(f"D6:G{last_row}", {
        "backgroundColor": YELLOW_AUTO,
        "horizontalAlignment": "CENTER",
        "numberFormat": {"type": "NUMBER", "pattern": "0.00"},
    })

    # 열 너비
    requests = [
        col_width(ws, 0, 80),   # A
        col_width(ws, 1, 100),  # B
        col_width(ws, 2, 160),  # C
        col_width(ws, 3, 140),  # D
        col_width(ws, 4, 130),  # E
        col_width(ws, 5, 130),  # F
        col_width(ws, 6, 130),  # G
    ]
    ws.spreadsheet.batch_update({"requests": requests})

    # 5행 높이
    ws.spreadsheet.batch_update({"requests": [row_height(ws, 4, 50)]})

    print("  ✓ 최종점수 시트 완료")


def col_width(ws, col_index, width_px):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": ws.id,
                "dimension": "COLUMNS",
                "startIndex": col_index,
                "endIndex": col_index + 1,
            },
            "properties": {"pixelSize": width_px},
            "fields": "pixelSize",
        }
    }


def row_height(ws, row_index, height_px):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": ws.id,
                "dimension": "ROWS",
                "startIndex": row_index,
                "endIndex": row_index + 1,
            },
            "properties": {"pixelSize": height_px},
            "fields": "pixelSize",
        }
    }


SPREADSHEET_ID = "1tqihK7NxqEEqohQqFlDsjNMXz531K1cpCtwRvni2uk8"


def get_or_create_sheet(spreadsheet, title, rows=50, cols=10):
    """시트가 있으면 반환, 없으면 새로 생성"""
    try:
        ws = spreadsheet.worksheet(title)
        ws.clear()
        return ws
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)


def main():
    parser = argparse.ArgumentParser(description="미인디 기말평가 점수집계표 Google Sheets 채우기")
    parser.add_argument(
        "--creds",
        default=os.environ.get("GOOGLE_CREDENTIALS", "credentials.json"),
        help="서비스 계정 JSON 파일 경로 (기본값: credentials.json)",
    )
    parser.add_argument(
        "--share",
        nargs="*",
        metavar="EMAIL",
        help="편집자로 공유할 이메일 주소 (여러 개 가능)",
    )
    parser.add_argument(
        "--public",
        action="store_true",
        help="링크 공유 활성화 (링크 소유자 편집 가능)",
    )
    args = parser.parse_args()

    if not os.path.exists(args.creds):
        print(f"오류: 인증 파일을 찾을 수 없습니다 → {args.creds}")
        print("\n해결 방법:")
        print("  1. Google Cloud Console에서 서비스 계정 JSON 키를 다운로드")
        print("  2. 파일 이름을 credentials.json으로 바꿔 이 폴더에 저장")
        print("  3. 또는 --creds /path/to/key.json 옵션으로 경로 지정")
        sys.exit(1)

    print("Google API 인증 중...")
    creds = Credentials.from_service_account_file(args.creds, scopes=SCOPES)
    gc = gspread.authorize(creds)

    print(f"기존 스프레드시트 열기 중... (ID: {SPREADSHEET_ID})")
    try:
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    except gspread.SpreadsheetNotFound:
        print("오류: 스프레드시트를 찾을 수 없습니다.")
        print("서비스 계정 이메일을 해당 구글 시트의 편집자로 공유했는지 확인하세요.")
        print(f"서비스 계정 이메일은 credentials.json 안의 'client_email' 값입니다.")
        sys.exit(1)

    print("시트 구성 중...")
    team_sheet = get_or_create_sheet(spreadsheet, "팀점수입력")
    final_sheet = get_or_create_sheet(spreadsheet, "최종점수")

    setup_team_sheet(team_sheet)
    setup_final_sheet(final_sheet)

    # 공유 설정
    if args.public:
        spreadsheet.share(None, perm_type="anyone", role="writer")
        print("  ✓ 링크 공유 활성화 (편집자)")

    if args.share:
        for email in args.share:
            spreadsheet.share(email, perm_type="user", role="writer")
            print(f"  ✓ 공유 완료: {email}")

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}"
    print(f"\n✅ 완료!")
    print(f"🔗 스프레드시트 URL:\n   {url}")
    print("\n다음 단계:")
    print("  1. 위 URL로 접속하여 시트 확인")
    print("  2. 팀점수입력 C·D열에 교수/강사 점수 입력")
    print("  3. 팀점수입력 F열에 Tally 조별 평가 평균 입력")
    print("  4. 최종점수 C열에 Tally 팀원 평가 평균 입력")


if __name__ == "__main__":
    main()
