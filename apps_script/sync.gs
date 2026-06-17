// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 동의대학교 미인디 기말평가 — Tally 자동 동기화
// [사용법] 구글 시트 → 확장 프로그램 → Apps Script
//          → 이 코드 전체 붙여넣기 → 저장 → 새로고침
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ─── 설정 ─────────────────────────────────────────
const TALLY_TOKEN    = "tly-xHfYqWNwxaATbpbg0GHsLlPBP5LKjb1P";
const FORM_PEER_EVAL = "VLNBql";  // 팀원 평가표
const FORM_TEAM_EVAL = "q4zGOO";  // 팀별 평가표

// Tally 질문 제목 키워드 — 필드명 확인 메뉴 실행 후 맞게 수정
const PEER_FIELD_EVALUATEE = "평가 대상";
const PEER_FIELD_SCORE     = "점수";
const TEAM_FIELD_TARGET    = "평가 조";
const TEAM_FIELD_EVALUATOR = "본인 조";
const TEAM_FIELD_SCORE     = "점수";
// ──────────────────────────────────────────────────

const STUDENTS = [
  [1,"박채운"],[1,"천하영"],[1,"함준우"],
  [2,"윤동규"],[2,"오진우"],[2,"정승환"],
  [3,"강윤중"],[3,"이유성"],[3,"장한길"],
  [4,"임소현"],[4,"임수정"],[4,"손나나"],
  [5,"백유민"],[5,"김가희"],[5,"박지환"],
  [6,"윤나영"],[6,"김시현"],[6,"이유진"],
  [7,"이보연"],
  [8,"배연주"],[8,"노윤송"],
  [9,"차아영"],[9,"최고운"],
];

// ─── 메뉴 등록 (시트 열릴 때 자동 실행) ──────────
// ※ 이 함수는 직접 실행하지 마세요. 저장 후 구글 시트를 새로고침하면 자동 실행됩니다.
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("📊 점수 동기화")
      .addItem("🔄 지금 동기화", "syncAll")
      .addSeparator()
      .addItem("🔍 Tally 필드명 확인", "inspectFields")
      .addSeparator()
      .addItem("⏰ 자동 동기화 켜기 (10분마다)", "setupTrigger")
      .addItem("⏹ 자동 동기화 끄기", "removeTriggers")
      .addToUi();
  } catch (e) {
    // 스크립트 에디터에서 직접 실행 시 무시
  }
}

// ─── Tally API 호출 ───────────────────────────────
function fetchSubmissions(formId) {
  const allSubs = [];
  let page = 1;

  while (true) {
    const url  = `https://api.tally.so/forms/${formId}/submissions?limit=200&page=${page}`;
    const resp = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: { Authorization: `Bearer ${TALLY_TOKEN}` },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) {
      throw new Error(`Tally API 오류 (${resp.getResponseCode()}): ${resp.getContentText()}`);
    }

    const data = JSON.parse(resp.getContentText());
    allSubs.push(...(data.submissions || []).filter(s => s.isCompleted));
    if (!data.hasMore) break;
    page++;
  }

  return allSubs;
}

function findValue(fields, keyword) {
  const kw = keyword.toLowerCase();
  for (const f of fields) {
    if (f.label && f.label.toLowerCase().includes(kw)) {
      const val = f.value;
      return Array.isArray(val) ? (val[0] ?? null) : val;
    }
  }
  return null;
}

// ─── 로우 데이터 저장 ─────────────────────────────
function saveRawPeer(ss, submissions) {
  const ws   = getOrCreateSheet(ss, "팀원평가_로우");
  const rows = [["번호", "제출시간", "피평가자", "점수 (1~5)"]];

  submissions.forEach((sub, i) => {
    const f = sub.fields || [];
    rows.push([
      i + 1,
      (sub.createdAt || "").replace("T", " ").slice(0, 19),
      findValue(f, PEER_FIELD_EVALUATEE) ?? "",
      findValue(f, PEER_FIELD_SCORE)     ?? "",
    ]);
  });

  ws.clearContents();
  ws.getRange(1, 1, rows.length, 4).setValues(rows);
  styleHeader(ws, 4);
}

function saveRawTeam(ss, submissions) {
  const ws   = getOrCreateSheet(ss, "팀별평가_로우");
  const rows = [["번호", "제출시간", "평가자 조", "평가 대상 조", "점수 (1~5)", "비고"]];

  submissions.forEach((sub, i) => {
    const f         = sub.fields || [];
    const target    = findValue(f, TEAM_FIELD_TARGET)    ?? "";
    const evaluator = findValue(f, TEAM_FIELD_EVALUATOR) ?? "";
    const score     = findValue(f, TEAM_FIELD_SCORE)     ?? "";
    rows.push([
      i + 1,
      (sub.createdAt || "").replace("T", " ").slice(0, 19),
      evaluator,
      target,
      score,
      String(evaluator) === String(target) ? "자기평가 제외" : "",
    ]);
  });

  ws.clearContents();
  ws.getRange(1, 1, rows.length, 6).setValues(rows);
  styleHeader(ws, 6);
}

// ─── 집계 → 시트 업데이트 ─────────────────────────
function updatePeerScores(ss, submissions) {
  const scores = {};
  let unmatched = 0;

  submissions.forEach(sub => {
    const f         = sub.fields || [];
    const evaluatee = findValue(f, PEER_FIELD_EVALUATEE);
    const scoreRaw  = findValue(f, PEER_FIELD_SCORE);
    if (evaluatee && scoreRaw !== null && scoreRaw !== "") {
      const score = parseFloat(scoreRaw);
      if (!isNaN(score)) {
        scores[evaluatee] = scores[evaluatee] || [];
        scores[evaluatee].push(score);
      }
    } else {
      unmatched++;
    }
  });

  if (unmatched) Logger.log(`팀원 평가 매핑 실패: ${unmatched}건`);

  const ws = ss.getSheetByName("최종점수");
  STUDENTS.forEach(([, name], i) => {
    const arr = scores[name] || [];
    const avg = arr.length ? Math.round(arr.reduce((a,b)=>a+b,0)/arr.length*100)/100 : "";
    ws.getRange(6 + i, 3).setValue(avg);
  });
}

function updateTeamScores(ss, submissions) {
  const scores = {};
  let unmatched = 0;

  submissions.forEach(sub => {
    const f         = sub.fields || [];
    const targetRaw = findValue(f, TEAM_FIELD_TARGET);
    const evalRaw   = findValue(f, TEAM_FIELD_EVALUATOR);
    const scoreRaw  = findValue(f, TEAM_FIELD_SCORE);

    if (targetRaw && scoreRaw !== null && scoreRaw !== "") {
      const target    = parseInt(targetRaw);
      const evaluator = evalRaw ? parseInt(evalRaw) : null;
      const score     = parseFloat(scoreRaw);
      if (!isNaN(target) && !isNaN(score) && evaluator !== target) {
        scores[target] = scores[target] || [];
        scores[target].push(score);
      }
    } else {
      unmatched++;
    }
  });

  if (unmatched) Logger.log(`팀별 평가 매핑 실패: ${unmatched}건`);

  const ws = ss.getSheetByName("팀점수입력");
  for (let i = 1; i <= 9; i++) {
    const arr = scores[i] || [];
    const avg = arr.length ? Math.round(arr.reduce((a,b)=>a+b,0)/arr.length*100)/100 : "";
    ws.getRange(3 + i, 6).setValue(avg);
  }
}

// ─── 메인 동기화 ──────────────────────────────────
function syncAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Tally에서 데이터 가져오는 중...", "📊 동기화", 60);

  try {
    const peerSubs = fetchSubmissions(FORM_PEER_EVAL);
    saveRawPeer(ss, peerSubs);
    updatePeerScores(ss, peerSubs);

    const teamSubs = fetchSubmissions(FORM_TEAM_EVAL);
    saveRawTeam(ss, teamSubs);
    updateTeamScores(ss, teamSubs);

    ss.toast(
      `팀원 평가 ${peerSubs.length}건 · 팀별 평가 ${teamSubs.length}건 완료`,
      "✅ 동기화 완료", 5
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(`오류:\n${e.message}`);
  }
}

// ─── Tally 필드명 확인 ────────────────────────────
function inspectFields() {
  const ui = SpreadsheetApp.getUi();
  let report = "각 폼의 실제 질문 제목 목록입니다.\n상단 설정의 키워드와 비교해 수정하세요.\n";

  for (const [formId, formName] of [[FORM_PEER_EVAL,"팀원 평가표"],[FORM_TEAM_EVAL,"팀별 평가표"]]) {
    report += `\n▶ ${formName}\n`;
    const subs = fetchSubmissions(formId);
    if (!subs.length) { report += "  응답 없음\n"; continue; }
    report += `  응답 ${subs.length}건\n`;
    (subs[0].fields || []).forEach(f => {
      report += `  • "${f.label}" → ${JSON.stringify(f.value)}\n`;
    });
  }

  ui.alert("Tally 필드 목록", report, ui.ButtonSet.OK);
}

// ─── 자동 동기화 트리거 ───────────────────────────
function setupTrigger() {
  removeTriggers();
  ScriptApp.newTrigger("syncAll").timeBased().everyMinutes(10).create();
  SpreadsheetApp.getUi().alert("✅ 10분마다 자동 동기화가 설정되었습니다.");
}

function removeTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "syncAll")
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ─── 유틸 ─────────────────────────────────────────
function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader(ws, numCols) {
  ws.getRange(1, 1, 1, numCols)
    .setBackground("#D9E8FB")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}
