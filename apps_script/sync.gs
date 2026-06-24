var TALLY_TOKEN    = "tly-xHfYqWNwxaATbpbg0GHsLlPBP5LKjb1P";
var SPREADSHEET_ID = "1tqihK7NxqEEqohQqFlDsjNMXz531K1cpCtwRvni2uk8";
var FORM_PEER_EVAL = "VLNBql";
var FORM_TEAM_EVAL = "q4zGOO";

// 팀원 평가표 - 건너뛸 questionId (본인 조, 학번 등 피평가자가 아닌 배열 응답)
var PEER_SKIP_QIDS = ["XEx7B4", "yxPbq4"];

// 팀별 평가표 - 본인 조 questionId
var TEAM_QID_EVALUATOR = "XExyPL";

var STUDENTS = [
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

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("📊 점수 동기화")
      .addItem("🔄 지금 동기화", "syncAll")
      .addSeparator()
      .addItem("🔍 응답 구조 확인", "logFirstResponse")
      .addSeparator()
      .addItem("⏰ 자동 동기화 켜기 (10분마다)", "setupTrigger")
      .addItem("⏹ 자동 동기화 끄기", "removeTriggers")
      .addToUi();
  } catch(e) {}
}

// ─── Tally API ────────────────────────────────────────────────
function fetchSubmissions(formId) {
  var allSubs = [];
  var page = 1;
  while (true) {
    var url  = "https://api.tally.so/forms/" + formId + "/submissions?limit=200&page=" + page;
    var resp = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: { Authorization: "Bearer " + TALLY_TOKEN },
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() !== 200) {
      throw new Error("Tally API 오류 (" + resp.getResponseCode() + "): " + resp.getContentText());
    }
    var data = JSON.parse(resp.getContentText());
    var subs = (data.submissions || []).filter(function(s){ return s.isCompleted; });
    allSubs = allSubs.concat(subs);
    if (!data.hasMore) break;
    page++;
  }
  return allSubs;
}

// questionId로 응답 값 찾기
function findById(responses, qid) {
  for (var i = 0; i < responses.length; i++) {
    if (responses[i].questionId === qid) {
      var ans = responses[i].answer;
      if (Array.isArray(ans)) return ans.length > 0 ? ans[0] : null;
      return ans;
    }
  }
  return null;
}

// 팀 이름에서 조번호 추출 (예: "2조 (윤동규...)" → 2)
function extractTeamNumber(teamStr) {
  if (!teamStr) return null;
  var m = String(teamStr).match(/^(\d+)조/);
  return m ? parseInt(m[1]) : null;
}

// ─── 팀원 평가 파싱: 배열 응답 = 피평가자, 이후 숫자 = 점수 ──
// 한 제출에 여러 팀원을 동시에 평가하는 구조에 대응
function parsePeerPairs(responses) {
  var pairs = [];
  var currentEvaluatee = null;
  var currentScores = [];

  for (var i = 0; i < responses.length; i++) {
    var qid = responses[i].questionId;
    var ans = responses[i].answer;

    // 본인 조, 학번 등 건너뜀
    var skip = false;
    for (var k = 0; k < PEER_SKIP_QIDS.length; k++) {
      if (qid === PEER_SKIP_QIDS[k]) { skip = true; break; }
    }
    if (skip) continue;

    if (Array.isArray(ans) && ans.length > 0) {
      // 배열 응답 = 피평가자 이름 선택
      if (currentEvaluatee !== null && currentScores.length > 0) {
        pairs.push({
          evaluatee: currentEvaluatee,
          avg: Math.round(currentScores.reduce(function(a,b){return a+b;},0)/currentScores.length*100)/100,
        });
      }
      currentEvaluatee = ans[0];
      currentScores = [];
    } else if (typeof ans === "number" || (typeof ans === "string" && ans !== "" && !isNaN(parseFloat(ans)))) {
      // 숫자 응답 = 점수
      if (currentEvaluatee !== null) {
        currentScores.push(parseFloat(ans));
      }
    }
  }

  // 마지막 피평가자 처리
  if (currentEvaluatee !== null && currentScores.length > 0) {
    pairs.push({
      evaluatee: currentEvaluatee,
      avg: Math.round(currentScores.reduce(function(a,b){return a+b;},0)/currentScores.length*100)/100,
    });
  }

  return pairs;
}

// ─── 팀별 평가 파싱: 배열 응답 = 평가 대상 조, 이후 숫자 = 점수 ──
function parseTeamPairs(responses) {
  var pairs = [];
  var currentTarget = null;
  var currentScores = [];

  for (var i = 0; i < responses.length; i++) {
    var qid = responses[i].questionId;
    var ans = responses[i].answer;

    if (qid === TEAM_QID_EVALUATOR) continue;

    if (Array.isArray(ans) && ans.length > 0) {
      if (currentTarget !== null && currentScores.length > 0) {
        pairs.push({
          target: currentTarget,
          avg: Math.round(currentScores.reduce(function(a,b){return a+b;},0)/currentScores.length*100)/100,
        });
      }
      currentTarget = extractTeamNumber(ans[0]);
      currentScores = [];
    } else if (typeof ans === "number" || (typeof ans === "string" && ans !== "" && !isNaN(parseFloat(ans)))) {
      if (currentTarget !== null) {
        currentScores.push(parseFloat(ans));
      }
    }
  }

  if (currentTarget !== null && currentScores.length > 0) {
    pairs.push({
      target: currentTarget,
      avg: Math.round(currentScores.reduce(function(a,b){return a+b;},0)/currentScores.length*100)/100,
    });
  }

  return pairs;
}

// ─── Google Sheets 유틸 ───────────────────────────────────────
function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader(ws, numCols) {
  ws.getRange(1, 1, 1, numCols)
    .setBackground("#D9E8FB")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

// ─── 로우 데이터 저장 ─────────────────────────────────────────
function saveRawPeer(ss, submissions) {
  var ws   = getOrCreateSheet(ss, "팀원평가_로우");
  var rows = [["번호","제출시간","피평가자","평균점수 (1~5)"]];
  for (var i = 0; i < submissions.length; i++) {
    var r     = submissions[i].responses || [];
    var pairs = parsePeerPairs(r);
    var ts    = (submissions[i].createdAt || "").replace("T"," ").slice(0,19);
    if (pairs.length === 0) {
      rows.push([i+1, ts, "?", ""]);
    }
    for (var p = 0; p < pairs.length; p++) {
      rows.push([i+1, ts, pairs[p].evaluatee, pairs[p].avg]);
    }
  }
  ws.clearContents();
  ws.getRange(1, 1, rows.length, 4).setValues(rows);
  styleHeader(ws, 4);
  Logger.log("팀원평가_로우 저장: " + submissions.length + "건");
}

function saveRawTeam(ss, submissions) {
  var ws   = getOrCreateSheet(ss, "팀별평가_로우");
  var rows = [["번호","제출시간","평가자 조","평가 대상 조","점수 (평균)","비고"]];
  for (var i = 0; i < submissions.length; i++) {
    var r            = submissions[i].responses || [];
    var evaluatorStr = findById(r, TEAM_QID_EVALUATOR) || "";
    var evaluator    = extractTeamNumber(evaluatorStr);
    var pairs        = parseTeamPairs(r);
    var ts           = (submissions[i].createdAt || "").replace("T"," ").slice(0,19);
    if (pairs.length === 0) {
      rows.push([i+1, ts, evaluatorStr, "?", "", "파싱실패"]);
    }
    for (var p = 0; p < pairs.length; p++) {
      rows.push([
        i+1, ts, evaluatorStr,
        pairs[p].target + "조",
        pairs[p].avg,
        evaluator === pairs[p].target ? "자기평가 제외" : "",
      ]);
    }
  }
  ws.clearContents();
  ws.getRange(1, 1, rows.length, 6).setValues(rows);
  styleHeader(ws, 6);
  Logger.log("팀별평가_로우 저장: " + submissions.length + "건");
}

// ─── 집계 및 시트 업데이트 ────────────────────────────────────
function updatePeerScores(ss, submissions) {
  var scores = {};
  var unmatched = 0;

  for (var i = 0; i < submissions.length; i++) {
    var r     = submissions[i].responses || [];
    var pairs = parsePeerPairs(r);
    if (pairs.length === 0) { unmatched++; continue; }
    for (var p = 0; p < pairs.length; p++) {
      var name = String(pairs[p].evaluatee);
      if (!scores[name]) scores[name] = [];
      scores[name].push(pairs[p].avg);
    }
  }

  if (unmatched) Logger.log("팀원 평가 매핑 실패: " + unmatched + "건");

  var ws = ss.getSheetByName("최종점수");
  for (var i = 0; i < STUDENTS.length; i++) {
    var name = STUDENTS[i][1];
    var arr  = scores[name] || [];
    var avg  = arr.length ? Math.round(arr.reduce(function(a,b){return a+b;},0)/arr.length*100)/100 : "";
    ws.getRange(6 + i, 3).setValue(avg);
  }
  Logger.log("최종점수 C열 업데이트 완료");
}

function updateTeamScores(ss, submissions) {
  var scores = {};
  var unmatched = 0;

  for (var i = 0; i < submissions.length; i++) {
    var r            = submissions[i].responses || [];
    var evaluatorStr = findById(r, TEAM_QID_EVALUATOR);
    var evaluator    = extractTeamNumber(evaluatorStr);
    var pairs        = parseTeamPairs(r);

    if (pairs.length === 0) { unmatched++; continue; }

    for (var p = 0; p < pairs.length; p++) {
      var target = pairs[p].target;
      if (target === null || evaluator === target) continue;
      if (!scores[target]) scores[target] = [];
      scores[target].push(pairs[p].avg);
    }
  }

  if (unmatched) Logger.log("팀별 평가 매핑 실패: " + unmatched + "건");

  var ws = ss.getSheetByName("팀점수입력");
  for (var i = 1; i <= 9; i++) {
    var arr = scores[i] || [];
    var avg = arr.length ? Math.round(arr.reduce(function(a,b){return a+b;},0)/arr.length*100)/100 : "";
    ws.getRange(3 + i, 6).setValue(avg);
  }
  Logger.log("팀점수입력 F열 업데이트 완료");
}

// ─── 메인 동기화 ─────────────────────────────────────────────
function syncAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
        || SpreadsheetApp.openById(SPREADSHEET_ID);

  try { ss.toast("Tally에서 데이터 가져오는 중...", "📊 동기화", 60); } catch(e) {}
  Logger.log("=== 동기화 시작 ===");

  var peerSubs = fetchSubmissions(FORM_PEER_EVAL);
  Logger.log("팀원 평가표: " + peerSubs.length + "건");
  saveRawPeer(ss, peerSubs);
  updatePeerScores(ss, peerSubs);

  var teamSubs = fetchSubmissions(FORM_TEAM_EVAL);
  Logger.log("팀별 평가표: " + teamSubs.length + "건");
  saveRawTeam(ss, teamSubs);
  updateTeamScores(ss, teamSubs);

  Logger.log("=== 동기화 완료 ===");
  try { ss.toast("완료!", "✅", 5); } catch(e) {}
}

// ─── 디버그: 첫 응답 구조 확인 ───────────────────────────────
function logFirstResponse() {
  var forms = [[FORM_PEER_EVAL,"팀원 평가표"],[FORM_TEAM_EVAL,"팀별 평가표"]];
  for (var fi = 0; fi < forms.length; fi++) {
    Logger.log("▶ " + forms[fi][1]);
    var subs = fetchSubmissions(forms[fi][0]);
    if (!subs.length) { Logger.log("  응답 없음"); continue; }
    Logger.log("  총 " + subs.length + "건");
    Logger.log(JSON.stringify(subs[0].responses, null, 2));
  }
}

// ─── 트리거 설정 ─────────────────────────────────────────────
function setupTrigger() {
  removeTriggers();
  ScriptApp.newTrigger("syncAll").timeBased().everyMinutes(10).create();
  SpreadsheetApp.getUi().alert("✅ 10분마다 자동 동기화 설정 완료");
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "syncAll") ScriptApp.deleteTrigger(triggers[i]);
  }
}
