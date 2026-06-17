const TALLY_TOKEN    = "tly-xHfYqWNwxaATbpbg0GHsLlPBP5LKjb1P";
const SPREADSHEET_ID = "1tqihK7NxqEEqohQqFlDsjNMXz531K1cpCtwRvni2uk8";
const FORM_PEER_EVAL = "VLNBql";
const FORM_TEAM_EVAL = "q4zGOO";

const PEER_FIELD_EVALUATEE = "평가 대상";
const PEER_FIELD_SCORE     = "점수";
const TEAM_FIELD_TARGET    = "평가 조";
const TEAM_FIELD_EVALUATOR = "본인 조";
const TEAM_FIELD_SCORE     = "점수";

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
  } catch(e) {}
}

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

function findValue(fields, keyword) {
  var kw = keyword.toLowerCase();
  for (var i = 0; i < fields.length; i++) {
    var f = fields[i];
    if (f.label && f.label.toLowerCase().indexOf(kw) !== -1) {
      var val = f.value;
      return Array.isArray(val) ? (val[0] !== undefined ? val[0] : null) : val;
    }
  }
  return null;
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader(ws, numCols) {
  ws.getRange(1, 1, 1, numCols)
    .setBackground("#D9E8FB")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

function saveRawPeer(ss, submissions) {
  var ws   = getOrCreateSheet(ss, "팀원평가_로우");
  var rows = [["번호","제출시간","피평가자","점수 (1~5)"]];
  for (var i = 0; i < submissions.length; i++) {
    var f = submissions[i].fields || [];
    rows.push([
      i + 1,
      (submissions[i].createdAt || "").replace("T"," ").slice(0,19),
      findValue(f, PEER_FIELD_EVALUATEE) || "",
      findValue(f, PEER_FIELD_SCORE)     || "",
    ]);
  }
  ws.clearContents();
  ws.getRange(1, 1, rows.length, 4).setValues(rows);
  styleHeader(ws, 4);
  Logger.log("팀원평가_로우 저장: " + submissions.length + "건");
}

function saveRawTeam(ss, submissions) {
  var ws   = getOrCreateSheet(ss, "팀별평가_로우");
  var rows = [["번호","제출시간","평가자 조","평가 대상 조","점수 (1~5)","비고"]];
  for (var i = 0; i < submissions.length; i++) {
    var f         = submissions[i].fields || [];
    var target    = findValue(f, TEAM_FIELD_TARGET)    || "";
    var evaluator = findValue(f, TEAM_FIELD_EVALUATOR) || "";
    var score     = findValue(f, TEAM_FIELD_SCORE)     || "";
    rows.push([
      i + 1,
      (submissions[i].createdAt || "").replace("T"," ").slice(0,19),
      evaluator, target, score,
      String(evaluator) === String(target) ? "자기평가 제외" : "",
    ]);
  }
  ws.clearContents();
  ws.getRange(1, 1, rows.length, 6).setValues(rows);
  styleHeader(ws, 6);
  Logger.log("팀별평가_로우 저장: " + submissions.length + "건");
}

function updatePeerScores(ss, submissions) {
  var scores = {};
  var unmatched = 0;
  for (var i = 0; i < submissions.length; i++) {
    var f         = submissions[i].fields || [];
    var evaluatee = findValue(f, PEER_FIELD_EVALUATEE);
    var scoreRaw  = findValue(f, PEER_FIELD_SCORE);
    if (evaluatee && scoreRaw !== null && scoreRaw !== "") {
      var score = parseFloat(scoreRaw);
      if (!isNaN(score)) {
        if (!scores[evaluatee]) scores[evaluatee] = [];
        scores[evaluatee].push(score);
      }
    } else { unmatched++; }
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
    var f         = submissions[i].fields || [];
    var targetRaw = findValue(f, TEAM_FIELD_TARGET);
    var evalRaw   = findValue(f, TEAM_FIELD_EVALUATOR);
    var scoreRaw  = findValue(f, TEAM_FIELD_SCORE);
    if (targetRaw && scoreRaw !== null && scoreRaw !== "") {
      var target    = parseInt(targetRaw);
      var evaluator = evalRaw ? parseInt(evalRaw) : null;
      var score     = parseFloat(scoreRaw);
      if (!isNaN(target) && !isNaN(score) && evaluator !== target) {
        if (!scores[target]) scores[target] = [];
        scores[target].push(score);
      }
    } else { unmatched++; }
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

function inspectFields() {
  var report = "각 폼의 실제 질문 제목 목록\n";
  var forms = [[FORM_PEER_EVAL,"팀원 평가표"],[FORM_TEAM_EVAL,"팀별 평가표"]];
  for (var fi = 0; fi < forms.length; fi++) {
    var formId   = forms[fi][0];
    var formName = forms[fi][1];
    report += "\n▶ " + formName + "\n";
    var subs = fetchSubmissions(formId);
    if (!subs.length) { report += "  응답 없음\n"; continue; }
    report += "  응답 " + subs.length + "건\n";
    var fields = subs[0].fields || [];
    for (var i = 0; i < fields.length; i++) {
      report += '  • "' + fields[i].label + '" → ' + JSON.stringify(fields[i].value) + "\n";
    }
  }
  SpreadsheetApp.getUi().alert("Tally 필드 목록", report, SpreadsheetApp.getUi().ButtonSet.OK);
}

function setupTrigger() {
  removeTriggers();
  ScriptApp.newTrigger("syncAll").timeBased().everyMinutes(10).create();
  SpreadsheetApp.getUi().alert("✅ 10분마다 자동 동기화 설정 완료");
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "syncAll") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
