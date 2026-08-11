/*******************************************************************
 * 아미쿠스 한국학교 체크인 백엔드 v2 (Google Apps Script)
 * 학교 계정(amicuskoreanschool@gmail.com)으로 배포하세요.
 *
 * 기능:
 *  1) GET ?action=roster : '학생들' 탭에서 명단(한글이름·성별·등록수업)
 *  2) GET ?action=today  : 오늘의 체크인/체크아웃 기록 (모든 기기 동기화)
 *  3) POST {action:'record'} : 기록 저장 → 시트 '체크인기록' 탭 + Notion
 *  4) 금요일 오후 9시 자동 요약 : 당일 반별 출석 정리 → 이메일 발송 + Notion
 *
 * 설정 (한 번만):
 *  A. "한국학교 check in/out" 시트 열기 → 확장 프로그램 → Apps Script → 붙여넣기
 *  B. Notion 연동(선택):
 *     1. www.notion.so/profile/integrations 에서 연동 생성(학교 워크스페이스)
 *     2. 프로젝트 설정 → 스크립트 속성:
 *          NOTION_TOKEN  = ntn_... 또는 secret_... 시크릿
 *          NOTION_PARENT = 기록을 둘 Notion 페이지 링크
 *        (그 Notion 페이지 ··· → 연결에 만든 연동 추가)
 *     3. setupNotionDatabase 함수 1회 실행 → "한국학교 체크인 기록" DB 자동 생성
 *  C. 요약 이메일 받는 주소(선택): 스크립트 속성 SUMMARY_EMAIL (쉼표로 여러 명)
 *     기본값: 이 계정(amicuskoreanschool@gmail.com)
 *  D. installFridayTrigger 함수 1회 실행 → 매주 금요일 오후 9시 자동 요약
 *  E. 배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자)
 *     → URL을 index.html 의 CONFIG.API_URL 에 붙여넣기
 *******************************************************************/

const SPREADSHEET_ID = "1_2wxyi2SrCtqO9LxQcJDAw8rQ1U6uRpG5UHgBON66cE";
const ROSTER_SHEET = "학생들";
const LOG_SHEET = "체크인기록";

/* ---------------- 명단 ---------------- */
function readRoster() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ROSTER_SHEET) || ss.getSheets()[0];
  const values = sheet.getDataRange().getValues();
  let hr = -1, cName = -1, cGender = -1, cCls = -1;
  for (let r = 0; r < Math.min(values.length, 10) && hr < 0; r++) {
    values[r].forEach((h, c) => {
      if (String(h).includes("한글이름")) { hr = r; }
    });
  }
  if (hr < 0) hr = 0;
  values[hr].forEach((h, c) => {
    const v = String(h);
    if (cName < 0 && v.includes("한글이름")) cName = c;
    if (cGender < 0 && v.includes("성별")) cGender = c;
    if (cCls < 0 && (v.includes("등록수업") || v.includes("수업") || v === "반")) cCls = c;
  });
  const roster = [];
  for (let r = hr + 1; r < values.length; r++) {
    const name = String(values[r][cName] || "").trim();
    if (!name) continue;
    roster.push({
      name: name,
      gender: String(values[r][cGender] || "").trim(),
      cls: String(values[r][cCls] || "").trim(),
    });
  }
  return roster;
}

/* ---------------- 기록 시트 ---------------- */
function logSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.appendRow(["날짜", "시간", "이름", "성별", "등록수업", "구분", "기록자", "ISO시각"]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function todayStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}
function readTodayLog() {
  const sh = logSheet_();
  const values = sh.getDataRange().getValues();
  const today = todayStr_();
  const log = [];
  for (let r = 1; r < values.length; r++) {
    const [d, time, name, gender, cls, gubun, by, iso] = values[r];
    const dStr = (d instanceof Date)
      ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d);
    if (dStr !== today) continue;
    log.push({
      id: name + "|" + cls, name: String(name), gender: String(gender), cls: String(cls),
      type: gubun === "등원" ? "in" : (gubun === "하원" ? "out" : "none"),
      by: String(by || "kiosk"), ts: String(iso || ""),
    });
  }
  return log;
}

/* ---------------- 웹 API ---------------- */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "";
  let payload;
  try {
    if (action === "roster") payload = { ok: true, roster: readRoster() };
    else if (action === "today") payload = { ok: true, log: readTodayLog() };
    else payload = { ok: true, service: "한국학교 체크인 API" };
  } catch (err) { payload = { ok: false, error: String(err) }; }
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const rec = JSON.parse(e.postData.contents);
  const t = rec.ts ? new Date(rec.ts) : new Date();
  const tz = Session.getScriptTimeZone();
  logSheet_().appendRow([
    Utilities.formatDate(t, tz, "yyyy-MM-dd"),
    Utilities.formatDate(t, tz, "HH:mm:ss"),
    rec.name, rec.gender, rec.cls,
    rec.type === "in" ? "등원" : (rec.type === "out" ? "하원" : "취소"),
    rec.by === "teacher" ? "선생님" : "키오스크",
    t.toISOString(),
  ]);
  let notion = "skipped";
  try { notion = pushToNotion(rec, t); } catch (err) { notion = "error: " + err; }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, notion: notion }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Notion ---------------- */
function notionHeaders_() {
  const token = PropertiesService.getScriptProperties().getProperty("NOTION_TOKEN");
  if (!token) return null;
  return { "Authorization": "Bearer " + token, "Notion-Version": "2022-06-28", "Content-Type": "application/json" };
}
function pushToNotion(rec, t) {
  const headers = notionHeaders_();
  const dbId = PropertiesService.getScriptProperties().getProperty("NOTION_DB_ID");
  if (!headers || !dbId) return "not configured";
  const body = {
    parent: { database_id: dbId },
    properties: {
      "이름": { title: [{ text: { content: rec.name } }] },
      "성별": { select: { name: rec.gender || "미상" } },
      "등록수업": { select: { name: rec.cls || "미지정" } },
      "구분": { select: { name: rec.type === "in" ? "등원" : (rec.type === "out" ? "하원" : "취소") } },
      "기록자": { select: { name: rec.by === "teacher" ? "선생님" : "키오스크" } },
      "시간": { date: { start: t.toISOString() } },
    },
  };
  const res = UrlFetchApp.fetch("https://api.notion.com/v1/pages",
    { method: "post", headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true });
  return res.getResponseCode() === 200 ? "ok" : res.getContentText();
}
function setupNotionDatabase() {
  const headers = notionHeaders_();
  if (!headers) throw new Error("스크립트 속성 NOTION_TOKEN 을 먼저 넣어주세요");
  const parent = PropertiesService.getScriptProperties().getProperty("NOTION_PARENT");
  if (!parent) throw new Error("스크립트 속성 NOTION_PARENT(페이지 링크)를 넣어주세요");
  const m = parent.replace(/-/g, "").match(/[0-9a-f]{32}/i);
  if (!m) throw new Error("NOTION_PARENT에서 페이지 ID를 찾지 못했습니다");
  const body = {
    parent: { type: "page_id", page_id: m[0] },
    title: [{ text: { content: "한국학교 체크인 기록" } }],
    properties: {
      "이름": { title: {} },
      "성별": { select: { options: [{ name: "남", color: "blue" }, { name: "여", color: "pink" }] } },
      "등록수업": { select: {} },
      "구분": { select: { options: [{ name: "등원", color: "green" }, { name: "하원", color: "yellow" }, { name: "취소", color: "gray" }] } },
      "기록자": { select: { options: [{ name: "키오스크", color: "default" }, { name: "선생님", color: "purple" }] } },
      "시간": { date: {} },
    },
  };
  const res = UrlFetchApp.fetch("https://api.notion.com/v1/databases",
    { method: "post", headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());
  if (!data.id) throw new Error("생성 실패: " + res.getContentText());
  PropertiesService.getScriptProperties().setProperty("NOTION_DB_ID", data.id);
  Logger.log("Notion DB 생성 완료: " + data.id);
}

/* ---------------- 금요일 자동 요약 ---------------- */
function buildSummary_() {
  const roster = readRoster();
  const log = readTodayLog();
  const last = {};
  log.forEach(r => { last[r.id] = r; });
  const classes = [...new Set(roster.map(s => s.cls))];
  const lines = [];
  let tIn = 0, tOut = 0, tNone = 0;
  classes.forEach(c => {
    const kids = roster.filter(s => s.cls === c);
    const stIn = [], stOut = [], stNone = [];
    kids.forEach(s => {
      const st = last[s.name + "|" + s.cls];
      if (st && st.type === "in") stIn.push(s.name);
      else if (st && st.type === "out") stOut.push(s.name);
      else stNone.push(s.name);
    });
    tIn += stIn.length; tOut += stOut.length; tNone += stNone.length;
    lines.push(
      "■ " + c + "  (등원 " + stIn.length + " / 하원 " + stOut.length + " / 미등원 " + stNone.length + ")\n" +
      (stIn.length ? "  · 등원: " + stIn.join(", ") + "\n" : "") +
      (stOut.length ? "  · 하원: " + stOut.join(", ") + "\n" : "") +
      (stNone.length ? "  · 미등원: " + stNone.join(", ") + "\n" : "")
    );
  });
  const today = todayStr_();
  const head = "[아미쿠스 한국학교] " + today + " 출석 요약\n" +
    "전체 " + roster.length + "명 — 등원 " + tIn + " / 하원 " + tOut + " / 미등원 " + tNone + "\n\n";
  return { subject: "[한국학교] " + today + " 출석 요약 (등원 " + tIn + "명)", body: head + lines.join("\n"), today: today };
}

function fridaySummary() {
  const s = buildSummary_();
  const to = PropertiesService.getScriptProperties().getProperty("SUMMARY_EMAIL")
          || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(to, s.subject, s.body);
  // Notion 요약 페이지 (설정된 경우)
  try {
    const headers = notionHeaders_();
    const parent = PropertiesService.getScriptProperties().getProperty("NOTION_PARENT");
    if (headers && parent) {
      const m = parent.replace(/-/g, "").match(/[0-9a-f]{32}/i);
      if (m) {
        UrlFetchApp.fetch("https://api.notion.com/v1/pages", {
          method: "post", headers: headers, muteHttpExceptions: true,
          payload: JSON.stringify({
            parent: { type: "page_id", page_id: m[0] },
            properties: { title: { title: [{ text: { content: "📋 " + s.today + " 출석 요약" } }] } },
            children: s.body.split("\n").filter(Boolean).slice(0, 90).map(line => ({
              object: "block", type: "paragraph",
              paragraph: { rich_text: [{ text: { content: line } }] },
            })),
          }),
        });
      }
    }
  } catch (e) { Logger.log("notion summary error: " + e); }
}

/* 1회 실행: 매주 금요일 오후 9시 자동 요약 트리거 설치 */
function installFridayTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "fridaySummary") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("fridaySummary")
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(21).create();
  Logger.log("금요일 오후 9시 트리거 설치 완료");
}

/* 테스트 */
function testRoster() { Logger.log(readRoster().length + "명"); }
function testSummary() { Logger.log(buildSummary_().body); }
