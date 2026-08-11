/*******************************************************************
 * 아미쿠스 한국학교 통합 백엔드 v3 (Google Apps Script)
 *  - 부모님 체크인/체크아웃 (키오스크)
 *  - 선생님 포털: 대시보드/명단수정/학습현황(숙제·받아쓰기)/시험성적/주간마무리
 *  - 관리자 포털: 전체 반
 *  - Notion 자동 기록 + 금요일 오후 9시 자동 요약 이메일
 *******************************************************************/

const SPREADSHEET_ID = "1_2wxyi2SrCtqO9LxQcJDAw8rQ1U6uRpG5UHgBON66cE";
const ROSTER_SHEET = "학생들";
const TEACHER_SHEET = "선생님들";
const LOG_SHEET = "체크인기록";
const LEARN_SHEET = "학습현황";
const EXAM_SHEET = "시험성적";
const WEEKLY_SHEET = "주간마무리";
const REASON_SHEET = "결석사유";
const ADMIN_EMAIL = "amicuskoreanschool@gmail.com";
const PORTAL_URL = "https://amicuskoreanschool-ship-it.github.io/checkin/teacher.html";

function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function tz_() { return Session.getScriptTimeZone(); }
function todayStr_() { return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd"); }
function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), "yyyy-MM-dd");
  return String(v || "").trim();
}
function getOrCreate_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}

/* ================= 학생 명단 ================= */
function rosterHeaderMap_() {
  const sheet = ss_().getSheetByName(ROSTER_SHEET) || ss_().getSheets()[0];
  const values = sheet.getDataRange().getValues();
  let hr = -1;
  for (let r = 0; r < Math.min(values.length, 10) && hr < 0; r++) {
    values[r].forEach(function(h) { if (String(h).includes("한글이름")) hr = r; });
  }
  if (hr < 0) hr = 0;
  const map = {};
  values[hr].forEach(function(h, c) { map[String(h).trim()] = c; });
  return { sheet: sheet, values: values, hr: hr, map: map };
}
function ensureAllergyCol_() {
  const ctx = rosterHeaderMap_();
  if (ctx.map["알러지"] !== undefined) return ctx;
  const col = ctx.values[ctx.hr].length + 1;
  ctx.sheet.getRange(ctx.hr + 1, col).setValue("알러지");
  return rosterHeaderMap_();
}
function readStudentsFull_(cls) {
  const ctx = ensureAllergyCol_();
  const m = ctx.map, values = ctx.values;
  const col = function(row, key) { return m[key] !== undefined ? String(row[m[key]] || "").trim() : ""; };
  const out = [];
  for (let r = ctx.hr + 1; r < values.length; r++) {
    const row = values[r];
    const name = col(row, "한글이름");
    if (!name) continue;
    const c = col(row, "등록수업");
    if (cls && cls !== "ALL" && c !== cls) continue;
    let birth = m["생년월일"] !== undefined ? row[m["생년월일"]] : "";
    if (birth instanceof Date) birth = Utilities.formatDate(birth, tz_(), "M/d/yyyy");
    out.push({
      row: r + 1, name: name,
      eng: col(row, "영어이름"), birth: String(birth || "").trim(),
      gender: col(row, "성별"), grade: col(row, "학년"), cls: c,
      parent: col(row, "부모성함"), phone: col(row, "전화번호"),
      email: col(row, "이메일주소"), allergy: col(row, "알러지"),
    });
  }
  return out;
}
function readRoster() {
  return readStudentsFull_("ALL").map(function(s) {
    return { name: s.name, gender: s.gender, cls: s.cls };
  });
}
function updateStudent_(req) {
  const ctx = ensureAllergyCol_();
  const m = ctx.map, sheet = ctx.sheet;
  const allow = { name:"한글이름", eng:"영어이름", birth:"생년월일", grade:"학년",
                  parent:"부모성함", phone:"전화번호", email:"이메일주소", allergy:"알러지", cls:"등록수업" };
  const r = Number(req.row);
  if (!r || r <= ctx.hr + 1 - 1) throw new Error("잘못된 행 번호");
  Object.keys(req.fields || {}).forEach(function(k) {
    const h = allow[k];
    if (h === undefined || m[h] === undefined) return;
    sheet.getRange(r, m[h] + 1).setValue(String(req.fields[k]));
  });
  return "ok";
}

/* ================= 선생님 ================= */
function readTeachers_() {
  const sh = ss_().getSheetByName(TEACHER_SHEET);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let r = 0; r < values.length; r++) {
    const cls = String(values[r][0] || "").trim();
    const name = String(values[r][1] || "").trim();
    const email = String(values[r][2] || "").trim();
    if (cls && name && cls.includes("반")) out.push({ cls: cls, name: name, email: email });
  }
  return out;
}

/* ================= 체크인 기록 ================= */
function logSheet_() {
  return getOrCreate_(LOG_SHEET, ["날짜","시간","이름","성별","등록수업","구분","기록자","ISO시각"]);
}
function readTodayLog() {
  const values = logSheet_().getDataRange().getValues();
  const today = todayStr_();
  const log = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (fmtDate_(row[0]) !== today) continue;
    const gubun = row[5];
    log.push({
      id: row[2] + "|" + row[4], name: String(row[2]), gender: String(row[3]), cls: String(row[4]),
      type: gubun === "등원" ? "in" : (gubun === "하원" ? "out" : "none"),
      by: String(row[6] || "키오스크"), ts: String(row[7] || ""),
    });
  }
  return log;
}

/* 날짜별 출석 이력 (학생별 마지막 상태 + 첫 등원 시각) */
function fmtTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), "HH:mm:ss");
  return String(v || "").trim();
}
function readAttendanceAll_(cls) {
  const values = logSheet_().getDataRange().getValues();
  const last = {};
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const c = String(row[4] || "").trim();
    if (cls && cls !== "ALL" && c !== cls) continue;
    const date = fmtDate_(row[0]);
    const gubun = row[5];
    const key = date + "|" + row[2] + "|" + c;
    if (!last[key]) last[key] = { date: date, name: String(row[2]), cls: c, type: "none", inAt: "" };
    const rec = last[key];
    if (gubun === "등원" && !rec.inAt) rec.inAt = fmtTime_(row[1]);
    rec.type = gubun === "등원" ? "in" : (gubun === "하원" ? "out" : "none");
  }
  return Object.keys(last).map(function(k){ return last[k]; });
}

/* ================= 학습현황 (숙제·받아쓰기) ================= */
function learnSheet_() {
  return getOrCreate_(LEARN_SHEET, ["날짜","반","이름","숙제","받아쓰기","기록자","수정시각"]);
}
function readLearning_(cls) {
  const values = learnSheet_().getDataRange().getValues();
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const c = String(row[1] || "").trim();
    if (cls && cls !== "ALL" && c !== cls) continue;
    out.push({ date: fmtDate_(row[0]), cls: c, name: String(row[2] || ""),
               hw: String(row[3] || ""), dict: String(row[4] || "") });
  }
  return out;
}
function setLearning_(req) {
  const sh = learnSheet_();
  const values = sh.getDataRange().getValues();
  const date = String(req.date), cls = String(req.cls), name = String(req.name);
  const now = Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd HH:mm");
  for (let r = 1; r < values.length; r++) {
    if (fmtDate_(values[r][0]) === date && String(values[r][1]).trim() === cls && String(values[r][2]).trim() === name) {
      sh.getRange(r+1, 4, 1, 4).setValues([[req.hw||"", req.dict||"", req.by||"", now]]);
      return "updated";
    }
  }
  sh.appendRow([date, cls, name, req.hw||"", req.dict||"", req.by||"", now]);
  return "added";
}

/* ================= 시험 성적 ================= */
function examSheet_() {
  return getOrCreate_(EXAM_SHEET, ["반","이름","중간시험","기말시험","수정시각"]);
}
function readExams_(cls) {
  const values = examSheet_().getDataRange().getValues();
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const c = String(row[0] || "").trim();
    if (cls && cls !== "ALL" && c !== cls) continue;
    out.push({ cls: c, name: String(row[1] || ""), mid: String(row[2] || ""), fin: String(row[3] || "") });
  }
  return out;
}
function setExam_(req) {
  const sh = examSheet_();
  const values = sh.getDataRange().getValues();
  const cls = String(req.cls), name = String(req.name);
  const now = Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd HH:mm");
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === cls && String(values[r][1]).trim() === name) {
      sh.getRange(r+1, 3, 1, 3).setValues([[req.mid||"", req.fin||"", now]]);
      return "updated";
    }
  }
  sh.appendRow([cls, name, req.mid||"", req.fin||"", now]);
  return "added";
}

/* ================= 주간 마무리 ================= */
function weeklySheet_() {
  return getOrCreate_(WEEKLY_SHEET, ["날짜","반","선생님","특이사항","제출시각"]);
}
function readWeekly_(cls) {
  const values = weeklySheet_().getDataRange().getValues();
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const c = String(row[1] || "").trim();
    if (cls && cls !== "ALL" && c !== cls) continue;
    out.push({ date: fmtDate_(row[0]), cls: c, teacher: String(row[2] || ""), note: String(row[3] || "") });
  }
  return out;
}
function setWeekly_(req) {
  const sh = weeklySheet_();
  const values = sh.getDataRange().getValues();
  const date = String(req.date), cls = String(req.cls);
  const now = Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd HH:mm");
  for (let r = 1; r < values.length; r++) {
    if (fmtDate_(values[r][0]) === date && String(values[r][1]).trim() === cls) {
      sh.getRange(r+1, 3, 1, 3).setValues([[req.teacher||"", req.note||"", now]]);
      return "updated";
    }
  }
  sh.appendRow([date, cls, req.teacher||"", req.note||"", now]);
  return "added";
}

/* ================= 결석 사유 ================= */
function reasonSheet_() {
  return getOrCreate_(REASON_SHEET, ["날짜","반","이름","결석사유","기록자","수정시각"]);
}
function readReasons_(cls) {
  const values = reasonSheet_().getDataRange().getValues();
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const c = String(row[1] || "").trim();
    if (cls && cls !== "ALL" && c !== cls) continue;
    out.push({ date: fmtDate_(row[0]), cls: c, name: String(row[2] || "").trim(),
               reason: String(row[3] || "") });
  }
  return out;
}
function setReason_(req) {
  const sh = reasonSheet_();
  const values = sh.getDataRange().getValues();
  const date = String(req.date), cls = String(req.cls), name = String(req.name);
  const now = Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd HH:mm");
  for (let r = 1; r < values.length; r++) {
    if (fmtDate_(values[r][0]) === date && String(values[r][1]).trim() === cls && String(values[r][2]).trim() === name) {
      sh.getRange(r+1, 4, 1, 3).setValues([[req.reason||"", req.by||"", now]]);
      return "updated";
    }
  }
  sh.appendRow([date, cls, name, req.reason||"", req.by||"", now]);
  return "added";
}

/* ================= 웹 API ================= */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || "";
  let payload;
  try {
    if (action === "roster") payload = { ok: true, roster: readRoster() };
    else if (action === "today") payload = { ok: true, log: readTodayLog() };
    else if (action === "teachers") payload = { ok: true, teachers: readTeachers_() };
    else if (action === "portal") {
      const cls = p.cls || "ALL";
      payload = { ok: true,
        students: readStudentsFull_(cls),
        log: readTodayLog(),
        att: readAttendanceAll_(cls),
        learning: readLearning_(cls),
        exams: readExams_(cls),
        weekly: readWeekly_(cls),
        reasons: readReasons_(cls),
        teachers: readTeachers_(),
      };
    }
    else payload = { ok: true, service: "한국학교 체크인 API v3" };
  } catch (err) { payload = { ok: false, error: String(err) }; }
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const req = JSON.parse(e.postData.contents);
  let result = {}, notion = "skipped";
  try {
    if (req.action === "record" || !req.action) {
      const t = req.ts ? new Date(req.ts) : new Date();
      logSheet_().appendRow([
        Utilities.formatDate(t, tz_(), "yyyy-MM-dd"),
        Utilities.formatDate(t, tz_(), "HH:mm:ss"),
        req.name, req.gender, req.cls,
        req.type === "in" ? "등원" : (req.type === "out" ? "하원" : "취소"),
        req.by === "teacher" ? "선생님" : "키오스크",
        t.toISOString(),
      ]);
      try { notion = pushToNotion(req, t); } catch (err) { notion = "error: " + err; }
      result = { ok: true, notion: notion };
    }
    else if (req.action === "updateStudent") result = { ok: true, r: updateStudent_(req) };
    else if (req.action === "setLearning") result = { ok: true, r: setLearning_(req) };
    else if (req.action === "setExam") result = { ok: true, r: setExam_(req) };
    else if (req.action === "setWeekly") result = { ok: true, r: setWeekly_(req) };
    else if (req.action === "setReason") result = { ok: true, r: setReason_(req) };
    else result = { ok: false, error: "unknown action" };
  } catch (err) { result = { ok: false, error: String(err) }; }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================= Notion ================= */
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
  if (!headers) throw new Error("NOTION_TOKEN 을 먼저 넣어주세요");
  const parent = PropertiesService.getScriptProperties().getProperty("NOTION_PARENT");
  if (!parent) throw new Error("NOTION_PARENT 를 넣어주세요");
  const m = parent.replace(/-/g, "").match(/[0-9a-f]{32}/i);
  const body = {
    parent: { type: "page_id", page_id: m[0] },
    title: [{ text: { content: "한국학교 체크인 기록" } }],
    properties: {
      "이름": { title: {} },
      "성별": { select: {} }, "등록수업": { select: {} },
      "구분": { select: {} }, "기록자": { select: {} }, "시간": { date: {} },
    },
  };
  const res = UrlFetchApp.fetch("https://api.notion.com/v1/databases",
    { method: "post", headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());
  if (!data.id) throw new Error("생성 실패: " + res.getContentText());
  PropertiesService.getScriptProperties().setProperty("NOTION_DB_ID", data.id);
}

/* ================= 금요일 자동 요약 (금 오후 10시) ================= */
function buildSummary_() {
  const roster = readRoster();
  const log = readTodayLog();
  const weekly = readWeekly_("ALL").filter(function(w){ return w.date === todayStr_(); });
  const last = {};
  log.forEach(function(r) { last[r.id] = r; });
  const classes = roster.map(function(s){return s.cls;}).filter(function(v,i,a){return a.indexOf(v)===i;}).sort();
  const rows = [];
  const lines = [];
  let tIn = 0, tNone = 0;
  classes.forEach(function(c) {
    const kids = roster.filter(function(s){return s.cls === c;});
    let nIn = 0; const stNone = [];
    kids.forEach(function(s) {
      const st = last[s.name + "|" + s.cls];
      if (st && (st.type === "in" || st.type === "out")) nIn++;
      else stNone.push(s.name);
    });
    tIn += nIn; tNone += stNone.length;
    const wk = weekly.filter(function(w){ return w.cls === c; })[0];
    rows.push({ cls: c, nIn: nIn, nNone: stNone.length, none: stNone, weekly: !!wk });
    lines.push("■ " + c + " — 등원 " + nIn + "명 / 미등원 " + stNone.length + "명" +
      (stNone.length ? " (" + stNone.join(", ") + ")" : "") +
      " · 주간마무리 " + (wk ? "제출" : "미제출"));
  });
  const today = todayStr_();
  const md = today.slice(5).replace("-", "/");
  const trs = rows.map(function(r){
    return "<tr><td style='padding:8px 14px;border-bottom:1px solid #F0E4D2;font-weight:700'>" + r.cls + "</td>" +
      "<td style='padding:8px 14px;border-bottom:1px solid #F0E4D2;text-align:center;color:#2E7D32;font-weight:700'>" + r.nIn + "</td>" +
      "<td style='padding:8px 14px;border-bottom:1px solid #F0E4D2;text-align:center;color:#C62828;font-weight:700'>" + r.nNone + "</td>" +
      "<td style='padding:8px 14px;border-bottom:1px solid #F0E4D2;text-align:center'>" + (r.weekly ? "✅ 제출" : "⏳ 미제출") + "</td></tr>";
  }).join("");
  const html =
    "<div style='font-family:Apple SD Gothic Neo,Malgun Gothic,sans-serif;max-width:640px;margin:0 auto;background:#FFF6EA;border-radius:16px;padding:26px'>" +
    "<h2 style='color:#E85D04;margin:0 0 4px'>🏫 아미쿠스 한국학교 · " + md + " (금) 주간 요약</h2>" +
    "<p style='font-size:16px;margin:14px 0'>전체 <b>" + roster.length + "명</b> 중 " +
    "<span style='color:#2E7D32;font-weight:800'>등원 " + tIn + "명</span> · " +
    "<span style='color:#C62828;font-weight:800'>미등원 " + tNone + "명</span></p>" +
    "<table style='border-collapse:collapse;width:100%;background:#FFFDF8;border-radius:12px;overflow:hidden;font-size:14px'>" +
    "<tr style='background:#F6E7CF'><th style='padding:9px 14px;text-align:left'>반</th><th style='padding:9px 14px'>등원</th><th style='padding:9px 14px'>미등원</th><th style='padding:9px 14px'>주간마무리</th></tr>" +
    trs + "</table>" +
    "<p style='color:#8A7A66;font-size:12.5px;margin-top:18px'>이 메일은 매주 금요일 오후 10시에 자동 발송됩니다 · <a href='" + PORTAL_URL + "'>선생님 포털 열기</a></p></div>";
  return {
    subject: "[아미쿠스 한국학교] " + md + " 주간 요약 — 등원 " + tIn + "명 / 미등원 " + tNone + "명",
    body: "아미쿠스 한국학교 " + today + " 주간 요약\n전체 " + roster.length + "명 — 등원 " + tIn + " / 미등원 " + tNone + "\n\n" + lines.join("\n"),
    html: html, today: today, weekly: weekly,
  };
}
function sendTeacherReminders_(today, weekly) {
  const teachers = readTeachers_();
  const byEmail = {};
  teachers.forEach(function(t) {
    if (!t.email) return;
    const done = weekly.some(function(w){ return w.cls === t.cls; });
    if (done) return;
    if (!byEmail[t.email]) byEmail[t.email] = { name: t.name, classes: [] };
    byEmail[t.email].classes.push(t.cls);
  });
  let sent = 0;
  Object.keys(byEmail).forEach(function(email) {
    const info = byEmail[email];
    const nm = info.name.replace(/\s*선생님\s*$/, "");
    const md = today.slice(5).replace("-", "/");
    const html =
      "<div style='font-family:Apple SD Gothic Neo,Malgun Gothic,sans-serif;max-width:560px;margin:0 auto;background:#FFF6EA;border-radius:16px;padding:26px'>" +
      "<h2 style='color:#E85D04;margin:0 0 12px'>📝 " + md + " 주간 마무리 안내</h2>" +
      "<p style='font-size:16px;line-height:1.7'><b>" + nm + " 선생님</b> 주간 마무리가 안되었습니다.<br>주간 마무리를 부탁드립니다.</p>" +
      "<p style='background:#FFFDF8;border-left:4px solid #FF8A3D;padding:12px 16px;border-radius:8px;font-size:14px'>" +
      "담당 반: <b>" + info.classes.join(", ") + "</b><br>" +
      "선생님 포털 → <b>주간마무리</b> 탭에서 작성해주세요.<br>특이사항이 없으면 <b>[특이사항 없음]</b> 버튼 클릭 OK!</p>" +
      "<p style='margin-top:16px'><a href='" + PORTAL_URL + "' style='background:#FF8A3D;color:#fff;text-decoration:none;font-weight:800;padding:11px 22px;border-radius:99px;display:inline-block'>선생님 포털 열기</a></p>" +
      "<p style='color:#8A7A66;font-size:12.5px;margin-top:18px'>아미쿠스 한국학교 · 이 메일은 자동 발송되었습니다.</p></div>";
    MailApp.sendEmail({
      to: email,
      subject: "[아미쿠스 한국학교] " + md + " 주간 마무리 부탁드립니다 📝",
      body: nm + " 선생님 주간 마무리가 안되었습니다. 주간 마무리를 부탁드립니다.\n" + PORTAL_URL,
      htmlBody: html,
    });
    sent++;
  });
  return sent;
}
function fridaySummary() {
  const s = buildSummary_();
  MailApp.sendEmail({ to: ADMIN_EMAIL, subject: s.subject, body: s.body, htmlBody: s.html });
  try { sendTeacherReminders_(s.today, s.weekly); } catch (e) { Logger.log("reminder error: " + e); }
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
            properties: { title: { title: [{ text: { content: "📋 " + s.today + " 요약" } }] } },
            children: s.body.split("\n").filter(Boolean).slice(0, 90).map(function(line) {
              return { object: "block", type: "paragraph",
                paragraph: { rich_text: [{ text: { content: line } }] } };
            }),
          }),
        });
      }
    }
  } catch (e) { Logger.log("notion summary error: " + e); }
}
function installFridayTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "fridaySummary") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("fridaySummary").timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(22).create();
}
/* 테스트용: 관리자 요약 + 선생님 리마인더를 즉시 발송해보기 */
function testFridayEmails() { fridaySummary(); }

/* 테스트 */
function testPortal() { Logger.log(JSON.stringify({
  students: readStudentsFull_("가람반").length,
  teachers: readTeachers_().length,
})); }
