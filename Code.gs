/** =========================
 *  Instruktorių skambučių / neatitikimų valdymas
 *  + istorija ContactHistory
 *  + "BLOGAS" scoring
 *  + drivers sync (optional)
 *  + AUTO-ASSIGN FREE DUE byloms pagal mažiausiai padarytų kontaktų
 *  + EMAIL priminimai (CLAIMED DUE) su BCC kopijomis + EmailLog
 *  + PRIVALOMAS IN/OUT + InOutLog ataskaita
 *  + getBadDrivers() – BLOGAS sąrašas UI
 *  ========================= */

const SPREADSHEET_ID = "1DAARlk_LoazXO9UgmjJJkjSKFFUR76bSyg2YT-9xTVc";

// Sheets
const SHEET_CASES      = "Cases";
const SHEET_HISTORY    = "ContactHistory";
const SHEET_DRIVERS    = "Drivers";
const SHEET_EMAIL_LOG  = "EmailLog";
const SHEET_INOUT_LOG  = "InOutLog";

// Source: DriverInspection – vairuotojų sinchronizavimas
const SOURCE_SPREADSHEET_ID = "1vOlJPkR0U0cRJdnMNtXz_C136kIUUfidfyZ6k28V1r8";
const SOURCE_SHEET_NAME = "DriverInspection";
const SOURCE_DRIVER_COL = 2;     // B: "Vardas Pavardė"
const SOURCE_HEADER_ROWS = 1;

// Instruktoriai
const INSTRUCTORS = [
  "Kęstutis Mikelevičius",
  "Jevgenij Burgan",
  ];

/** Instruktorių el. paštai */
const INSTRUCTOR_EMAILS = {
  "Kęstutis Mikelevičius": "kestutis.m.agriciaus@gmail.com",
  "Jevgenij Burgan":        "jevgenijb17@gmail.com",
  };

// Fallback email jei instruktorius nerastas
const FALLBACK_REMINDER_EMAIL = "saulius.transportas@gmail.com";

// BCC kopijos visiems siunčiamiems priminimams
const AUDIT_BCC_EMAILS = [
  "saulius.transportas@gmail.com",
  "mc@glogistics.lt"
].filter(Boolean).join(",");

// Taisyklės
const AUTO_RELEASE_DAYS         = 7;
const RULE_NO_ANSWER_BUSINESS_DAYS = 1;  // +1 d.d.
const RULE_RECHECK_BUSINESS_DAYS   = 5;  // +5 d.d.
const RULE_OK_DAYS                 = 60; // +60 d.

// BLOGAS scoring
const BAD_WINDOW_DAYS     = 60;
const BAD_SCORE_THRESHOLD = 8;
const BAD_FAILS_THRESHOLD = 3;

// AUTO-ASSIGN
const AUTO_ASSIGN_MAX_PER_RUN = 50;
const ASSIGN_WORK_WINDOW_DAYS = 30;

// Darbo rezultatų outcome
const WORK_OUTCOMES = new Set(["NEATSILIEPIA","NEATITINKA","ATITINKA","UZDARYTI"]);

/** ===== Web app ===== */
function doGet(e) {
  ensureSheets_();
  const page = (e && e.parameter && e.parameter.page) || "index";
  if (page === "analysis") {
    return HtmlService.createHtmlOutputFromFile("Analysis")
      .setTitle("Eco Analizė · DanDanPro Logistics")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === "report") {
    return HtmlService.createHtmlOutputFromFile("Report")
      .setTitle("Instruktorių ataskaita · DanDanPro Logistics")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === "drivers") {
    return HtmlService.createHtmlOutputFromFile("Drivers")
      .setTitle("Vairuotojų analizė · DanDanPro Logistics")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === "naujokai") {
    return HtmlService.createHtmlOutputFromFile("Naujokai")
      .setTitle("Naujokų analizė · DanDanPro Logistics")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Drivers CheckList · DanDanPro Logistics")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ===== Meniu funkcijos ===== */
function ensureSheets() { ensureSheets_(); }

/** ===== Grąžina Analysis puslapio URL ===== */
function getAnalysisUrl() {
  const base = ScriptApp.getService().getUrl();
  return base ? base + "?page=analysis" : "";
}

function runHourly() { sendDueInstructorReminders(); /* autoAssignDueCases() - išjungta */ }

/** ===== Install triggers (paleisti vieną kartą) ===== */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if ([
      "autoReleaseStaleClaims",
      "syncDriversFromInspection",
      "sendDueInstructorReminders",
      "autoAssignDueCases",
      "runHourly"
    ].includes(fn)) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("autoReleaseStaleClaims")
    .timeBased().atHour(6).everyDays(1).create();

  ScriptApp.newTrigger("syncDriversFromInspection")
    .timeBased().atHour(6).everyDays(1).create();

  ScriptApp.newTrigger("runHourly")
    .timeBased().everyHours(1).create();
}

/** ===== TEST: el. pašto siuntimas ===== */
function testEmailNow() {
  ensureSheets_();
  const to = FALLBACK_REMINDER_EMAIL;
  const subject = "TESTAS: DriversCheckList email siuntimas";
  const body =
    "Jei gavai – siuntimas veikia.\n\n" +
    "Laikas: " + new Date() + "\n" +
    "BCC: " + (AUDIT_BCC_EMAILS || "(tuščia)") + "\n";

  try {
    GmailApp.sendEmail(to, subject, body, { bcc: AUDIT_BCC_EMAILS });
    logEmail_("TEST", to, AUDIT_BCC_EMAILS, subject, 0, "", "");
  } catch (e) {
    logEmail_("TEST", to, AUDIT_BCC_EMAILS, subject, 0, "", String(e && e.message ? e.message : e));
    throw e;
  }
}

/** ===== Init data for UI ===== */
function getInitData() {
  ensureSheets_();
  const localDrivers = getLocalDrivers_();
  const inspectionDrivers = getDriversFromInspection_();
  const merged = uniqueNames_([].concat(localDrivers, inspectionDrivers));
  return { instructors: INSTRUCTORS, drivers: merged };
}

/** ===== Queue: rodo CALLBACK/RECHECK visada, OK tik jei suėjęs terminas ===== */
function getQueue(queryText) {
  ensureSheets_();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_CASES);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const hm = headerMap_(header);
  const c = (name, fallback) => (hm[name] ? hm[name] : fallback);

  const cCaseId      = c("CaseID", 1);
  const cDriver      = c("Vairuotojas", 2);
  const cStatus      = c("Statusas", 3);
  const cRisk        = c("Rizika", 4);
  const cCats        = c("Kategorijos", 5);
  const cSummary     = c("Paskutinės pastabos", 6);
  const cStage       = c("Stage", 7);
  const cNext        = c("Sekantis kontaktas", 8);
  const cClaimStatus = c("ClaimStatus", 9);
  const cClaimedBy   = c("Paimta (kas)", 10);
  const cClaimedAt   = c("Paimta (kada)", 11);
  const cUpdatedAt   = c("Atnaujinta (kada)", 12);
  const cBadFlag     = c("BadFlag", 13);
  const cBadScore    = c("BadScore", 14);
  const cBadReasons  = c("BadReasons", 15);

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const now = new Date();
  const q = String(queryText || "").trim().toLowerCase();
  const out = [];

  data.forEach(r => {
    const caseId = String(r[cCaseId - 1] || "").trim();
    const driverName = String(r[cDriver - 1] || "").trim();
    if (!caseId || !driverName) return;

    const status = String(r[cStatus - 1] || "OPEN").trim();
    if (status === "CLOSED") return;

    const risk = String(r[cRisk - 1] || "VIDUTINĖ").trim();
    const categories = String(r[cCats - 1] || "").trim();
    const summary = String(r[cSummary - 1] || "").trim();
    const stage = String(r[cStage - 1] || "OK").trim();
    const nextContactAt = r[cNext - 1];
    const claimStatus = String(r[cClaimStatus - 1] || "FREE").trim();
    const claimedBy = String(r[cClaimedBy - 1] || "").trim();
    const claimedAt = r[cClaimedAt - 1];
    const updatedAt = r[cUpdatedAt - 1];
    const badFlag = !!r[cBadFlag - 1];
    const badScore = Number(r[cBadScore - 1] || 0);
    const badReasons = String(r[cBadReasons - 1] || "").trim();

    const stageU = stage.toUpperCase();
    let show = false;
    if (stageU === "CALLBACK" || stageU === "RECHECK") show = true;
    else if (nextContactAt instanceof Date && nextContactAt <= now) show = true;
    if (!show) return;

    if (q) {
      const hay = [driverName,status,risk,categories,summary,stage,claimStatus,claimedBy,
        badFlag?"BLOGAS":"",String(badScore),badReasons,
        nextContactAt instanceof Date ? fmt_(nextContactAt) : "",
        claimedAt instanceof Date ? fmt_(claimedAt) : "",
        updatedAt instanceof Date ? fmt_(updatedAt) : ""
      ].join(" ").toLowerCase();
      if (!hay.includes(q)) return;
    }

    out.push({
      caseId, driverName, status, risk, categories, summary, stage,
      nextContactAt: (nextContactAt instanceof Date) ? fmt_(nextContactAt) : "",
      claimStatus, claimedBy,
      claimedAt: (claimedAt instanceof Date) ? fmt_(claimedAt) : "",
      updatedAt: (updatedAt instanceof Date) ? fmt_(updatedAt) : "",
      badFlag, badScore, badReasons
    });
  });

  out.sort((a, b) => {
    const aa = a.nextContactAt || "9999-99-99 99:99";
    const bb = b.nextContactAt || "9999-99-99 99:99";
    return aa.localeCompare(bb);
  });

  return out;
}

/** ===== Case details + history ===== */
function getCaseDetails(caseId) {
  ensureSheets_();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shC = ss.getSheetByName(SHEET_CASES);
  const shH = ss.getSheetByName(SHEET_HISTORY);

  const row = findCaseRowById_(shC, caseId);
  if (!row) throw new Error("Byla nerasta.");

  const lastCol = shC.getLastColumn();
  const header = shC.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const hm = headerMap_(header);
  const c = (name, fallback) => (hm[name] ? hm[name] : fallback);

  const r = shC.getRange(row, 1, 1, lastCol).getValues()[0];

  const details = {
    caseId:      String(r[c("CaseID",1)-1] || ""),
    driverName:  String(r[c("Vairuotojas",2)-1] || ""),
    status:      String(r[c("Statusas",3)-1] || "OPEN"),
    risk:        String(r[c("Rizika",4)-1] || "VIDUTINĖ"),
    categories:  String(r[c("Kategorijos",5)-1] || ""),
    summary:     String(r[c("Paskutinės pastabos",6)-1] || ""),
    stage:       String(r[c("Stage",7)-1] || "OK"),
    nextContactAt: (r[c("Sekantis kontaktas",8)-1] instanceof Date) ? fmt_(r[c("Sekantis kontaktas",8)-1]) : "",
    claimStatus: String(r[c("ClaimStatus",9)-1] || "FREE"),
    claimedBy:   String(r[c("Paimta (kas)",10)-1] || ""),
    claimedAt:   (r[c("Paimta (kada)",11)-1] instanceof Date) ? fmt_(r[c("Paimta (kada)",11)-1]) : "",
    updatedAt:   (r[c("Atnaujinta (kada)",12)-1] instanceof Date) ? fmt_(r[c("Atnaujinta (kada)",12)-1]) : "",
    badFlag:     !!r[c("BadFlag",13)-1],
    badScore:    Number(r[c("BadScore",14)-1] || 0),
    badReasons:  String(r[c("BadReasons",15)-1] || "")
  };

  const history = [];
  const lastH = shH.getLastRow();
  if (lastH >= 2) {
    const hData = shH.getRange(2, 1, lastH - 1, shH.getLastColumn()).getValues();
    hData.forEach(rr => {
      if (String(rr[1] || "") === details.caseId) {
        history.push({
          time:          (rr[0] instanceof Date) ? fmt_(rr[0]) : "",
          caseId:        String(rr[1] || ""),
          driverName:    String(rr[2] || ""),
          instructor:    String(rr[3] || ""),
          outcome:       String(rr[4] || ""),
          categories:    String(rr[5] || ""),
          notes:         String(rr[6] || ""),
          stage:         String(rr[7] || ""),
          nextContactAt: (rr[8] instanceof Date) ? fmt_(rr[8]) : (rr[8] ? String(rr[8]) : ""),
          source:        String(rr[9] || "")
        });
      }
    });
  }

  history.sort((a, b) => String(b.time).localeCompare(String(a.time)));
  return { details, history };
}

/** ===== Claim ===== */
function claimCase(caseId, instructorName) {
  ensureSheets_();
  const instr = String(instructorName || "").trim();
  if (!INSTRUCTORS.includes(instr)) throw new Error("Neteisingas instruktorius.");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_CASES);
  const row = findCaseRowById_(sh, caseId);
  if (!row) throw new Error("Byla nerasta.");

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
    const hm = headerMap_(header);
    const c = (name, fallback) => (hm[name] ? hm[name] : fallback);

    const r = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    const claimStatus = String(r[c("ClaimStatus",9)-1] || "FREE").trim();
    const claimedBy   = String(r[c("Paimta (kas)",10)-1] || "").trim();

    if (claimStatus === "CLAIMED") {
      return { ok: false, message: `Jau paimta: ${claimedBy || "—"}` };
    }

    const now = new Date();
    sh.getRange(row, c("ClaimStatus",9)).setValue("CLAIMED");
    sh.getRange(row, c("Paimta (kas)",10)).setValue(instr);
    sh.getRange(row, c("Paimta (kada)",11)).setValue(now);
    sh.getRange(row, c("Atnaujinta (kada)",12)).setValue(now);

    appendHistory_(ss, {
      caseId, driverName: String(r[c("Vairuotojas",2)-1] || ""),
      instructor: instr, outcome: "PAIMTA", categories: "",
      notes: "Byla paimta instruktoriaus.",
      stage: String(r[c("Stage",7)-1] || ""),
      nextContactAt: r[c("Sekantis kontaktas",8)-1] instanceof Date ? r[c("Sekantis kontaktas",8)-1] : "",
      source: "Sistema"
    });

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** ===== Release ===== */
function releaseCase(caseId, instructorName) {
  ensureSheets_();
  const instr = String(instructorName || "").trim();
  if (!INSTRUCTORS.includes(instr)) throw new Error("Neteisingas instruktorius.");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_CASES);
  const row = findCaseRowById_(sh, caseId);
  if (!row) throw new Error("Byla nerasta.");

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
    const hm = headerMap_(header);
    const c = (name, fallback) => (hm[name] ? hm[name] : fallback);

    const r = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    const claimStatus = String(r[c("ClaimStatus",9)-1] || "FREE").trim();
    const claimedBy   = String(r[c("Paimta (kas)",10)-1] || "").trim();

    if (claimStatus !== "CLAIMED") return { ok: true };

    if (normalizeName_(claimedBy) !== normalizeName_(instr)) {
      return { ok: false, message: `Negali atlaisvinti. Paimta: ${claimedBy}` };
    }

    const now = new Date();
    sh.getRange(row, c("ClaimStatus",9)).setValue("FREE");
    sh.getRange(row, c("Paimta (kas)",10)).setValue("");
    sh.getRange(row, c("Paimta (kada)",11)).setValue("");
    sh.getRange(row, c("Atnaujinta (kada)",12)).setValue(now);

    appendHistory_(ss, {
      caseId, driverName: String(r[c("Vairuotojas",2)-1] || ""),
      instructor: instr, outcome: "ATLAISVINTA", categories: "",
      notes: "Byla atlaisvinta.",
      stage: String(r[c("Stage",7)-1] || ""),
      nextContactAt: r[c("Sekantis kontaktas",8)-1] instanceof Date ? r[c("Sekantis kontaktas",8)-1] : "",
      source: "Sistema"
    });

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** ===== Save entry ===== */
function saveCaseEntry(payload) {
  ensureSheets_();
  if (!payload) throw new Error("Tuščias payload.");

  const instr = String(payload.instructorName || "").trim();
  if (!INSTRUCTORS.includes(instr)) throw new Error("Neteisingas instruktorius.");

  const driverName = String(payload.driverName || "").trim();
  if (!driverName) throw new Error("Vairuotojas privalomas.");

  const ioDirection = String(payload.ioDirection || "").trim().toUpperCase();
  if (!["IN", "OUT"].includes(ioDirection)) throw new Error("Pasirinkimas IN/OUT privalomas.");

  const risk = String(payload.risk || "VIDUTINĖ").trim();
  const categoriesArr = Array.isArray(payload.categories) ? payload.categories : [];
  const categories = categoriesArr.map(String).map(s => s.trim()).filter(Boolean).join("; ");

  const outcome = String(payload.outcome || "").trim();
  if (!outcome) throw new Error("Rezultatas privalomas.");

  const notes = String(payload.notes || "").trim();
  const caseIdFromUi = String(payload.caseId || "").trim();

  upsertLocalDriver_(driverName);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shC = ss.getSheetByName(SHEET_CASES);

  const now = new Date();
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);

  try {
    let row = 0;
    let caseId = "";

    if (caseIdFromUi) row = findCaseRowById_(shC, caseIdFromUi);
    if (!row) row = findCaseRowByDriver_(shC, driverName);

    const lastCol = shC.getLastColumn();
    const header = shC.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
    const hm = headerMap_(header);
    const c = (name, fallback) => (hm[name] ? hm[name] : fallback);

    if (!row) {
      caseId = genCaseId_();
      shC.appendRow([
        caseId, driverName, "OPEN", risk, categories, notes,
        "OK", "", "FREE", "", "", now,
        false, 0, "",
        "", "", "",
        "", ""
      ]);
      row = shC.getLastRow();
    } else {
      caseId = String(shC.getRange(row, c("CaseID",1)).getValue());
    }

    let stage = "OK";
    let status = "IN_PROGRESS";
    let nextContactAt = "";

    if (outcome === "NEATSILIEPIA") {
      stage = "CALLBACK"; nextContactAt = addBusinessDays_(now, RULE_NO_ANSWER_BUSINESS_DAYS); status = "IN_PROGRESS";
    } else if (outcome === "NEATITINKA") {
      stage = "RECHECK"; nextContactAt = addBusinessDays_(now, RULE_RECHECK_BUSINESS_DAYS); status = "IN_PROGRESS";
    } else if (outcome === "ATITINKA") {
      stage = "OK"; nextContactAt = addDays_(now, RULE_OK_DAYS); status = "OPEN";
    } else if (outcome === "UZDARYTI") {
      stage = "OK"; nextContactAt = ""; status = "CLOSED";
    } else {
      throw new Error("Neteisingas rezultatas.");
    }

    shC.getRange(row, c("Statusas",3)).setValue(status);
    shC.getRange(row, c("Rizika",4)).setValue(risk);
    shC.getRange(row, c("Kategorijos",5)).setValue(categories);
    shC.getRange(row, c("Paskutinės pastabos",6)).setValue(notes);
    shC.getRange(row, c("Stage",7)).setValue(stage);
    shC.getRange(row, c("Sekantis kontaktas",8)).setValue(nextContactAt);
    shC.getRange(row, c("Atnaujinta (kada)",12)).setValue(now);

    appendHistory_(ss, { caseId, driverName, instructor: instr, outcome, categories, notes, stage, nextContactAt, source: "Rankinis" });
    appendInOutLog_(ss, { caseId, driverName, instructor: instr, ioDirection, outcome, categories, notes, source: "Rankinis" });
    recalcBadForCase_(ss, shC, row, caseId);

    return {
      ok: true, caseId, stage,
      nextContactAt: (nextContactAt instanceof Date) ? fmt_(nextContactAt) : ""
    };
  } finally {
    lock.releaseLock();
  }
}

/** ===== BLOGAS vairuotojų sąrašas (NAUJAS) ===== */
function getBadDrivers() {
  ensureSheets_();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_CASES);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const hm = headerMap_(header);
  const c = (name, fallback) => (hm[name] ? hm[name] : fallback);

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const out = [];

  data.forEach(r => {
    const badFlag = !!r[c("BadFlag", 13) - 1];
    if (!badFlag) return;

    const caseId     = String(r[c("CaseID", 1) - 1] || "").trim();
    const driverName = String(r[c("Vairuotojas", 2) - 1] || "").trim();
    if (!caseId || !driverName) return;

    out.push({
      caseId,
      driverName,
      status:     String(r[c("Statusas", 3) - 1] || "OPEN"),
      risk:       String(r[c("Rizika", 4) - 1] || ""),
      categories: String(r[c("Kategorijos", 5) - 1] || ""),
      summary:    String(r[c("Paskutinės pastabos", 6) - 1] || ""),
      badScore:   Number(r[c("BadScore", 14) - 1] || 0),
      badReasons: String(r[c("BadReasons", 15) - 1] || ""),
      updatedAt:  (r[c("Atnaujinta (kada)", 12) - 1] instanceof Date)
                    ? fmt_(r[c("Atnaujinta (kada)", 12) - 1]) : ""
    });
  });

  // Pirma aktyvūs (pagal badScore mažėjančiai), tada uždaryti
  out.sort((a, b) => {
    if (a.status === "CLOSED" && b.status !== "CLOSED") return 1;
    if (a.status !== "CLOSED" && b.status === "CLOSED") return -1;
    return b.badScore - a.badScore;
  });

  return out;
}

/** ===== Recalculate BLOGAS ===== */
function recalcBadForCase_(ss, shC, caseRow, caseId) {
  const shH = ss.getSheetByName(SHEET_HISTORY);
  const lastH = shH.getLastRow();
  if (lastH < 2) return;

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - BAD_WINDOW_DAYS);

  const h = shH.getRange(2, 1, lastH - 1, shH.getLastColumn()).getValues();
  let score = 0, cntBad = 0, cntNo = 0;
  const catCounts = {};

  h.forEach(r => {
    const t = r[0]; const cid = String(r[1] || "");
    const outcome = String(r[4] || ""); const cats = String(r[5] || "");
    if (cid !== caseId) return;
    if (!(t instanceof Date)) return;
    if (t < cutoff) return;
    if (outcome === "NEATITINKA") { score += 3; cntBad++; }
    if (outcome === "NEATSILIEPIA") { score += 2; cntNo++; }
    if (cats) { cats.split(";").map(x=>x.trim()).filter(Boolean).forEach(c => { score += 1; catCounts[c] = (catCounts[c]||0)+1; }); }
  });

  const badFlag = (score >= BAD_SCORE_THRESHOLD) || (cntBad >= BAD_FAILS_THRESHOLD);
  const topCats = Object.keys(catCounts).sort((a,b)=>(catCounts[b]-catCounts[a])).slice(0,6).map(k=>`${k}(${catCounts[k]})`).join(", ");
  const reasons = [`Langas: ${BAD_WINDOW_DAYS} d.`,`Taškai: ${score}`,`Neatitinka: ${cntBad}`,`Neatsiliepia: ${cntNo}`,topCats?`Kategorijos: ${topCats}`:""].filter(Boolean).join("; ");

  const lastCol = shC.getLastColumn();
  const header = shC.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const hm2 = headerMap_(header);
  if (!hm2["BadFlag"] || !hm2["BadScore"] || !hm2["BadReasons"]) return;

  shC.getRange(caseRow, hm2["BadFlag"]).setValue(badFlag);
  shC.getRange(caseRow, hm2["BadScore"]).setValue(score);
  shC.getRange(caseRow, hm2["BadReasons"]).setValue(reasons);
}

/** ===== Auto-release stale CLAIMs ===== */
function autoReleaseStaleClaims() {
  ensureSheets_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_CASES);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const hm = headerMap_(header);
  const c = (name, fallback) => (hm[name] ? hm[name] : fallback);

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const now = new Date();

  data.forEach((r, idx) => {
    const row = idx + 2;
    const caseId     = String(r[c("CaseID",1)-1] || "");
    const driverName = String(r[c("Vairuotojas",2)-1] || "");
    const claimStatus = String(r[c("ClaimStatus",9)-1] || "FREE");
    const claimedBy  = String(r[c("Paimta (kas)",10)-1] || "");
    const updatedAt  = r[c("Atnaujinta (kada)",12)-1];

    if (claimStatus !== "CLAIMED") return;
    if (!(updatedAt instanceof Date)) return;

    const ageDays = (now.getTime() - updatedAt.getTime()) / (1000 * 60 * 60 * 24);
    if (ageDays >= AUTO_RELEASE_DAYS) {
      sh.getRange(row, c("ClaimStatus",9)).setValue("FREE");
      sh.getRange(row, c("Paimta (kas)",10)).setValue("");
      sh.getRange(row, c("Paimta (kada)",11)).setValue("");
      sh.getRange(row, c("Atnaujinta (kada)",12)).setValue(now);

      appendHistory_(ss, {
        caseId, driverName, instructor: "SYSTEM", outcome: "AUTO_RELEASE", categories: "",
        notes: `Automatiškai atlaisvinta po ${AUTO_RELEASE_DAYS} d. be veiksmų. Buvo paėmęs: ${claimedBy}`,
        stage: String(r[c("Stage",7)-1] || ""),
        nextContactAt: (r[c("Sekantis kontaktas",8)-1] instanceof Date) ? r[c("Sekantis kontaktas",8)-1] : "",
        source: "Sistema"
      });
      recalcBadForCase_(ss, sh, row, caseId);
    }
  });
}

/** ===== Sync drivers from DriverInspection ===== */
function syncDriversFromInspection() {
  ensureSheets_();
  const names = getDriversFromInspection_();
  if (!names.length) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_DRIVERS);
  const existing = getLocalDrivers_();
  const merged = uniqueNames_([].concat(existing, names));

  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).clearContent();
  if (merged.length) sh.getRange(2, 1, merged.length, 1).setValues(merged.map(x => [x]));
}

/** ===== AUTO-ASSIGN ===== */
function autoAssignDueCases() {
  // AUTO-ASSIGN IŠJUNGTAS – bylos priskiriamos tik rankiniu būdu
  return;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_CASES);

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
    const hm = headerMap_(header);

    ["AutoAssignedAt","AutoAssignedTo"].forEach(n => {
      if (!hm[n]) sh.getRange(1, sh.getLastColumn() + 1).setValue(n);
    });

    const lastCol2 = sh.getLastColumn();
    const header2 = sh.getRange(1, 1, 1, lastCol2).getValues()[0].map(h => String(h || "").trim());
    const hm2 = headerMap_(header2);
    const c = (name, fallback) => (hm2[name] ? hm2[name] : fallback);

    const cCaseId      = c("CaseID", 1);
    const cDriver      = c("Vairuotojas", 2);
    const cStatus      = c("Statusas", 3);
    const cStage       = c("Stage", 7);
    const cNext        = c("Sekantis kontaktas", 8);
    const cClaimStatus = c("ClaimStatus", 9);
    const cClaimedBy   = c("Paimta (kas)", 10);
    const cClaimedAt   = c("Paimta (kada)", 11);
    const cUpdatedAt   = c("Atnaujinta (kada)", 12);
    const cAutoAt      = c("AutoAssignedAt", 19);
    const cAutoTo      = c("AutoAssignedTo", 20);

    const now = new Date();
    const data = sh.getRange(2, 1, lastRow - 1, lastCol2).getValues();
    const counts = getInstructorWorkCounts_(ss, ASSIGN_WORK_WINDOW_DAYS);
    let assigned = 0;

    for (let i = 0; i < data.length; i++) {
      if (assigned >= AUTO_ASSIGN_MAX_PER_RUN) break;
      const r = data[i]; const rowNr = i + 2;
      const status = String(r[cStatus - 1] || "").trim();
      if (status === "CLOSED") continue;
      const next = r[cNext - 1];
      if (!(next instanceof Date)) continue;
      if (next > now) continue;
      const claimStatus = String(r[cClaimStatus - 1] || "FREE").trim();
      if (claimStatus !== "FREE") continue;

      const chosen = pickLeastBusyInstructor_(counts);
      if (!chosen) continue;

      const caseId = String(r[cCaseId - 1] || "").trim();
      const driver = String(r[cDriver - 1] || "").trim();
      const stage  = String(r[cStage - 1] || "").trim();

      sh.getRange(rowNr, cClaimStatus).setValue("CLAIMED");
      sh.getRange(rowNr, cClaimedBy).setValue(chosen);
      sh.getRange(rowNr, cClaimedAt).setValue(now);
      sh.getRange(rowNr, cUpdatedAt).setValue(now);
      sh.getRange(rowNr, cAutoAt).setValue(now);
      sh.getRange(rowNr, cAutoTo).setValue(chosen);

      appendHistory_(ss, {
        caseId, driverName: driver, instructor: "SYSTEM", outcome: "AUTO_ASSIGN", categories: "",
        notes: `Automatiškai priskirta instruktoriui: ${chosen}. Priežastis: suėjęs terminas ir byla nepaimta.`,
        stage, nextContactAt: next, source: "Sistema"
      });

      counts[chosen] = (counts[chosen] || 0) + 1;
      assigned++;
    }
  } finally {
    lock.releaseLock();
  }
}

/** ===== EMAIL priminimai ===== */
function sendDueInstructorReminders() {
  ensureSheets_();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_CASES);

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const lastCol = sh.getLastColumn();
    let header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
    let hm = headerMap_(header);

    ["ReminderSentFor","ReminderSentAt","ReminderSentTo"].forEach(n => {
      if (!hm[n]) sh.getRange(1, sh.getLastColumn() + 1).setValue(n);
    });

    const lastCol2 = sh.getLastColumn();
    header = sh.getRange(1, 1, 1, lastCol2).getValues()[0].map(h => String(h || "").trim());
    hm = headerMap_(header);
    const c = (name, fallback) => (hm[name] ? hm[name] : fallback);

    const cCaseId      = c("CaseID", 1);
    const cDriver      = c("Vairuotojas", 2);
    const cStatus      = c("Statusas", 3);
    const cStage       = c("Stage", 7);
    const cNext        = c("Sekantis kontaktas", 8);
    const cClaimStatus = c("ClaimStatus", 9);
    const cClaimedBy   = c("Paimta (kas)", 10);
    const cSentFor     = c("ReminderSentFor", 16);
    const cSentAt      = c("ReminderSentAt", 17);
    const cSentTo      = c("ReminderSentTo", 18);

    const data = sh.getRange(2, 1, lastRow - 1, lastCol2).getValues();
    const now = new Date();
    const bucket = {};

    data.forEach((r, idx) => {
      const rowNr = idx + 2;
      const status = String(r[cStatus - 1] || "").trim();
      if (status === "CLOSED") return;
      const next = r[cNext - 1];
      if (!(next instanceof Date)) return;
      if (next > now) return;
      const claimStatus = String(r[cClaimStatus - 1] || "FREE").trim();
      if (claimStatus !== "CLAIMED") return;
      const claimedBy = String(r[cClaimedBy - 1] || "").trim();
      if (!claimedBy) return;

      const sentKey = dateKey_(r[cSentFor - 1]);
      const nextKey = dateKey_(next);
      if (sentKey && sentKey === nextKey) return;

      const email = getInstructorEmail_(claimedBy) || FALLBACK_REMINDER_EMAIL;
      if (!email) return;

      const caseId = String(r[cCaseId - 1] || "").trim();
      const driver = String(r[cDriver - 1] || "").trim();
      const stage  = String(r[cStage - 1] || "").trim();

      if (!bucket[claimedBy]) bucket[claimedBy] = { email, items: [] };
      bucket[claimedBy].items.push({ rowNr, caseId, driver, stage, next });
    });

    const appUrl = ScriptApp.getService().getUrl() || "";

    Object.keys(bucket).forEach(instrName => {
      const pack = bucket[instrName];
      const items = pack.items || [];
      if (!items.length) return;

      items.sort((a, b) => a.next.getTime() - b.next.getTime());

      const lines = items.map(it => `- ${it.driver} | ${it.stage || "—"} | Terminas: ${fmt_(it.next)} | ${it.caseId}`).join("\n");
      const subject = `PRIMINIMAS: suėjo terminas kontaktui (${items.length})`;
      const body = `Sveiki, ${instrName},\n\nSuėjo terminas sekančiam kontaktui (Cases lape):\n\n${lines}\n\n${appUrl ? ("Atidaryti sistemą: " + appUrl) : ""}\n\n— Sistema`;
      const caseIds = items.map(x => x.caseId).join(",");

      try {
        GmailApp.sendEmail(pack.email, subject, body, { bcc: AUDIT_BCC_EMAILS });
        const sentAt = new Date();
        items.forEach(it => {
          sh.getRange(it.rowNr, cSentFor).setValue(it.next);
          sh.getRange(it.rowNr, cSentAt).setValue(sentAt);
          sh.getRange(it.rowNr, cSentTo).setValue(pack.email);
        });
        logEmail_("REMINDER", pack.email, AUDIT_BCC_EMAILS, subject, items.length, caseIds, "");
      } catch (e) {
        logEmail_("REMINDER", pack.email, AUDIT_BCC_EMAILS, subject, items.length, caseIds, String(e && e.message ? e.message : e));
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function getInstructorEmail_(name) {
  return INSTRUCTOR_EMAILS[String(name || "").trim()] || "";
}

/** ===== EmailLog ===== */
function logEmail_(type, to, bcc, subject, itemCount, caseIds, errorMsg) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_EMAIL_LOG);
    if (!sh) return;
    sh.appendRow([new Date(), String(type||""), String(to||""), String(bcc||""), String(subject||""), Number(itemCount||0), String(caseIds||""), String(errorMsg||"")]);
  } catch (e) {}
}

/** ===== InOutLog ===== */
function appendInOutLog_(ss, o) {
  try {
    const sh = ss.getSheetByName(SHEET_INOUT_LOG);
    if (!sh) return;
    sh.appendRow([new Date(), String(o.caseId||""), String(o.driverName||""), String(o.instructor||""), String(o.ioDirection||""), String(o.outcome||""), String(o.categories||""), String(o.notes||""), String(o.source||"")]);
  } catch (e) {}
}

/** ===== Sheet bootstrap ===== */
function ensureSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let shC = ss.getSheetByName(SHEET_CASES);
  if (!shC) shC = ss.insertSheet(SHEET_CASES);
  if (shC.getLastRow() === 0) {
    shC.appendRow(["CaseID","Vairuotojas","Statusas","Rizika","Kategorijos","Paskutinės pastabos","Stage","Sekantis kontaktas","ClaimStatus","Paimta (kas)","Paimta (kada)","Atnaujinta (kada)","BadFlag","BadScore","BadReasons","ReminderSentFor","ReminderSentAt","ReminderSentTo","AutoAssignedAt","AutoAssignedTo"]);
    shC.setFrozenRows(1);
  } else {
    const header = shC.getRange(1, 1, 1, shC.getLastColumn()).getValues()[0].map(h => String(h||"").trim());
    const hm = headerMap_(header);
    ["BadFlag","BadScore","BadReasons","ReminderSentFor","ReminderSentAt","ReminderSentTo","AutoAssignedAt","AutoAssignedTo"]
      .forEach(n => { if (!hm[n]) shC.getRange(1, shC.getLastColumn()+1).setValue(n); });
    shC.setFrozenRows(1);
  }

  let shH = ss.getSheetByName(SHEET_HISTORY);
  if (!shH) shH = ss.insertSheet(SHEET_HISTORY);
  if (shH.getLastRow() === 0) {
    shH.appendRow(["Laikas","CaseID","Vairuotojas","Instruktorius","Rezultatas","Kategorijos","Pastabos","Stage","Sekantis kontaktas","Šaltinis"]);
    shH.setFrozenRows(1);
  }

  let shD = ss.getSheetByName(SHEET_DRIVERS);
  if (!shD) { shD = ss.insertSheet(SHEET_DRIVERS); shD.appendRow(["Vairuotojas"]); shD.setFrozenRows(1); }

  let shE = ss.getSheetByName(SHEET_EMAIL_LOG);
  if (!shE) shE = ss.insertSheet(SHEET_EMAIL_LOG);
  if (shE.getLastRow() === 0) { shE.appendRow(["Laikas","Tipas","To","Bcc","Subject","Items","CaseIds","Error"]); shE.setFrozenRows(1); }

  let shIO = ss.getSheetByName(SHEET_INOUT_LOG);
  if (!shIO) shIO = ss.insertSheet(SHEET_INOUT_LOG);
  if (shIO.getLastRow() === 0) { shIO.appendRow(["Laikas","CaseID","Vairuotojas","Instruktorius","IN_OUT","Rezultatas","Kategorijos","Pastabos","Šaltinis"]); shIO.setFrozenRows(1); }

  try {
    shC.getRange("H:H").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    shC.getRange("K:K").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    shC.getRange("L:L").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    shH.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    shH.getRange("I:I").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    ss.getSheetByName(SHEET_EMAIL_LOG).getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    ss.getSheetByName(SHEET_INOUT_LOG).getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
  } catch (e) {}
}

/** ===== Helpers ===== */
function headerMap_(headerArr) {
  const m = {};
  (headerArr||[]).forEach((h, idx) => { const key=String(h||"").trim(); if(key&&!m[key]) m[key]=idx+1; });
  return m;
}

function getDriversFromInspection_() {
  try {
    const src = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    const sh = src.getSheetByName(SOURCE_SHEET_NAME);
    if (!sh) return [];
    const last = sh.getLastRow();
    if (last <= SOURCE_HEADER_ROWS) return [];
    return uniqueNames_(sh.getRange(SOURCE_HEADER_ROWS+1, SOURCE_DRIVER_COL, last-SOURCE_HEADER_ROWS, 1).getValues().flat().map(v=>String(v||"").trim()).filter(Boolean));
  } catch (e) { return []; }
}

function getLocalDrivers_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_DRIVERS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues().flat().map(v=>String(v||"").trim()).filter(Boolean);
}

function upsertLocalDriver_(name) {
  const n = String(name||"").trim(); if (!n) return;
  const existing = getLocalDrivers_();
  const norm = normalizeName_(n);
  if (existing.some(x => normalizeName_(x) === norm)) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss.getSheetByName(SHEET_DRIVERS).appendRow([n]);
}

function uniqueNames_(arr) {
  const seen = new Set(); const out = [];
  (arr||[]).forEach(v => { const s=String(v||"").trim().replace(/\s+/g," "); if(!s) return; const k=s.toLowerCase(); if(seen.has(k)) return; seen.add(k); out.push(s); });
  return out;
}

function genCaseId_() {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
  return `CASE_${stamp}_${Math.floor(Math.random()*900+100)}`;
}

function findCaseRowById_(sh, caseId) {
  const id = String(caseId||"").trim(); if (!id) return 0;
  const last = sh.getLastRow(); if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last-1, 1).getValues().flat();
  const idx = vals.findIndex(v => String(v) === id);
  return idx >= 0 ? idx + 2 : 0;
}

function findCaseRowByDriver_(sh, driverName) {
  const dn = normalizeName_(driverName);
  const last = sh.getLastRow(); if (last < 2) return 0;
  const vals = sh.getRange(2, 2, last-1, 1).getValues().flat();
  const idx = vals.findIndex(v => normalizeName_(v) === dn);
  return idx >= 0 ? idx + 2 : 0;
}

function normalizeName_(s) {
  return String(s||"").trim().replace(/\s+/g," ").toLowerCase();
}

function appendHistory_(ss, o) {
  ss.getSheetByName(SHEET_HISTORY).appendRow([new Date(), String(o.caseId||""), String(o.driverName||""), String(o.instructor||""), String(o.outcome||""), String(o.categories||""), String(o.notes||""), String(o.stage||""), o.nextContactAt||"", String(o.source||"")]);
}

function fmt_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
}

function addDays_(dateObj, days) {
  const d = new Date(dateObj.getTime()); d.setDate(d.getDate() + Number(days)); return d;
}

function addBusinessDays_(dateObj, businessDays) {
  let d = new Date(dateObj.getTime()); let left = Number(businessDays)||0;
  while (left > 0) { d.setDate(d.getDate()+1); const day=d.getDay(); if(day!==0&&day!==6) left--; }
  return d;
}

function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyyMMddHHmm");
  return String(v||"").trim();
}

function getInstructorWorkCounts_(ss, windowDays) {
  const counts = {}; INSTRUCTORS.forEach(n => counts[n] = 0);
  const shH = ss.getSheetByName(SHEET_HISTORY);
  const lastH = shH.getLastRow(); if (lastH < 2) return counts;
  const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - Number(windowDays||30));
  shH.getRange(2, 1, lastH-1, shH.getLastColumn()).getValues().forEach(r => {
    const t=r[0]; if(!(t instanceof Date)) return; if(t<cutoff) return;
    const instr=String(r[3]||"").trim(); const outcome=String(r[4]||"").trim();
    if(!INSTRUCTORS.includes(instr)) return; if(!WORK_OUTCOMES.has(outcome)) return;
    counts[instr]=(counts[instr]||0)+1;
  });
  return counts;
}

function pickLeastBusyInstructor_(counts) {
  let best="", bestVal=1e18;
  INSTRUCTORS.forEach(n => { const v=Number(counts[n]||0); if(v<bestVal){ bestVal=v; best=n; } });
  return best;
}

/** ===== HTML suderinamumo adapteriai ===== */
function getLists() {
  const init = getInitData();
  return { drivers: init.drivers || [] };
}

function getCases(queryText) {
  return (getQueue(queryText||"")||[]).map(it => ({
    caseId:       it.caseId,
    driverName:   it.driverName,
    instructor:   it.claimedBy || "",
    categories:   it.categories || "",
    status:       it.claimStatus || "FREE",
    claimedAt:    it.claimedAt || "",
    nextContactAt: it.nextContactAt || ""
  }));
}

function getCaseById(caseId) {
  const res = getCaseDetails(caseId); const d = res.details || {};
  return { caseId: d.caseId||"", driverName: d.driverName||"", instructor: d.claimedBy||"", categories: d.categories||"", status: d.claimStatus||"FREE", claimedAt: d.claimedAt||"", nextContactAt: d.nextContactAt||"" };
}


/** ===== Rankinis BadFlag išvalymas (vairuotojas pasitaisė) ===== */
const ADMIN_NAME = "Saulius";

function clearBadFlag(caseId, requestorName) {
  ensureSheets_();
  const who = String(requestorName || "").trim();
  if (who !== ADMIN_NAME) {
    return { ok: false, message: "Tik administratorius gali išvalyti BadFlag." };
  }
  const cid = String(caseId || "").trim();
  if (!cid) return { ok: false, message: "Nenurodytas CaseID." };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_CASES);
  const row = findCaseRowById_(sh, cid);
  if (!row) return { ok: false, message: "Byla nerasta." };

  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const hm = headerMap_(header);

  if (!hm["BadFlag"] || !hm["BadScore"] || !hm["BadReasons"]) {
    return { ok: false, message: "Nerasti BadFlag stulpeliai." };
  }

  const driverName = String(sh.getRange(row, hm["Vairuotojas"] || 2).getValue() || "");

  sh.getRange(row, hm["BadFlag"]).setValue(false);
  sh.getRange(row, hm["BadScore"]).setValue(0);
  sh.getRange(row, hm["BadReasons"]).setValue("Rankiniu būdu išvalyta " + fmt_(new Date()));

  // Įrašome į istoriją
  appendHistory_(ss, {
    caseId: cid,
    driverName: driverName,
    instructor: "SYSTEM",
    outcome: "BAD_CLEARED",
    categories: "",
    notes: "BadFlag rankiniu būdu išvalyta – vairuotojas pasitaisė.",
    stage: "",
    nextContactAt: "",
    source: "Rankinis"
  });

  return { ok: true };
}


/** ===== Analizė: Volvo ir Mercedes eco duomenys ===== */
function getAnalysisData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function readSheet(sheetName, colPlate, colName, colBal, colKuras) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return { error: "Lapas \"" + sheetName + "\" nerastas." };

    const lastRow = sh.getLastRow();
    if (lastRow < 1) return { good: [], bad: [], neutral: [], fuelGood: [], fuelBad: [], byLetter: [] };

    const maxCol = Math.max(colPlate, colName, colBal, colKuras);
    const data = sh.getRange(1, 1, lastRow, maxCol).getValues();

    const good = [], bad = [], neutral = [];
    const fuelGood = [], fuelBad = [];
    let countGoodBal = 0, countBadBal = 0, countNeutral = 0;
    let countFuelGood = 0, countFuelBad = 0;
    let sumBal = 0, sumFuel = 0, countBal = 0, countFuel = 0;

    // Pagal pirmą numerio raidę: { "A": { sumBal, cntBal, sumFuel, cntFuel } }
    const letterMap = {};

    data.forEach(function(row) {
      const plate = String(row[colPlate - 1] || "").trim();
      const name  = String(row[colName  - 1] || "").trim();
      const bal   = parseFloat(String(row[colBal   - 1] || "").replace(",", "."));
      const kuras = parseFloat(String(row[colKuras - 1] || "").replace(",", "."));

      if (!name) return;

      const balValid  = !isNaN(bal);
      const fuelValid = !isNaN(kuras);

      if (balValid)  { sumBal  += bal;   countBal++;  }
      if (fuelValid) { sumFuel += kuras; countFuel++; }

      let balStatus = "";
      if (balValid && bal < 9)         { balStatus = "bad";     countBadBal++;  }
      else if (balValid && bal > 9.61) { balStatus = "good";    countGoodBal++; }
      else                             { balStatus = "neutral"; countNeutral++; }

      let fuelStatus = "";
      if (fuelValid && kuras > 26.9) fuelStatus = "bad";
      else if (fuelValid && kuras < 24) fuelStatus = "good";

      const entry = { name, bal: balValid ? bal : null, kuras: fuelValid ? kuras : null, balStatus, fuelStatus };

      if (balStatus === "bad")        bad.push(entry);
      else if (balStatus === "good")  good.push(entry);
      else                            neutral.push(entry);

      if (fuelStatus === "bad")       fuelBad.push(entry);
      else if (fuelStatus === "good") fuelGood.push(entry);

      // Grupuoti pagal pirmą numerio raidę
      if (plate) {
        const letter = plate.charAt(0).toUpperCase();
        if (!letterMap[letter]) letterMap[letter] = { sumBal: 0, cntBal: 0, sumFuel: 0, cntFuel: 0, count: 0 };
        letterMap[letter].count++;
        if (balValid)  { letterMap[letter].sumBal  += bal;   letterMap[letter].cntBal++;  }
        if (fuelValid) { letterMap[letter].sumFuel += kuras; letterMap[letter].cntFuel++; }
      }
    });

    bad.sort(function(a,b){ return (a.bal||0)-(b.bal||0); });
    good.sort(function(a,b){ return (b.bal||0)-(a.bal||0); });
    fuelBad.sort(function(a,b){ return (b.kuras||0)-(a.kuras||0); });
    fuelGood.sort(function(a,b){ return (a.kuras||0)-(b.kuras||0); });

    // Raidžių sąrašas abėcėlės tvarka (A=seniausias)
    const byLetter = Object.keys(letterMap).sort().map(function(letter) {
      const d = letterMap[letter];
      return {
        letter: letter,
        count:   d.count,
        avgBal:  d.cntBal  > 0 ? Math.round(d.sumBal  / d.cntBal  * 100) / 100 : null,
        avgFuel: d.cntFuel > 0 ? Math.round(d.sumFuel / d.cntFuel * 100) / 100 : null
      };
    });

    return {
      good, bad, neutral, fuelGood, fuelBad,
      byLetter,
      stats: {
        total: countBal,
        goodBal: countGoodBal,
        badBal: countBadBal,
        neutralBal: countNeutral,
        avgBal:  countBal  > 0 ? Math.round(sumBal  / countBal  * 100) / 100 : null,
        avgFuel: countFuel > 0 ? Math.round(sumFuel / countFuel * 100) / 100 : null,
        fuelGoodCount: fuelGood.length,
        fuelBadCount:  fuelBad.length
      }
    };
  }

  return {
    volvo:    readSheet("Volvo",    1, 2, 7, 4),   // A=1 numeris, B=2 vardas, G=7 balas, D=4 kuras
    mercedes: readSheet("Mercedes", 1, 2, 4, 10),  // A=1 numeris, B=2 vardas, D=4 balas, J=10 kuras
    updated:  new Date().toLocaleString("lt-LT")
  };
}


/** ===== Instruktorių veiklos ataskaita pagal laikotarpį ===== */
function getInstructorReport(dateFrom, dateTo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shH = ss.getSheetByName(SHEET_HISTORY);
  const shC = ss.getSheetByName(SHEET_CASES);

  const from = dateFrom ? new Date(dateFrom) : null;
  const to   = dateTo   ? new Date(dateTo)   : null;
  if (to) to.setHours(23, 59, 59, 999);

  const TRACKED = ["NEATSILIEPIA","NEATITINKA","ATITINKA","UZDARYTI","PAIMTA","ATLAISVINTA","AUTO_ASSIGN","AUTO_RELEASE","BAD_CLEARED"];
  const INSTRUCTORS_RPT = ["Kęstutis Mikelevičius","Jevgenij Burgan","Vilma Rustamova"];

  // ── ContactHistory ──
  const histRows = shH.getLastRow() > 1
    ? shH.getRange(2, 1, shH.getLastRow()-1, shH.getLastColumn()).getValues()
    : [];

  // Dienos -> instruktorius -> outcome -> count
  const byDay   = {}; // { "2026-04-01": { "Vilma": { ATITINKA: 2, ... } } }
  const byInstr = {}; // { "Vilma": { ATITINKA: 5, total: 10, ... } }
  const allDays = new Set();

  INSTRUCTORS_RPT.forEach(n => { byInstr[n] = { total: 0 }; TRACKED.forEach(o => byInstr[n][o] = 0); });

  // Kategorijų skaičiavimas: instruktorius -> kategorija -> skaičius
  const byCat = {}; // { "Vilma": { DOCS: 5, TIME: 3, ... } }
  INSTRUCTORS_RPT.forEach(n => byCat[n] = {});

  histRows.forEach(r => {
    const t       = r[0];
    const instr   = String(r[3] || "").trim();
    const outcome = String(r[4] || "").trim();
    if (!(t instanceof Date)) return;
    if (from && t < from) return;
    if (to   && t > to)   return;
    if (!INSTRUCTORS_RPT.includes(instr)) return;
    if (!TRACKED.includes(outcome)) return;

    const dayKey = Utilities.formatDate(t, Session.getScriptTimeZone(), "yyyy-MM-dd");
    allDays.add(dayKey);

    if (!byDay[dayKey]) byDay[dayKey] = {};
    if (!byDay[dayKey][instr]) { byDay[dayKey][instr] = { total: 0 }; TRACKED.forEach(o => byDay[dayKey][instr][o] = 0); }

    byDay[dayKey][instr][outcome] = (byDay[dayKey][instr][outcome] || 0) + 1;
    byDay[dayKey][instr].total++;
    byInstr[instr][outcome]++;
    byInstr[instr].total++;

    // Kategorijos (stulpelis r[5] ContactHistory)
    const cats = String(r[5] || "").trim();
    if (cats && ["NEATSILIEPIA","NEATITINKA","ATITINKA","UZDARYTI"].includes(outcome)) {
      cats.split(/[;,]+/).map(x => x.trim()).filter(Boolean).forEach(cat => {
        byCat[instr][cat] = (byCat[instr][cat] || 0) + 1;
      });
    }
  });

  // ── Cases: kiek bylų paimta per laikotarpį ──
  const caseRows = shC.getLastRow() > 1
    ? shC.getRange(2, 1, shC.getLastRow()-1, shC.getLastColumn()).getValues()
    : [];

  const caseHeader = shC.getRange(1,1,1,shC.getLastColumn()).getValues()[0].map(h=>String(h||"").trim());
  const hm = {};
  caseHeader.forEach((h,i) => { if(h && !hm[h]) hm[h] = i; });

  const casesByInstr = {};
  INSTRUCTORS_RPT.forEach(n => casesByInstr[n] = { open: 0, closed: 0 });

  caseRows.forEach(r => {
    const claimedBy = String(r[hm["Paimta (kas)"] ?? 9] || "").trim();
    const status    = String(r[hm["Statusas"]     ?? 2] || "").trim();
    const updAt     = r[hm["Atnaujinta (kada)"] ?? 11];
    if (!INSTRUCTORS_RPT.includes(claimedBy)) return;
    if (from && updAt instanceof Date && updAt < from) return;
    if (to   && updAt instanceof Date && updAt > to)   return;
    if (status === "CLOSED") casesByInstr[claimedBy].closed++;
    else casesByInstr[claimedBy].open++;
  });

  // ── Dienų sąrašas surūšiuotas ──
  const days = Array.from(allDays).sort();

  // ── Timeline duomenys grafikui ──
  // Kiekvienai dienai – kiekvieno instruktoriaus total veiksmų
  const timeline = days.map(day => {
    const entry = { day };
    INSTRUCTORS_RPT.forEach(n => {
      entry[n] = byDay[day] && byDay[day][n] ? byDay[day][n].total : 0;
    });
    return entry;
  });

  // ── Detalūs duomenys per dieną kiekvienam instruktoriui ──
  const dailyDetail = days.map(day => {
    const row = { day };
    INSTRUCTORS_RPT.forEach(n => {
      row[n] = byDay[day] && byDay[day][n] ? byDay[day][n] : { total: 0 };
      TRACKED.forEach(o => { if (!row[n][o]) row[n][o] = 0; });
    });
    return row;
  });

  return {
    instructors: INSTRUCTORS_RPT,
    tracked:     TRACKED,
    byInstr:     byInstr,
    casesByInstr:casesByInstr,
    timeline:    timeline,
    dailyDetail: dailyDetail,
    days:        days,
    dateFrom:    dateFrom || "",
    dateTo:      dateTo   || "",
    generated:   new Date().toLocaleString("lt-LT"),
    byCat:       byCat
  };
}


/** ===== Vairuotojų demografijos analizė ===== */
function getDriversAnalysis(dateFrom, dateTo) {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const from = dateFrom ? new Date(dateFrom) : null;
  const to   = dateTo   ? new Date(dateTo)   : null;
  if (to) to.setHours(23, 59, 59, 999);

  function parseSheet(sheetName, dateCol, natCol, stazCol) {
    // Ieškome lapo – tikslus pavadinimas arba be tarpų/didžių raidžių
    let sh = ss.getSheetByName(sheetName);
    if (!sh) {
      const norm = s => s.toLowerCase().replace(/\s+/g,"");
      sh = ss.getSheets().find(s => norm(s.getName()) === norm(sheetName)) || null;
    }
    if (!sh) {
      const available = ss.getSheets().map(s => s.getName()).join(", ");
      return { error: "Lapas \"" + sheetName + "\" nerastas. Esami lapai: " + available, total: 0, natList: [], stazGroups: {"0–1 m.":0,"1–3 m.":0,"3–5 m.":0,"5+ m.":0}, avgStaz: null };
    }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { total: 0, natList: [], stazGroups: {"0–1 m.":0,"1–3 m.":0,"3–5 m.":0,"5+ m.":0}, avgStaz: null };

    const maxCol = Math.max(dateCol, natCol, stazCol);
    const data   = sh.getRange(2, 1, lastRow - 1, maxCol).getValues();

    const natStats   = {};
    const stazGroups = { "0–1 m.": 0, "1–3 m.": 0, "3–5 m.": 0, "5+ m.": 0 };
    let totalStaz = 0, stazCount = 0, total = 0;
    const rows = [];

    data.forEach(r => {
      const dateVal = r[dateCol - 1];
      const nat     = String(r[natCol  - 1] || "").trim();
      const stazRaw = r[stazCol - 1];
      const staz    = parseFloat(String(stazRaw || "").replace(",", "."));

      // Jei tautybė tuščia – priskiriam "Nenurodyta"
      const natFinal = nat || "Nenurodyta";

      // Datos filtras – tik jei data yra datos tipo
      if (dateVal instanceof Date) {
        if (from && dateVal < from) return;
        if (to   && dateVal > to)   return;
      }
      // Jei langelis tuščias arba ne data – rodome be filtravimo

      total++;

      // Tautybė
      if (!natStats[natFinal]) natStats[natFinal] = { count: 0, totalStaz: 0, stazList: [] };
      natStats[natFinal].count++;

      // Stažas
      const stazValid = !isNaN(staz) && staz >= 0;
      if (stazValid) {
        natStats[natFinal].totalStaz += staz;
        natStats[natFinal].stazList.push(staz);
        totalStaz += staz;
        stazCount++;

        if      (staz < 1) stazGroups["0–1 m."]++;
        else if (staz < 3) stazGroups["1–3 m."]++;
        else if (staz < 5) stazGroups["3–5 m."]++;
        else               stazGroups["5+ m."]++;
      }

      rows.push({ nat: natFinal, staz: stazValid ? staz : null, date: dateVal instanceof Date ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd") : "" });
    });

    // Tautybių statistika – vidurkiai
    const natList = Object.keys(natStats).map(natKey => {
      const d = natStats[natKey];
      const avgS = d.stazList.length ? Math.round(d.totalStaz / d.stazList.length * 100) / 100 : null;
      return { nat: natKey, count: d.count, avgStaz: avgS, totalStaz: d.totalStaz };
    }).sort((a, b) => b.count - a.count);

    // Pagal mėnesį (kalendorinė diagrama)
    const byMonth = {};
    rows.forEach(r => {
      if (!r.date) return;
      const m = r.date.slice(0, 7); // "2026-04"
      byMonth[m] = (byMonth[m] || 0) + 1;
    });
    const monthList = Object.keys(byMonth).sort().map(m => ({ month: m, count: byMonth[m] }));

    return {
      total,
      natList,
      stazGroups,
      avgStaz: stazCount ? Math.round(totalStaz / stazCount * 100) / 100 : null,
      rows,
      monthList
    };
  }

  return {
    dirbantys: parseSheet("DirbantysVairuotojai", 5, 10, 11), // E=5, J=10, K=11
    atleisti:  parseSheet("AtleistiVairuotojai",  6, 10, 11), // F=6, J=10, K=11
    dateFrom:  dateFrom || "",
    dateTo:    dateTo   || "",
    generated: new Date().toLocaleString("lt-LT")
  };
}


/** ===== Naujokų analizė ===== */
function getNaujokai(dateFrom, dateTo) {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh   = ss.getSheetByName("NAUJOKAI");

  if (!sh) {
    const available = ss.getSheets().map(s => s.getName()).join(", ");
    return { error: "Lapas \"NAUJOKAI\" nerastas. Esami lapai: " + available };
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { total: 0, rows: [], natStats: {}, testStats: {}, lStats: {}, pStats: {} };

  const from = dateFrom ? new Date(dateFrom) : null;
  const to   = dateTo   ? new Date(dateTo)   : null;
  if (to) to.setHours(23, 59, 59, 999);

  // A=1, C=3, I=9, L=12, M=13, P=16
  const data = sh.getRange(2, 1, lastRow - 1, 16).getValues();

  const natStats  = {};  // tautybė -> count
  const testVals  = [];  // testo rezultatai (skaičiai)
  const lStats    = { "Ž": 0, "P": 0, "Grįžo": 0, "Kita": 0 };
  const pStats    = { "Geras": 0, "Vidutinis": 0, "Blogas": 0, "Nenurodyta": 0 };
  const mStats    = { "T": 0, "N": 0, "Nenurodyta": 0 };
  const byMonth   = {};
  const rows      = [];
  let total = 0;

  data.forEach(r => {
    const dateVal = r[0];  // A
    const nat     = String(r[2]  || "").trim();  // C
    const testRaw = r[8];                        // I
    const lRaw    = String(r[11] || "").trim();  // L
    const mRaw    = String(r[12] || "").trim();  // M
    const pRaw    = String(r[15] || "").trim();  // P

    // Datos filtras
    if (dateVal instanceof Date) {
      if (from && dateVal < from) return;
      if (to   && dateVal > to)   return;
    }

    // Bent tautybė arba data turi būti
    if (!nat && !(dateVal instanceof Date)) return;

    total++;

    // Tautybė
    const natF = nat || "Nenurodyta";
    natStats[natF] = (natStats[natF] || 0) + 1;

    // Testo rezultatas
    const testNum = parseFloat(String(testRaw || "").replace(",", "."));
    if (!isNaN(testNum)) testVals.push(testNum);

    // L reikšmė
    const lUp = lRaw.toUpperCase();
    if      (lUp === "Ž" || lUp === "Z")  lStats["Ž"]++;
    else if (lUp === "P")                  lStats["P"]++;
    else if (lRaw.toLowerCase().includes("grįžo") || lRaw.toLowerCase().includes("grizо") || lRaw.toLowerCase().includes("gri")) lStats["Grįžo"]++;
    else if (lRaw)                         lStats["Kita"]++;

    // M tinkamumas
    const mUp = mRaw.toUpperCase();
    if      (mUp === "T") mStats["T"]++;
    else if (mUp === "N") mStats["N"]++;
    else                  mStats["Nenurodyta"]++;

    // P kokybė
    const pLow = pRaw.toLowerCase();
    if      (pLow.includes("ger"))  pStats["Geras"]++;
    else if (pLow.includes("vid"))  pStats["Vidutinis"]++;
    else if (pLow.includes("blog")) pStats["Blogas"]++;
    else                            pStats["Nenurodyta"]++;

    // Pagal mėnesį
    if (dateVal instanceof Date) {
      const m = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM");
      byMonth[m] = (byMonth[m] || 0) + 1;
    }

    rows.push({
      date:  dateVal instanceof Date ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "",
      nat:   natF,
      test:  !isNaN(testNum) ? testNum : null,
      l:     lRaw || "—",
      p:     pRaw || "—"
    });
  });

  // Tautybių sąrašas
  const natList = Object.keys(natStats)
    .map(n => ({ nat: n, count: natStats[n] }))
    .sort((a, b) => b.count - a.count);

  // Testo statistika
  const testAvg  = testVals.length ? Math.round(testVals.reduce((a,b)=>a+b,0) / testVals.length * 100) / 100 : null;
  const testMin  = testVals.length ? Math.min(...testVals) : null;
  const testMax  = testVals.length ? Math.max(...testVals) : null;

  // Testo pasiskirstymas grupėmis
  const testGroups = { "0–40": 0, "41–60": 0, "61–80": 0, "81–100": 0 };
  testVals.forEach(v => {
    if      (v <= 40) testGroups["0–40"]++;
    else if (v <= 60) testGroups["41–60"]++;
    else if (v <= 80) testGroups["61–80"]++;
    else              testGroups["81–100"]++;
  });

  const monthList = Object.keys(byMonth).sort().map(m => ({ month: m, count: byMonth[m] }));

  return {
    total,
    natList,
    testAvg, testMin, testMax,
    testGroups,
    lStats,
    mStats,
    pStats,
    monthList,
    dateFrom: dateFrom || "",
    dateTo:   dateTo   || "",
    generated: new Date().toLocaleString("lt-LT")
  };
}

function saveNote(payload) {
  payload = payload || {};
  const cats = String(payload.categories||"").split(/[;,]+/).map(s=>s.trim()).filter(Boolean);
  const mapped = {
    instructorName: String(payload.instructor||"").trim(),
    driverName:     String(payload.driverName||"").trim(),
    caseId:         String(payload.caseId||"").trim(),
    ioDirection:    String(payload.ioDirection||"").trim().toUpperCase(),
    risk:           "VIDUTINĖ",
    categories:     cats,
    outcome:        String(payload.outcome||"").trim(),
    notes:          String(payload.notes||"").trim()
  };
  const res = saveCaseEntry(mapped);
  return { ok: true, message: `Išsaugota. Stage: ${res.stage}. Sekantis kontaktas: ${res.nextContactAt || "—"}`, caseId: res.caseId || "" };
}
