/**
 * Google Apps Script — CompModels DB
 *
 * SETUP (one time):
 *  1. Create a Google Sheet. Add two tabs:
 *       "Submissions"  — incoming form entries (private)
 *       "Verified"     — curated entries shown on the dashboard (public read)
 *  2. Extensions → Apps Script → paste this file, save.
 *  3. Run addHeaders() once (select from the function dropdown, click Run).
 *  4. Deploy → New deployment → Web App
 *       Execute as: Me
 *       Who has access: Anyone
 *  5. Copy the Web App URL → paste into submit.qmd as SUBMIT_URL.
 *  6. Publish the Verified sheet as CSV:
 *       File → Share → Publish to web → Sheet: Verified → CSV → Publish
 *       Copy that URL → paste into database.qmd as SHEET_CSV_URL.
 *
 * CURATION WORKFLOW:
 *  - New submissions land in "Submissions" with status = "pending".
 *  - Review them, then call approveRow(rowNumber) from the Script editor
 *    (or add a button/menu item — see addReviewMenu() below).
 *  - Approved entries are copied to "Verified" and the original row is
 *    marked status = "approved".
 */

// ── Column schema ─────────────────────────────────────────────────────────────
const COLUMNS = [
  "timestamp", "author", "year", "doi_url", "behavior", "behavior_category",
  "theory_name", "theory_group", "architecture",
  "mechanism_derived", "mechanism_rationale",
  "cross_arch_validation", "cross_arch_method", "delta_elpd",
  "jitai_derived", "jitai_who", "jitai_when",
  "study_design", "n_participants", "study_duration_days",
  "ema_assessments_per_day", "outcome_type", "notes",
  "submitter_email", "status"
];

// ── Receive form submissions ──────────────────────────────────────────────────
function doPost(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Submissions");
    if (!sheet) throw new Error("Sheet 'Submissions' not found");

    const data = JSON.parse(e.postData.contents);
    const row  = COLUMNS.map(col => {
      if (col === "timestamp") return new Date().toISOString();
      if (col === "status")    return "pending";
      return data[col] !== undefined ? String(data[col]) : "";
    });
    sheet.appendRow(row);

    // Notify maintainer (uncomment and add your email once deployed)
    // MailApp.sendEmail(
    //   "christopher.jones@medma.uni-heidelberg.com",
    //   "[CompModels DB] New submission: " + (data.author || "?") + " (" + (data.year || "?") + ")",
    //   "Review at: https://docs.google.com/spreadsheets/d/" +
    //     SpreadsheetApp.getActiveSpreadsheet().getId()
    // );

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", db: "CompModels DB" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── One-time setup ────────────────────────────────────────────────────────────
function addHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ["Submissions", "Verified"].forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const range = sheet.getRange(1, 1, 1, COLUMNS.length);
    range.setValues([COLUMNS]);
    range.setFontWeight("bold");
    range.setBackground("#2c3e50");
    range.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  });

  // Seed the Verified sheet with the two confirmed entries from Jones (2025/2022)
  const verified = ss.getSheetByName("Verified");
  const today = new Date().toISOString().split("T")[0];

  verified.appendRow([
    today,
    "Jones et al.", "2025",
    "https://doi.org/10.1080/08870446.2025.000000",
    "Snacking", "Dietary behavior",
    "Temporal Self-Regulation Theory (TST)", "Self-regulation / Cybernetic",
    "Cybernetic difference equation",
    "Yes",
    "TST sequentiality identified as dominant mechanism: high-frequency goal-directed behavior generates feedback across episodes.",
    "Yes", "LOO-CV", "-10.8",
    "Yes",
    "Posterior non-corrector probability P(β_fail > −0.01 | data) as continuous targeting score",
    "Goal-attainment failure windows (GA below person mean)",
    "EMA", "91", "14", "5",
    "Binary (snacking event)",
    "Person-level β_fail heterogeneity (SD = 0.091). Cross-architecture comparison vs. dual-process logistic.",
    "", "verified"
  ]);

  verified.appendRow([
    today,
    "Jones & Schüz", "2022",
    "https://doi.org/10.1111/bjhp.12580",
    "Smoking", "Substance use",
    "Temporal Self-Regulation Theory (TST)", "Self-regulation / Cybernetic",
    "Dual-process logistic",
    "Yes",
    "TST synchronicity identified as dominant mechanism: within-episode cue competition formalised as dual-process.",
    "Yes", "LOO-CV", "-12.3",
    "Yes",
    "CPD below median (lighter smokers show stronger cue effects)",
    "High social cue + high urge + low intention windows (~6.8% of assessments)",
    "EMA", "49", "21", "5",
    "Binary (smoking event)",
    "CPD moderates cue dependence (w3 = −0.127). TST interplay term (Cues × Intention) significant.",
    "", "verified"
  ]);

  SpreadsheetApp.getUi().alert("Headers and seed data added to both sheets.");
}

// ── Curation: approve a pending row ──────────────────────────────────────────
/**
 * Call approveRow() from the custom menu (see addReviewMenu),
 * or directly: approveRow(2) to approve row 2 of Submissions.
 */
function approveRow(rowNumber) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const submissions = ss.getSheetByName("Submissions");
  const verified    = ss.getSheetByName("Verified");

  if (!rowNumber) {
    // If called from menu: use the active cell's row
    rowNumber = submissions.getActiveCell().getRow();
  }
  if (rowNumber <= 1) {
    SpreadsheetApp.getUi().alert("Select a data row (not the header).");
    return;
  }

  const row = submissions.getRange(rowNumber, 1, 1, COLUMNS.length).getValues()[0];
  const statusIdx = COLUMNS.indexOf("status");

  if (row[statusIdx] === "approved") {
    SpreadsheetApp.getUi().alert("Row already approved.");
    return;
  }

  // Copy to Verified with status = "verified"
  row[statusIdx] = "verified";
  verified.appendRow(row);

  // Mark original as approved
  submissions.getRange(rowNumber, statusIdx + 1).setValue("approved");
  submissions.getRange(rowNumber, statusIdx + 1).setBackground("#c8f0d8");

  SpreadsheetApp.getUi().alert(
    "Row " + rowNumber + " approved and copied to Verified sheet."
  );
}

// ── Custom menu in the spreadsheet UI ────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("CompModels DB")
    .addItem("Approve selected row", "approveRow")
    .addItem("Setup / re-add headers", "addHeaders")
    .addToUi();
}
