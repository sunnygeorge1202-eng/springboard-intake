// ─────────────────────────────────────────────────────────────
// Springboard Talent — Intake Form Handler
// Deploy as: Web App → Execute as Me → Anyone can access
// ─────────────────────────────────────────────────────────────

const ROOT_FOLDER_NAME = "Springboard Talent Intake";
const SHEET_NAME = "Responses";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const root = getOrCreateFolder(ROOT_FOLDER_NAME, DriveApp.getRootFolder());
    const audioFolder = getOrCreateFolder("Audio Responses", root);
    const sheet = getOrCreateSheet(root);

    // ── Save audio files ──
    const audioLinks = { history: "", target: "", outcome: "" };
    const candidateName = (data.name || "Unknown").replace(/[^a-zA-Z0-9 ]/g, "").trim();
    const dateStr = new Date().toISOString().slice(0, 10);
    const candidateFolderName = candidateName + " — " + dateStr;

    const audioKeys = ["history", "target", "outcome"];
    const audioLabels = {
      history: "career-history",
      target: "target-direction",
      outcome: "outcome"
    };

    let candidateAudioFolder = null;

    for (const key of audioKeys) {
      const b64 = data["audio_" + key];
      const mime = data["audio_" + key + "_mime"] || "audio/webm";
      if (b64) {
        if (!candidateAudioFolder) {
          candidateAudioFolder = getOrCreateFolder(candidateFolderName, audioFolder);
        }
        const ext = mime.includes("ogg") ? ".ogg" : ".webm";
        const filename = audioLabels[key] + ext;
        const bytes = Utilities.base64Decode(b64);
        const blob = Utilities.newBlob(bytes, mime, filename);
        const file = candidateAudioFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        audioLinks[key] = file.getUrl();
      }
    }

    // ── Append row to sheet ──
    const timestamp = data.timestamp || new Date().toISOString();
    const row = [
      timestamp,
      data.name || "",
      data.email || "",
      data.phone || "",
      data.city || "",
      data.linkedin || "",
      data.relocation || "",
      data.reloc_cities || "",
      data.qual || "",
      data.college || "",
      data.gradyr || "",
      data.ug || "",
      data.certs || "",
      data.jobtitle || "",
      data.employer || "",
      data.tenure || "",
      data.fn || "",
      data.industry || "",
      data.yoe || "",
      data.history_text || data.history_transcript || "",
      audioLinks.history,
      data.situation || "",
      data.urgency || "",
      data.target_text || data.target_transcript || "",
      audioLinks.target,
      data.ctc || "",
      data.ectc || "",
      data.comp_flex || "",
      data.outcome_text || data.outcome_transcript || "",
      audioLinks.outcome,
      candidateAudioFolder ? candidateAudioFolder.getUrl() : ""
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("Error: " + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateFolder(name, parent) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

function getOrCreateSheet(rootFolder) {
  // Look for existing spreadsheet in root folder
  const files = rootFolder.getFilesByName(SHEET_NAME);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
    // Move to root folder
    const file = DriveApp.getFileById(ss.getId());
    rootFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    // Write header row
    const headers = [
      "Timestamp", "Name", "Email", "Phone", "City", "LinkedIn",
      "Relocation", "Relocation Cities",
      "Highest Qual", "College", "Year of Passing", "UG Degree", "Certifications",
      "Job Title", "Employer", "Tenure", "Function", "Industry", "Years of Exp",
      "Career History (text)", "Career History (audio)",
      "Situation", "Urgency",
      "Target Direction (text)", "Target Direction (audio)",
      "Current CTC", "Expected CTC", "Comp Flexibility",
      "Outcome (text)", "Outcome (audio)",
      "Audio Folder Link"
    ];
    ss.getActiveSheet().setName(SHEET_NAME);
    ss.getActiveSheet().appendRow(headers);
    // Format header row
    const headerRange = ss.getActiveSheet().getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#085041");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    ss.getActiveSheet().setFrozenRows(1);
  }
  return ss.getSheetByName(SHEET_NAME);
}

// ── Test function — run manually to verify setup ──
function testSetup() {
  const root = getOrCreateFolder(ROOT_FOLDER_NAME, DriveApp.getRootFolder());
  const sheet = getOrCreateSheet(root);
  Logger.log("Root folder: " + root.getUrl());
  Logger.log("Sheet: " + sheet.getParent().getUrl());
  Logger.log("Setup OK");
}
