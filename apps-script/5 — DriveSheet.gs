function applicationsRoot_() {
  return DriveApp.getFolderById(CFG.APPLICATIONS_FOLDER_ID);
}

function findChildFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

function findChildFile_(parent, name) {
  var it = parent.getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}

function ensureCompanyFolder_(company) {
  var root = applicationsRoot_();
  var existing = findChildFolder_(root, company);
  if (existing) return existing;
  return root.createFolder(company);
}

function docAlreadyExists_(company, role) {
  var folder = findChildFolder_(applicationsRoot_(), company);
  if (!folder) return false;
  return Boolean(findChildFile_(folder, role));
}

function createRoleDoc_(job, applyDay) {
  var folder = ensureCompanyFolder_(job.company);
  var existing = findChildFile_(folder, job.title);
  if (existing) return existing.getUrl();
  var body = [
    job.title,
    job.company,
    job.location,
    job.url,
    "Match: " + job.matchPercent + "%",
    "Recommended apply: " + applyDay,
    "Reasons: " + (job.matchReasons || []).join("; "),
    "",
    String(job.description || "").slice(0, 12000)
  ].join("\n");
  var doc = DocumentApp.create(job.title);
  doc.getBody().setText(body);
  doc.saveAndClose();
  var gdoc = DriveApp.getFileById(doc.getId());
  folder.addFile(gdoc);
  DriveApp.getRootFolder().removeFile(gdoc);
  return gdoc.getUrl();
}

function trackerSheet_() {
  return SpreadsheetApp.openById(CFG.TRACKER_SHEET_ID);
}

function torontoParts_(now) {
  var tz = "America/Toronto";
  return {
    date: Utilities.formatDate(now, tz, "yyyy-MM-dd"),
    time: Utilities.formatDate(now, tz, "HH:mm:ss")
  };
}

function suggestionHeaders_() {
  return [
    "recommendation_date",
    "run_time",
    "run_label",
    "company",
    "role",
    "match_percent",
    "recommended_apply_day",
    "location",
    "url",
    "drive_doc",
    "reasons"
  ];
}

function statusHeaders_() {
  return [
    "status",
    "company",
    "role",
    "match_percent",
    "recommended_apply_day",
    "progress",
    "gmail_signal",
    "url",
    "drive_doc",
    "jd_text"
  ];
}

function ensureHeaderRow_(sheet, headers) {
  var width = headers.length;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, width).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }
  var last = Math.max(sheet.getLastColumn(), width);
  var header = sheet.getRange(1, 1, 1, last).getValues()[0];
  if (String(header[0]).toLowerCase() !== String(headers[0]).toLowerCase()) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, width).setValues([headers]);
  }
  sheet.setFrozenRows(1);
}

function findCompanyRoleRow_(sheet, company, role, companyCol, roleCol) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (
      String(values[i][companyCol]).toLowerCase() === String(company).toLowerCase() &&
      String(values[i][roleCol]).toLowerCase() === String(role).toLowerCase()
    ) {
      return i + 1;
    }
  }
  return 0;
}

function linkedinHeaders_() {
  return [
    "found_date",
    "found_time",
    "company",
    "role",
    "location",
    "url",
    "job_id",
    "email_subject"
  ];
}

function reservedTabName_(name) {
  return (
    name === CFG.SUGGESTIONS_TAB ||
    name === CFG.STATUS_TAB ||
    name === CFG.LINKEDIN_TAB ||
    name === "Folder scores"
  );
}

function ensureTrackerTabs_() {
  var ss = trackerSheet_();
  var status = ss.getSheetByName(CFG.STATUS_TAB);
  if (!status) {
    var first = ss.getSheets()[0];
    if (first && !reservedTabName_(first.getName())) {
      first.setName(CFG.STATUS_TAB);
      status = first;
    } else {
      status = ss.insertSheet(CFG.STATUS_TAB, 0);
    }
  }
  ensureHeaderRow_(status, statusHeaders_());
  if (status.getMaxColumns() >= 10) status.hideColumns(10);

  var suggestions = ss.getSheetByName(CFG.SUGGESTIONS_TAB);
  if (!suggestions) {
    suggestions = ss.insertSheet(CFG.SUGGESTIONS_TAB, 0);
  }
  ensureHeaderRow_(suggestions, suggestionHeaders_());

  var linkedinName = CFG.LINKEDIN_TAB || "Suggestions - Linkedin";
  var linkedin = ss.getSheetByName(linkedinName);
  if (!linkedin) {
    linkedin = ss.insertSheet(linkedinName);
  }
  ensureHeaderRow_(linkedin, linkedinHeaders_());
  return { ss: ss, suggestions: suggestions, status: status, linkedin: linkedin };
}

function appendSuggestionRow_(job, applyDay, driveUrl, runAt, runLabel) {
  var tabs = ensureTrackerTabs_();
  var stamp = torontoParts_(runAt || new Date());
  tabs.suggestions.appendRow([
    stamp.date,
    stamp.time,
    runLabel || "manual",
    job.company,
    job.title,
    job.matchPercent,
    applyDay,
    job.location || "",
    job.url,
    driveUrl || "",
    (job.matchReasons || []).join("; ")
  ]);
}

function ensureApplicationStatusRow_(job, applyDay, driveUrl, progress, gmailSignal) {
  var tabs = ensureTrackerTabs_();
  var sheet = tabs.status;
  if (findCompanyRoleRow_(sheet, job.company, job.title, 1, 2)) return;
  sheet.appendRow([
    "suggested",
    job.company,
    job.title,
    job.matchPercent,
    applyDay,
    progress || "not_applied",
    gmailSignal || "none",
    job.url,
    driveUrl || "",
    String(job.description || "").slice(0, 12000)
  ]);
}

function alreadyRecommended_(company, role) {
  var tabs = ensureTrackerTabs_();
  if (findCompanyRoleRow_(tabs.suggestions, company, role, 3, 4)) return true;
  if (findCompanyRoleRow_(tabs.status, company, role, 1, 2)) return true;
  return docAlreadyExists_(company, role);
}

function progressMeansApplied_(progress) {
  return (
    progress === "applied" ||
    progress === "interview_or_assessment" ||
    progress === "rejected"
  );
}

function jobFromStatusRow_(row) {
  return {
    company: String(row[1] || ""),
    title: String(row[2] || ""),
    matchPercent: row[3],
    location: "",
    url: String(row[7] || ""),
    description: String(row[9] || ""),
    matchReasons: []
  };
}

function copyJdIfApplied_(sheet, rowNumber, row, applyDay) {
  if (String(row[8] || "").trim()) return String(row[8]);
  if (!progressMeansApplied_(String(row[5] || ""))) return "";
  var job = jobFromStatusRow_(row);
  if (!job.company || !job.title) return "";
  var driveUrl = createRoleDoc_(job, applyDay || String(row[4] || ""));
  sheet.getRange(rowNumber, 9).setValue(driveUrl);
  sheet.getRange(rowNumber, 1).setValue("filed");
  stampDriveUrlOnSuggestions_(job.company, job.title, driveUrl);
  return driveUrl;
}

function stampDriveUrlOnSuggestions_(company, role, driveUrl) {
  var sheet = ensureTrackerTabs_().suggestions;
  var row = findCompanyRoleRow_(sheet, company, role, 3, 4);
  if (row) sheet.getRange(row, 10).setValue(driveUrl);
}

function recordSuggestedRole_(job, applyDay, driveUrl, progress, gmailSignal, runAt, runLabel) {
  appendSuggestionRow_(job, applyDay, driveUrl, runAt, runLabel);
  ensureApplicationStatusRow_(job, applyDay, driveUrl, progress, gmailSignal);
}

function inferGmailProgress_(company, role) {
  var q =
    "newer_than:21d (" +
    '"' +
    company +
    '" OR subject:' +
    company.split(" ")[0] +
    ") (applied OR application OR interview OR assessment OR rejected OR \"thank you for applying\" OR \"not moving forward\")";
  try {
    var threads = GmailApp.search(q, 0, 8);
    for (var i = 0; i < threads.length; i++) {
      var subj = String(threads[i].getFirstMessageSubject() || "").toLowerCase();
      var snip = String(threads[i].getMessages()[0].getPlainBody() || "").slice(0, 400).toLowerCase();
      var blob = subj + " " + snip;
      if (role && blob.indexOf(String(role).toLowerCase().slice(0, 18)) === -1 && blob.indexOf(company.toLowerCase()) === -1) {
        continue;
      }
      if (/not moving forward|decided not to|unfortunately|rejected|not be moving/.test(blob)) {
        return { progress: "rejected", signal: threads[i].getFirstMessageSubject() };
      }
      if (/interview|assessment|codesignal/.test(blob)) {
        return { progress: "interview_or_assessment", signal: threads[i].getFirstMessageSubject() };
      }
      if (/thank you for applying|application received|we.ve received your application|successfully applied/.test(blob)) {
        return { progress: "applied", signal: threads[i].getFirstMessageSubject() };
      }
    }
  } catch (e) {
    return { progress: "unknown", signal: "gmail search failed" };
  }
  return { progress: "not_applied", signal: "none" };
}

function refreshGmailProgressOnSheet_() {
  var sheet = ensureTrackerTabs_().status;
  var values = sheet.getDataRange().getValues();
  var filed = 0;
  for (var i = 1; i < values.length; i++) {
    var company = values[i][1];
    var role = values[i][2];
    if (!company) continue;
    var inferred = inferGmailProgress_(String(company), String(role || ""));
    if (inferred.progress !== "not_applied") {
      sheet.getRange(i + 1, 6, 1, 2).setValues([[inferred.progress, inferred.signal]]);
      values[i][5] = inferred.progress;
      values[i][6] = inferred.signal;
    }
    var driveUrl = copyJdIfApplied_(sheet, i + 1, values[i], values[i][4]);
    if (driveUrl && !String(values[i][8] || "").trim()) filed += 1;
  }
  SpreadsheetApp.getActive().toast(
    filed ? "Gmail updated. Copied " + filed + " JD" + (filed === 1 ? "" : "s") + " into Applications." : "Gmail progress updated"
  );
}
