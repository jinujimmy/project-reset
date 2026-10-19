function applicationsRoot_() {
  return DriveApp.getFolderById(CFG.APPLICATIONS_FOLDER_ID);
}

function findChildFolder_(parent, name) {
  var want = String(name || "").toLowerCase().trim();
  if (!want) return null;
  var exact = parent.getFoldersByName(name);
  if (exact.hasNext()) return exact.next();
  var it = parent.getFolders();
  while (it.hasNext()) {
    var folder = it.next();
    if (String(folder.getName() || "").toLowerCase().trim() === want) return folder;
  }
  return null;
}

function findChildFile_(parent, name) {
  var want = String(name || "").toLowerCase().trim();
  if (!want) return null;
  var exact = parent.getFilesByName(name);
  if (exact.hasNext()) return exact.next();
  var it = parent.getFiles();
  while (it.hasNext()) {
    var file = it.next();
    if (String(file.getName() || "").toLowerCase().trim() === want) return file;
  }
  return null;
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
    "Match: " + (job.matchPercent === "" || job.matchPercent == null ? "n/a" : job.matchPercent + "%"),
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
    "reasons",
    "feedback"
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

function normalizeHeaderName_(value) {
  var t = String(value || "")
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (t === "not applied") return "applied";
  return t;
}

function ensureHeaderRow_(sheet, headers) {
  var width = headers.length;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, width).setValues([headers]);
    sheet.getRange(1, 1, 1, width).clearDataValidations().setFontWeight("bold");
    sheet.setFrozenRows(1);
    return;
  }
  var last = Math.max(sheet.getLastColumn(), width);
  var header = sheet.getRange(1, 1, 1, last).getValues()[0];
  var first = normalizeHeaderName_(header[0]);
  var expected = normalizeHeaderName_(headers[0]);
  var row1IsData = /^\d{4}-\d{2}-\d{2}/.test(String(header[0] || "").trim()) || first === "yes" || first === "no";
  if (row1IsData || (first && first !== expected)) {
    if (first !== expected) sheet.insertRowBefore(1);
  }
  sheet.getRange(1, 1, 1, width).setValues([headers]);
  sheet.getRange(1, 1, 1, width).clearDataValidations().setFontWeight("bold");
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
    "email_subject",
    "easy_apply",
    "source",
    "feedback",
    "applied",
    "drive_doc"
  ];
}

function reservedTabName_(name) {
  return (
    name === CFG.SUGGESTIONS_TAB ||
    name === CFG.STATUS_TAB ||
    name === CFG.LINKEDIN_TAB ||
    name === (CFG.CATALOG_TAB || "Catalog") ||
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
  ensureFeedbackDropdown_(suggestions, suggestionHeaders_().length);

  var linkedinName = CFG.LINKEDIN_TAB || "Suggestions - Linkedin";
  var linkedin = ss.getSheetByName(linkedinName);
  if (!linkedin) {
    linkedin = ss.insertSheet(linkedinName);
  }
  ensureHeaderRow_(linkedin, linkedinHeaders_());
  compactLinkedinBlankRows_(linkedin);
  var appliedCol = headerColIndex_(linkedin, "applied") + 1;
  var feedbackCol = headerColIndex_(linkedin, "feedback") + 1;
  ensureFeedbackDropdown_(linkedin, feedbackCol > 0 ? feedbackCol : 11);
  ensureAppliedDropdown_(linkedin, appliedCol > 0 ? appliedCol : 12);

  var catalog = ensureCatalogTab_(ss);
  return { ss: ss, suggestions: suggestions, status: status, linkedin: linkedin, catalog: catalog };
}

function dataRowCount_(sheet) {
  return Math.max(sheet.getLastRow() - 1, 1);
}

function ensureFeedbackDropdown_(sheet, col) {
  if (!col || col < 1) return;
  var n = dataRowCount_(sheet);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["skip company", "skip role", "expired", "apply later"], true)
    .setAllowInvalid(true)
    .setHelpText("skip company = hide this employer. skip role = hide this job title. expired = hide this posting. apply later = keep the row, do not skip next run.")
    .build();
  sheet.getRange(2, col, n, 1).setDataValidation(rule);
}

function ensureAppliedDropdown_(sheet, col) {
  if (!col || col < 1) return;
  var n = dataRowCount_(sheet);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["No", "Yes"], true)
    .setAllowInvalid(true)
    .setHelpText("Manual: Yes = applied, No = not applied. Import does not change this.")
    .build();
  sheet.getRange(2, col, n, 1).setDataValidation(rule);
  var urlCol = headerColIndex_(sheet, "url");
  var companyCol = headerColIndex_(sheet, "company");
  var width = Math.max(sheet.getLastColumn(), col);
  var rows = sheet.getRange(2, 1, n, width).getValues();
  var appliedVals = [];
  var wrote = false;
  for (var i = 0; i < rows.length; i++) {
    var hasJob =
      (urlCol >= 0 && String(rows[i][urlCol] || "").trim()) ||
      (companyCol >= 0 && String(rows[i][companyCol] || "").trim());
    var cur = String(rows[i][col - 1] || "").trim();
    var low = cur.toLowerCase();
    var next = cur;
    if (!hasJob) {
      if (cur) {
        next = "";
        wrote = true;
      }
    } else if (!low || low === "not applied" || low === "no") next = "No";
    else if (low === "applied" || low === "yes") next = "Yes";
    if (next !== cur) wrote = true;
    appliedVals.push([next]);
  }
  if (wrote) sheet.getRange(2, col, n, 1).setValues(appliedVals);
}

function compactLinkedinBlankRows_(sheet) {
  if (!sheet) return;
  var last = sheet.getLastRow();
  if (last < 2) return;
  var urlCol = headerColIndex_(sheet, "url");
  var companyCol = headerColIndex_(sheet, "company");
  var jobIdCol = headerColIndex_(sheet, "job_id");
  var values = sheet.getRange(2, 1, last - 1, Math.max(sheet.getLastColumn(), 1)).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var url = urlCol >= 0 ? String(values[i][urlCol] || "").trim() : "";
    var company = companyCol >= 0 ? String(values[i][companyCol] || "").trim() : "";
    var jobId = jobIdCol >= 0 ? String(values[i][jobIdCol] || "").trim() : "";
    if (!url && !company && !jobId) sheet.deleteRow(i + 2);
  }
}

function isLinkedinAppliedYes_(value) {
  var v = String(value || "").trim().toLowerCase();
  return v === "yes" || v === "applied";
}

function ensureJdInboxFolder_() {
  var root = applicationsRoot_();
  return findChildFolder_(root, "_Inbox") || root.createFolder("_Inbox");
}

function looksLikeFullJd_(text) {
  var t = String(text || "").replace(/\s+/g, " ").trim();
  if (t.length < 700) return false;
  if (/sign in|log in to continue|verify you are human|jobright instant alert|you may be a fit|always be the first to apply/i.test(t)) {
    return false;
  }
  return true;
}

function fetchPublicJd_(url) {
  var u = String(url || "");
  if (!/indeed\.com|glassdoor\.com/i.test(u)) return "";
  try {
    var res = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() >= 400) return "";
    var text = htmlToAlertText_(res.getContentText() || "").slice(0, 12000);
    return looksLikeFullJd_(text) ? text : "";
  } catch (e) {
    return "";
  }
}

function scoreInboxName_(filename, job) {
  var n = String(filename || "").toLowerCase();
  var company = String(job.company || "").toLowerCase().trim();
  var role = String(job.role || "").toLowerCase().trim();
  var score = 0;
  if (job.jobId && n.indexOf(String(job.jobId).toLowerCase()) !== -1) score += 6;
  if (job.url) {
    var m = String(job.url).match(/[?&](?:jl|jk)=([a-z0-9]+)/i) || String(job.url).match(/\/info\/([a-f0-9]+)/i);
    if (m && n.indexOf(m[1].toLowerCase()) !== -1) score += 6;
  }
  if (company.length >= 4 && n.indexOf(company) !== -1) score += 3;
  var first = company.split(/\s+/)[0] || "";
  if (first.length >= 4 && n.indexOf(first) !== -1) score += 2;
  var words = role.replace(/[^a-z0-9\s]/g, " ").split(/\s+/).filter(function (w) {
    return w.length > 3;
  }).slice(0, 5);
  var hits = 0;
  words.forEach(function (w) {
    if (n.indexOf(w) !== -1) hits += 1;
  });
  if (hits >= 2) score += 2;
  return score;
}

function linkedinRowJob_(sheet, rowNumber) {
  var width = Math.max(sheet.getLastColumn(), linkedinHeaders_().length);
  var header = sheet.getRange(1, 1, 1, width).getValues()[0];
  function col(name) {
    var want = String(name).toLowerCase();
    for (var i = 0; i < header.length; i++) {
      if (String(header[i] || "").toLowerCase() === want) return i;
    }
    return -1;
  }
  var row = sheet.getRange(rowNumber, 1, 1, width).getValues()[0];
  return {
    row: rowNumber,
    company: String(row[col("company")] || "").trim(),
    role: String(row[col("role")] || "").trim(),
    url: String(row[col("url")] || "").trim(),
    jobId: String(row[col("job_id")] || "").trim(),
    location: String(row[col("location")] || "").trim(),
    applied: row[col("applied")],
    driveDoc: col("drive_doc") >= 0 ? String(row[col("drive_doc")] || "").trim() : "",
    docCol: col("drive_doc")
  };
}

function stampLinkedinDriveDoc_(sheet, rowNumber, docCol, url) {
  if (docCol >= 0 && url) sheet.getRange(rowNumber, docCol + 1).setValue(url);
}

function placeInboxPdf_(file, company, role) {
  var folder = ensureCompanyFolder_(company);
  var dest = role + ".pdf";
  if (file.getName() !== dest) file.setName(dest);
  folder.addFile(file);
  var inbox = findChildFolder_(applicationsRoot_(), "_Inbox");
  if (inbox) {
    try {
      inbox.removeFile(file);
    } catch (e) {}
  }
  return file.getUrl();
}

function findInboxPdfForJob_(job) {
  var inbox = findChildFolder_(applicationsRoot_(), "_Inbox");
  if (!inbox) return null;
  var it = inbox.getFiles();
  var best = null;
  var bestScore = 0;
  var ties = 0;
  while (it.hasNext()) {
    var file = it.next();
    var mime = String(file.getMimeType() || "");
    if (mime && mime !== MimeType.PDF && !/\.pdf$/i.test(file.getName())) continue;
    var score = scoreInboxName_(file.getName(), job);
    if (score > bestScore) {
      best = file;
      bestScore = score;
      ties = 1;
    } else if (score === bestScore && score > 0) {
      ties += 1;
    }
  }
  if (!best || bestScore < 3 || ties > 1) return null;
  return best;
}

function copyLinkedinJdForRow_(sheet, rowNumber) {
  if (rowNumber < 2) return "";
  var job = linkedinRowJob_(sheet, rowNumber);
  if (!isLinkedinAppliedYes_(job.applied)) return "";
  if (job.driveDoc) return job.driveDoc;
  if (!job.company || !job.role) return "";
  if (docAlreadyExists_(job.company, job.role)) {
    var folder = findChildFolder_(applicationsRoot_(), job.company);
    var existing = folder ? findChildFile_(folder, job.role) : null;
    var existingUrl = existing ? existing.getUrl() : "";
    stampLinkedinDriveDoc_(sheet, rowNumber, job.docCol, existingUrl);
    return existingUrl;
  }
  var inboxFile = findInboxPdfForJob_(job);
  if (inboxFile) {
    var pdfUrl = placeInboxPdf_(inboxFile, job.company, job.role);
    stampLinkedinDriveDoc_(sheet, rowNumber, job.docCol, pdfUrl);
    return pdfUrl;
  }
  var desc = fetchPublicJd_(job.url);
  if (!desc) return "";
  var driveUrl = createRoleDoc_(
    {
      title: job.role,
      company: job.company,
      location: job.location,
      url: job.url,
      matchPercent: "",
      matchReasons: ["full posting from the public job page"],
      description: desc
    },
    torontoParts_(new Date()).date
  );
  stampLinkedinDriveDoc_(sheet, rowNumber, job.docCol, driveUrl);
  return driveUrl;
}

function fileInboxJdPdfs() {
  ensureJdInboxFolder_();
  var sheet = ensureTrackerTabs_().linkedin;
  var appliedCol = headerColIndex_(sheet, "applied");
  var last = sheet.getLastRow();
  var filed = 0;
  var unmatched = 0;
  if (appliedCol < 0 || last < 2) {
    SpreadsheetApp.getActive().toast("Set up tabs first.");
    return;
  }
  var inbox = findChildFolder_(applicationsRoot_(), "_Inbox");
  var files = inbox ? inbox.getFiles() : null;
  if (!files || !files.hasNext()) {
    SpreadsheetApp.getActive().toast("Applications/_Inbox is empty. Print the posting to PDF and drop it there.");
    return;
  }
  while (files.hasNext()) {
    var file = files.next();
    var jobs = [];
    for (var r = 2; r <= last; r++) {
      var job = linkedinRowJob_(sheet, r);
      if (!isLinkedinAppliedYes_(job.applied) || job.driveDoc) continue;
      jobs.push(job);
    }
    var best = null;
    var bestScore = 0;
    var ties = 0;
    jobs.forEach(function (job) {
      var score = scoreInboxName_(file.getName(), job);
      if (score > bestScore) {
        best = job;
        bestScore = score;
        ties = 1;
      } else if (score === bestScore && score > 0) {
        ties += 1;
      }
    });
    if (!best || bestScore < 3 || ties > 1) {
      unmatched += 1;
      continue;
    }
    var url = placeInboxPdf_(file, best.company, best.role);
    stampLinkedinDriveDoc_(sheet, best.row, best.docCol, url);
    filed += 1;
  }
  SpreadsheetApp.getActive().toast(
    filed
      ? "Filed " + filed + " PDF" + (filed === 1 ? "" : "s") + " into company folders" + (unmatched ? ". " + unmatched + " left in _Inbox (rename with company name)." : "")
      : "Could not match inbox PDFs. Rename them like “Tokyo Smoke Data Analyst.pdf”."
  );
}

function copyAppliedLinkedinJds() {
  fileInboxJdPdfs();
  var sheet = ensureTrackerTabs_().linkedin;
  var appliedCol = headerColIndex_(sheet, "applied");
  var docCol = headerColIndex_(sheet, "drive_doc");
  if (appliedCol < 0) return;
  var values = sheet.getDataRange().getValues();
  var filed = 0;
  var missing = 0;
  for (var i = 1; i < values.length; i++) {
    if (!isLinkedinAppliedYes_(values[i][appliedCol])) continue;
    if (docCol >= 0 && String(values[i][docCol] || "").trim()) continue;
    if (copyLinkedinJdForRow_(sheet, i + 1)) filed += 1;
    else missing += 1;
  }
  SpreadsheetApp.getActive().toast(
    (filed ? "Filed " + filed + " full JD" + (filed === 1 ? "" : "s") + ". " : "") +
      (missing ? missing + " applied row(s) still need a print-to-PDF in Applications/_Inbox." : filed ? "" : "No new full JDs.")
  );
}

function headerColIndex_(sheet, name) {
  var last = Math.max(sheet.getLastColumn(), 1);
  var header = sheet.getRange(1, 1, 1, last).getValues()[0];
  var want = normalizeHeaderName_(name);
  for (var i = 0; i < header.length; i++) {
    if (normalizeHeaderName_(header[i]) === want) return i;
  }
  return -1;
}

function normalizeFeedbackLabel_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function collectFeedbackSkips_() {
  var tabs = ensureTrackerTabs_();
  var companies = {};
  var roles = {};
  var expiredUrls = {};
  var expiredJobs = {};
  function ingest_(sheet, companyKey, roleKey) {
    if (!sheet) return;
    var fb = headerColIndex_(sheet, "feedback");
    var companyCol = headerColIndex_(sheet, companyKey);
    var roleCol = headerColIndex_(sheet, roleKey);
    var urlCol = headerColIndex_(sheet, "url");
    if (fb < 0 || companyCol < 0 || roleCol < 0) return;
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var label = normalizeFeedbackLabel_(values[i][fb]);
      var company = String(values[i][companyCol] || "").toLowerCase().trim();
      var role = String(values[i][roleCol] || "").toLowerCase().replace(/\s+/g, " ").trim();
      if (label === "skip company" && company) companies[company] = true;
      if (label === "skip role" && role) roles[role] = true;
      if (label === "expired") {
        if (company && role) expiredJobs[company + "\t" + role] = true;
        if (urlCol >= 0) {
          var url = normalizeJobUrl_(values[i][urlCol]);
          if (url) expiredUrls[url] = true;
        }
      }
    }
  }
  ingest_(tabs.linkedin, "company", "role");
  ingest_(tabs.suggestions, "company", "role");
  return { companies: companies, roles: roles, expiredUrls: expiredUrls, expiredJobs: expiredJobs };
}

function normalizeJobUrl_(url) {
  var u = String(url || "").trim();
  if (!u) return "";
  u = u.replace(/[?&#]utm_[^&#]*/gi, "");
  u = u.replace(/[?&]$/, "");
  return u.toLowerCase();
}

function isExpiredPosting_(company, role, url, skips) {
  var bag = skips || collectFeedbackSkips_();
  var key =
    String(company || "").toLowerCase().trim() +
    "\t" +
    String(role || "").toLowerCase().replace(/\s+/g, " ").trim();
  if (bag.expiredJobs && bag.expiredJobs[key]) return true;
  var n = normalizeJobUrl_(url);
  return Boolean(n && bag.expiredUrls && bag.expiredUrls[n]);
}

function isSkippedByFeedback_(company, role, skips, url) {
  var bag = skips || collectFeedbackSkips_();
  var c = String(company || "").toLowerCase().trim();
  var r = String(role || "").toLowerCase().replace(/\s+/g, " ").trim();
  if (c) {
    var names = Object.keys(bag.companies || {});
    for (var i = 0; i < names.length; i++) {
      if (c.indexOf(names[i]) !== -1 || names[i].indexOf(c) !== -1) return true;
    }
  }
  if (r && bag.roles && bag.roles[r]) return true;
  return isExpiredPosting_(company, role, url, bag);
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

function alreadyRecommendedKey_(company, role) {
  return String(company || "").toLowerCase().trim() + "\t" + String(role || "").toLowerCase().trim();
}

function collectAppliedListingKeys_() {
  var tabs = ensureTrackerTabs_();
  var keys = {};
  var urls = {};
  function add_(company, role, url) {
    var key = alreadyRecommendedKey_(company, role);
    if (key !== "\t") keys[key] = true;
    var n = normalizeJobUrl_(url);
    if (n) urls[n] = true;
  }
  var linkedin = tabs.linkedin;
  if (linkedin) {
    var appliedCol = headerColIndex_(linkedin, "applied");
    var companyCol = headerColIndex_(linkedin, "company");
    var roleCol = headerColIndex_(linkedin, "role");
    var urlCol = headerColIndex_(linkedin, "url");
    var values = linkedin.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (appliedCol >= 0 && !isLinkedinAppliedYes_(values[i][appliedCol])) continue;
      if (appliedCol < 0) continue;
      add_(
        companyCol >= 0 ? values[i][companyCol] : "",
        roleCol >= 0 ? values[i][roleCol] : "",
        urlCol >= 0 ? values[i][urlCol] : ""
      );
    }
  }
  var status = tabs.status;
  if (status) {
    var rows = status.getDataRange().getValues();
    for (var s = 1; s < rows.length; s++) {
      if (!progressMeansApplied_(String(rows[s][5] || ""))) continue;
      add_(rows[s][1], rows[s][2], rows[s][7]);
    }
  }
  return { keys: keys, urls: urls };
}

function isAlreadyAppliedListing_(company, role, url, bag) {
  var applied = bag || collectAppliedListingKeys_();
  var key = alreadyRecommendedKey_(company, role);
  if (key !== "\t" && applied.keys && applied.keys[key]) return true;
  var n = normalizeJobUrl_(url);
  return Boolean(n && applied.urls && applied.urls[n]);
}

function collectAlreadyRecommendedKeys_() {
  var tabs = ensureTrackerTabs_();
  var keys = {};
  function ingest_(sheet, companyCol, roleCol) {
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var key = alreadyRecommendedKey_(values[i][companyCol], values[i][roleCol]);
      if (key !== "\t") keys[key] = true;
    }
  }
  ingest_(tabs.suggestions, 3, 4);
  ingest_(tabs.status, 1, 2);
  return keys;
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
