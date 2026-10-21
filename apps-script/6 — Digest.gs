function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Canada ATS")
    .addItem("Run digest now", "runDigestNow")
    .addItem("Set up Suggestions and Application status tabs", "setupTrackerTabs")
    .addItem("Score Applications folder", "scoreApplicationsFolder")
    .addItem("Refresh Gmail progress", "refreshGmailProgressOnSheet_")
    .addItem("Import LinkedIn job links from Gmail", "importLinkedInJobsFromGmail")
    .addItem("Copy JDs for applied Linkedin rows", "copyAppliedLinkedinJds")
    .addItem("File inbox PDFs into company folders", "fileInboxJdPdfs")
    .addSeparator()
    .addItem("Install 8 AM / 6 PM triggers", "installDigestTriggers")
    .addItem("Remove digest triggers", "removeDigestTriggers")
    .addToUi();
}

function setupTrackerTabs() {
  ensureTrackerTabs_();
  ensureJdInboxFolder_();
  SpreadsheetApp.getActive().toast("Tabs ready. Drop print-to-PDF JDs in Applications/_Inbox.");
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== (CFG.LINKEDIN_TAB || "Suggestions - Linkedin")) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  var row = e.range.getRow();
  if (row < 2) return;
  var statusCol = linkedinApplyStatusCol_(sheet) + 1;
  var dateCol = headerColIndex_(sheet, "applied_date") + 1;
  var col = e.range.getColumn();
  try {
    if (dateCol > 0 && col === dateCol && String(e.value || "").trim()) {
      if (statusCol > 0) sheet.getRange(row, statusCol).setValue("Yes");
      copyLinkedinJdForRow_(sheet, row);
      return;
    }
    if (statusCol > 0 && col === statusCol && isLinkedinAppliedYes_(e.value)) {
      copyLinkedinJdForRow_(sheet, row);
    }
  } catch (err) {}
}

function runMorningDigest() {
  runDigest_("morning");
}

function runEveningDigest() {
  runDigest_("evening");
}

function runDigestNow() {
  runDigest_("manual");
}

function installDigestTriggers() {
  removeDigestTriggers();
  ScriptApp.newTrigger("runMorningDigest").timeBased().atHour(8).everyDays(1).inTimezone("America/Toronto").create();
  ScriptApp.newTrigger("runEveningDigest").timeBased().atHour(18).everyDays(1).inTimezone("America/Toronto").create();
  SpreadsheetApp.getActive().toast("Triggers installed: 8:00 and 18:00 America/Toronto");
}

function removeDigestTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === "runMorningDigest" || fn === "runEveningDigest" || fn === "continueDigest") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function runDigest_(label) {
  ensureTrackerTabs_({ skipCompact: true, skipDropdowns: true });
  if (label !== "continue") {
    try {
      importLinkedInJobsFromGmail(true);
    } catch (e) {}
  }
  var runAt = new Date();
  var started = Date.now();
  var props = PropertiesService.getScriptProperties();
  var cursor = Number(props.getProperty("boardCursor") || "0");
  var boards = activeCatalog_();
  var applyDay = recommendedApplyDay_();
  var newMatches = [];
  var errors = [];
  var considered = 0;
  var end = Math.min(boards.length, cursor + CFG.BOARD_BATCH);
  var feedbackSkips = collectFeedbackSkips_();
  var already = collectAlreadyRecommendedKeys_();
  var applied = collectAppliedListingKeys_();

  for (var i = cursor; i < end; i++) {
    if (Date.now() - started > CFG.MAX_RUNTIME_MS) {
      end = i;
      break;
    }
    var fetched = fetchBoardJobs_(boards[i]);
    if (fetched.error) errors.push(boards[i].company + ": " + fetched.error);
    (fetched.jobs || []).forEach(function (job) {
      if (isAlreadyAppliedListing_(job.company, job.title, job.url, applied)) return;
      if (isSkippedByFeedback_(job.company, job.title, feedbackSkips, job.url)) return;
      if (!isCanadaEligible_(job.location, job.description)) return;
      var scored = scoreAgainstTargetJd_(job.description);
      job.matchPercent = scored.percent;
      job.matchReasons = scored.reasons;
      considered += 1;
      if (scored.percent < CFG.MATCH_THRESHOLD) return;
      if (!isAnalyticsIcJd_(job.description)) return;
      var recKey = alreadyRecommendedKey_(job.company, job.title);
      if (already[recKey]) return;
      if (docAlreadyExists_(job.company, job.title)) {
        already[recKey] = true;
        return;
      }
      already[recKey] = true;
      recordSuggestedRole_(job, applyDay, "", "not_applied", "none", runAt, label);
      newMatches.push({
        company: job.company,
        title: job.title,
        percent: job.matchPercent,
        url: job.url,
        driveUrl: "",
        location: job.location,
        reasons: job.matchReasons
      });
    });
  }

  if (end < boards.length) {
    props.setProperty("boardCursor", String(end));
    var pending = JSON.parse(props.getProperty("pendingMatches") || "[]");
    props.setProperty("pendingMatches", JSON.stringify(pending.concat(newMatches)));
    ScriptApp.newTrigger("continueDigest").timeBased().after(60 * 1000).create();
    return;
  }

  var pendingFinal = JSON.parse(props.getProperty("pendingMatches") || "[]").concat(newMatches);
  props.deleteProperty("boardCursor");
  props.deleteProperty("pendingMatches");
  if (pendingFinal.length) sendDigestEmail_(pendingFinal, applyDay, label, errors);
}

function continueDigest() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "continueDigest") ScriptApp.deleteTrigger(t);
  });
  runDigest_("continue");
}

function sendDigestEmail_(matches, applyDay, label, errors) {
  var lines = [
    "Canada ATS digest (" + label + "). Recommended apply day: " + applyDay + ".",
    "Open the job link, then apply. The JD is copied into Applications after Gmail shows you applied.",
    "Coverage ≥70% and JD body is a DA/PA/BI/insights IC seat. Posted title ignored.",
    ""
  ];
  matches.forEach(function (m, idx) {
    lines.push(idx + 1 + ". " + m.company + " — " + m.title + " — " + m.percent + "% — " + m.location);
    lines.push(m.url);
    if (m.driveUrl) lines.push("Doc: " + m.driveUrl);
    lines.push("");
  });
  if (errors.length) {
    lines.push("Board errors:");
    errors.forEach(function (e) {
      lines.push("- " + e);
    });
  }
  GmailApp.sendEmail(
    CFG.EMAIL,
    "Canada ATS digest — " + matches.length + " new match" + (matches.length === 1 ? "" : "es"),
    lines.join("\n")
  );
}

function scoreApplicationsFolder() {
  var root = applicationsRoot_();
  var folders = root.getFolders();
  var rows = [["company", "role", "percent", "pass", "reasons"]];
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName() === "ARCHIVE") continue;
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType() !== MimeType.GOOGLE_DOCS) continue;
      var text = DocumentApp.openById(file.getId()).getBody().getText();
      var scored = scoreAgainstTargetJd_(text);
      var ic = isAnalyticsIcJd_(text);
      var pass = scored.percent >= CFG.MATCH_THRESHOLD && ic;
      rows.push([
        folder.getName(),
        file.getName(),
        scored.percent,
        pass ? "yes" : "no",
        (scored.reasons || []).join(" | ") + (ic ? "" : " | not analytics IC seat")
      ]);
    }
  }
  var ss = trackerSheet_();
  var tab = ss.getSheetByName("Folder scores") || ss.insertSheet("Folder scores");
  tab.clear();
  tab.getRange(1, 1, rows.length, 5).setValues(rows);
  SpreadsheetApp.getActive().toast("Wrote " + (rows.length - 1) + " folder scores");
}
