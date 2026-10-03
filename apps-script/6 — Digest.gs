function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Canada ATS")
    .addItem("Run digest now", "runDigestNow")
    .addItem("Set up Suggestions and Application status tabs", "setupTrackerTabs")
    .addItem("Score Applications folder", "scoreApplicationsFolder")
    .addItem("Refresh Gmail progress", "refreshGmailProgressOnSheet_")
    .addSeparator()
    .addItem("Install 8 AM / 6 PM triggers", "installDigestTriggers")
    .addItem("Remove digest triggers", "removeDigestTriggers")
    .addToUi();
  try {
    ensureTrackerTabs_();
  } catch (e) {}
}

function setupTrackerTabs() {
  ensureTrackerTabs_();
  SpreadsheetApp.getActive().toast("Tabs ready: Suggestions (new recs) and Application status (your apply list)");
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
  ensureTrackerTabs_();
  var runAt = new Date();
  var started = Date.now();
  var props = PropertiesService.getScriptProperties();
  var cursor = Number(props.getProperty("boardCursor") || "0");
  var boards = defaultCatalog_().filter(function (b) {
    return b.enabled;
  });
  var applyDay = recommendedApplyDay_();
  var newMatches = [];
  var errors = [];
  var considered = 0;
  var end = Math.min(boards.length, cursor + CFG.BOARD_BATCH);

  for (var i = cursor; i < end; i++) {
    if (Date.now() - started > CFG.MAX_RUNTIME_MS) {
      end = i;
      break;
    }
    var fetched = fetchBoardJobs_(boards[i]);
    if (fetched.error) errors.push(boards[i].company + ": " + fetched.error);
    (fetched.jobs || []).forEach(function (job) {
      if (isAlreadyAppliedCompany_(job.company)) return;
      if (!isCanadaEligible_(job.location, job.description)) return;
      var scored = scoreAgainstTargetJd_(job.description);
      job.matchPercent = scored.percent;
      job.matchReasons = scored.reasons;
      considered += 1;
      if (scored.percent < CFG.MATCH_THRESHOLD) return;
      if (!isAnalyticsIcJd_(job.description)) return;
      if (docAlreadyExists_(job.company, job.title)) return;
      var driveUrl = createRoleDoc_(job, applyDay);
      var gmail = inferGmailProgress_(job.company, job.title);
      recordSuggestedRole_(job, applyDay, driveUrl, gmail.progress, gmail.signal, runAt, label);
      newMatches.push({
        company: job.company,
        title: job.title,
        percent: job.matchPercent,
        url: job.url,
        driveUrl: driveUrl,
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
  refreshGmailProgressOnSheet_();
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
