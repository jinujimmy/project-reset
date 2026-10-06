function importLinkedInJobsFromGmail(optSilent) {
  var tabs = ensureTrackerTabs_();
  var sheet = tabs.linkedin;
  var known = linkedinKnownIds_(sheet);
  var days = CFG.LINKEDIN_LOOKBACK_DAYS || 21;
  var query =
    "newer_than:" +
    days +
    "d (from:jobalerts-noreply@linkedin.com OR from:jobs-noreply@linkedin.com)";
  var added = 0;
  var seenThisRun = {};
  var start = 0;
  var pageSize = 50;
  var maxThreads = 250;

  while (start < maxThreads) {
    var threads = GmailApp.search(query, start, pageSize);
    if (!threads.length) break;
    threads.forEach(function (thread) {
      thread.getMessages().forEach(function (msg) {
        var subject = String(msg.getSubject() || "");
        var body = linkedinMessageText_(msg);
        var jobs = parseLinkedInJobs_(body, subject);
        jobs.forEach(function (job) {
          if (!job.jobId || known[job.jobId] || seenThisRun[job.jobId]) return;
          seenThisRun[job.jobId] = true;
          var stamp = torontoParts_(msg.getDate() || new Date());
          sheet.appendRow([
            stamp.date,
            stamp.time,
            job.company,
            job.role,
            job.location,
            job.url,
            job.jobId,
            subject
          ]);
          known[job.jobId] = true;
          added += 1;
        });
      });
    });
    if (threads.length < pageSize) break;
    start += pageSize;
  }

  if (optSilent !== true) {
    SpreadsheetApp.getActive().toast(
      added ? "Added " + added + " LinkedIn job link" + (added === 1 ? "" : "s") : "No new LinkedIn job links"
    );
  }
  return added;
}

function linkedinMessageText_(msg) {
  var plain = "";
  var html = "";
  try {
    plain = String(msg.getPlainBody() || "");
  } catch (e) {}
  try {
    html = String(msg.getBody() || "");
  } catch (e2) {}
  html = html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<[^>]+>/g, " ");
  return decodeLinkedInEntities_(plain + "\n" + html);
}

function linkedinKnownIds_(sheet) {
  var known = {};
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var id = String(values[i][6] || "").trim();
    var url = String(values[i][5] || "");
    if (!id) id = linkedinJobIdFromUrl_(url);
    if (id) known[id] = true;
  }
  return known;
}

function parseLinkedInJobs_(body, subject) {
  var text = decodeLinkedInEntities_(String(body || "").replace(/\r/g, "\n"));
  var jobs = {};
  var re = /https?:\/\/(?:[\w.-]+\.)?linkedin\.com\/(?:comm\/)?jobs\/view\/(\d+)/gi;
  var match;
  while ((match = re.exec(text))) {
    var jobId = match[1];
    if (jobs[jobId]) continue;
    var around = text.slice(Math.max(0, match.index - 80), match.index + 40).toLowerCase();
    if (around.indexOf("search-results") !== -1 && around.indexOf("view job") === -1) continue;
    var before = text.slice(Math.max(0, match.index - 700), match.index);
    var parsed = parseLinkedInCardText_(before);
    if (!parsed.role) {
      var fromSubject = parseLinkedInSubject_(subject);
      parsed.role = fromSubject.role;
      if (!parsed.company) parsed.company = fromSubject.company;
    }
    jobs[jobId] = {
      jobId: jobId,
      url: "https://www.linkedin.com/jobs/view/" + jobId,
      role: parsed.role,
      company: parsed.company,
      location: parsed.location
    };
  }
  return Object.keys(jobs).map(function (id) {
    return jobs[id];
  });
}

function parseLinkedInCardText_(chunk) {
  var lines = String(chunk || "")
    .split("\n")
    .map(function (line) {
      return line.replace(/\u00a0/g, " ").replace(/\u034F/g, "").replace(/\s+/g, " ").trim();
    })
    .filter(function (line) {
      if (!line) return false;
      if (/^view job:/i.test(line)) return false;
      if (/^apply with resume/i.test(line)) return false;
      if (/^top applicant$/i.test(line)) return false;
      if (/^this company is actively hiring$/i.test(line)) return false;
      if (/school alumni/i.test(line)) return false;
      if (/^https?:\/\//i.test(line)) return false;
      if (/^see all jobs/i.test(line)) return false;
      if (/^view all jobs/i.test(line)) return false;
      if (/your job alert/i.test(line)) return false;
      if (/new jobs match/i.test(line)) return false;
      if (/expand your search/i.test(line)) return false;
      if (/recommendations based on/i.test(line)) return false;
      if (/^this email was intended/i.test(line)) return false;
      if (/ jobs$/i.test(line) && line.split(" ").length <= 4) return false;
      return true;
    });
  var role = "";
  var company = "";
  var location = "";
  if (lines.length >= 3) {
    role = lines[lines.length - 3];
    company = lines[lines.length - 2];
    location = lines[lines.length - 1];
  } else if (lines.length === 2) {
    role = lines[0];
    company = lines[1];
  } else if (lines.length === 1) {
    role = lines[0];
  }
  return { role: role, company: company, location: location };
}

function parseLinkedInSubject_(subject) {
  var text = String(subject || "").trim();
  var at = text.match(/^(.*?)\s+at\s+(.+)$/i);
  if (at) return { role: at[1].trim(), company: at[2].trim() };
  return { role: text, company: "" };
}

function linkedinJobIdFromUrl_(url) {
  var match = String(url || "").match(/\/jobs\/view\/(\d+)/i);
  return match ? match[1] : "";
}

function decodeLinkedInEntities_(text) {
  return String(text || "")
    .replace(/&amp;/gi, "&")
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#(\d+);/g, function (_, n) {
      return String.fromCharCode(Number(n));
    });
}
