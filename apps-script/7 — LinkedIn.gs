function importLinkedInJobsFromGmail() {
  var tabs = ensureTrackerTabs_();
  var sheet = tabs.linkedin;
  var known = linkedinKnownIds_(sheet);
  var days = CFG.LINKEDIN_LOOKBACK_DAYS || 21;
  var query =
    "newer_than:" +
    days +
    "d (from:jobalerts-noreply@linkedin.com OR from:jobs-noreply@linkedin.com)";
  var threads = GmailApp.search(query, 0, 50);
  var added = 0;
  var seenThisRun = {};

  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (msg) {
      var body = String(msg.getPlainBody() || "") + "\n" + String(msg.getBody() || "");
      var jobs = parseLinkedInJobs_(body, msg.getSubject() || "");
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
          msg.getSubject() || ""
        ]);
        known[job.jobId] = true;
        added += 1;
      });
    });
  });

  SpreadsheetApp.getActive().toast(
    added ? "Added " + added + " LinkedIn job link" + (added === 1 ? "" : "s") : "No new LinkedIn job links"
  );
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
  var jobs = [];
  var cards = text.split(/-{8,}/);
  cards.forEach(function (card) {
    var urlMatch = card.match(/https?:\/\/(?:www\.)?linkedin\.com\/(?:comm\/)?jobs\/view\/(\d+)/i);
    if (!urlMatch) return;
    if (/jobs\/search-results/i.test(card) && !/view job:/i.test(card)) return;
    var jobId = urlMatch[1];
    var lines = card
      .split("\n")
      .map(function (line) {
        return line.replace(/\u00a0/g, " ").replace(/[͏]/g, "").trim();
      })
      .filter(function (line) {
        return (
          line &&
          !/^view job:/i.test(line) &&
          !/^apply with resume/i.test(line) &&
          !/^top applicant$/i.test(line) &&
          !/^https?:\/\//i.test(line)
        );
      });
    var role = "";
    var company = "";
    var location = "";
    var start = 0;
    if (lines[0] && /your job alert/i.test(lines[0])) start = 1;
    if (lines[start] && /new jobs match/i.test(lines[start])) start += 1;
    if (lines[start]) role = lines[start];
    if (lines[start + 1]) company = lines[start + 1];
    if (lines[start + 2] && !/^see all jobs/i.test(lines[start + 2])) location = lines[start + 2];
    if (!role) {
      var parsed = parseLinkedInSubject_(subject);
      role = parsed.role;
      if (!company) company = parsed.company;
    }
    jobs.push({
      jobId: jobId,
      url: "https://www.linkedin.com/jobs/view/" + jobId,
      role: role,
      company: company,
      location: location
    });
  });

  if (!jobs.length) {
    var fallback = text.match(/https?:\/\/(?:www\.)?linkedin\.com\/(?:comm\/)?jobs\/view\/(\d+)/gi) || [];
    fallback.forEach(function (raw) {
      var jobId = linkedinJobIdFromUrl_(raw);
      if (!jobId) return;
      var parsed = parseLinkedInSubject_(subject);
      jobs.push({
        jobId: jobId,
        url: "https://www.linkedin.com/jobs/view/" + jobId,
        role: parsed.role,
        company: parsed.company,
        location: ""
      });
    });
  }

  var unique = [];
  var seen = {};
  jobs.forEach(function (job) {
    if (seen[job.jobId]) return;
    seen[job.jobId] = true;
    unique.push(job);
  });
  return unique;
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
    .replace(/&gt;/gi, ">");
}
