function importLinkedInJobsFromGmail(optSilent) {
  var tabs = ensureTrackerTabs_();
  var sheet = tabs.linkedin;
  var known = linkedinKnownRows_(sheet);
  backfillAlertSource_(sheet);
  var feedbackSkips = { companies: {}, roles: {} };
  try {
    feedbackSkips = collectFeedbackSkips_();
  } catch (eSkip) {}
  var days = CFG.LINKEDIN_LOOKBACK_DAYS || 21;
  var label = CFG.JOB_ALERT_LABEL || "Job Alerts";
  var query =
    "newer_than:" +
    days +
    "d (label:\"" +
    label +
    "\" OR from:noreply@jobright.ai OR from:donotreply@jobalert.indeed.com OR from:donotreply@match.indeed.com OR from:noreply@glassdoor.com OR from:jobalerts-noreply@linkedin.com OR from:jobs-noreply@linkedin.com)";
  var added = 0;
  var seenThisRun = {};
  var start = 0;
  var pageSize = 50;
  var maxThreads = 300;

  while (start < maxThreads) {
    var threads = GmailApp.search(query, start, pageSize);
    if (!threads.length) break;
    threads.forEach(function (thread) {
      thread.getMessages().forEach(function (msg) {
        var subject = String(msg.getSubject() || "");
        var html = "";
        try {
          html = String(msg.getBody() || "");
        } catch (eHtml) {}
        var body = linkedinMessageText_(msg);
        var blob = (html || "") + "\n" + (body || "");
        var jobs = [];
        try {
          jobs = jobs.concat(parseLinkedInJobs_(body, subject));
        } catch (eLi) {}
        try {
          jobs = jobs.concat(parseJobrightJobs_(blob, subject));
        } catch (eJr) {}
        try {
          jobs = jobs.concat(parseIndeedJobs_(blob, subject));
        } catch (eIn) {}
        try {
          jobs = jobs.concat(parseGlassdoorJobs_(blob, subject));
        } catch (eGd) {}
        jobs.forEach(function (job) {
          if (!job.jobId) return;
          if (known[job.jobId]) {
            if (job.easyApply && String(sheet.getRange(known[job.jobId], 9).getValue() || "") !== "yes") {
              sheet.getRange(known[job.jobId], 9).setValue("yes");
            }
            return;
          }
          if (seenThisRun[job.jobId]) return;
          if (isSkippedByFeedback_(job.company, job.role, feedbackSkips, job.url)) return;
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
            subject,
            job.easyApply ? "yes" : "",
            job.source || alertSourceFromUrl_(job.url),
            ""
          ]);
          known[job.jobId] = sheet.getLastRow();
          added += 1;
        });
      });
    });
    if (threads.length < pageSize) break;
    start += pageSize;
  }

  if (optSilent !== true) {
    SpreadsheetApp.getActive().toast(
      added ? "Added " + added + " job link" + (added === 1 ? "" : "s") : "No new job links"
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

function alertSourceFromUrl_(url) {
  var u = String(url || "").toLowerCase();
  if (u.indexOf("linkedin.com") !== -1) return "LinkedIn";
  if (u.indexOf("jobright.ai") !== -1) return "Jobright";
  if (u.indexOf("indeed.com") !== -1) return "Indeed";
  if (u.indexOf("glassdoor.com") !== -1) return "Glassdoor";
  return "";
}

function backfillAlertSource_(sheet) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][9] || "").trim()) continue;
    var src = alertSourceFromUrl_(values[i][5]);
    if (src) sheet.getRange(i + 1, 10).setValue(src);
  }
}

function linkedinKnownRows_(sheet) {
  var known = {};
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var id = String(values[i][6] || "").trim();
    var url = String(values[i][5] || "");
    if (!id) id = alertJobIdFromUrl_(url);
    if (id) known[id] = i + 1;
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
      location: parsed.location,
      easyApply: /apply with resume|easy apply/i.test(before),
      source: "LinkedIn"
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

function parseJobrightJobs_(html, subject) {
  var raw = decodeLinkedInEntities_(String(html || "").replace(/\r/g, "\n"));
  var jobs = {};
  var re = /https?:\/\/(?:www\.)?jobright\.ai\/jobs\/info\/([a-f0-9]+)/gi;
  var match;
  while ((match = re.exec(raw))) {
    var jobId = match[1];
    if (jobs[jobId]) continue;
    var before = raw
      .slice(Math.max(0, match.index - 2500), match.index)
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div|tr|h[1-6])>/gi, "\n")
      .replace(/<[^>]+>/g, " ");
    var parsed = parseJobrightCardText_(before);
    if (!parsed.role || !parsed.company) {
      var fromSubject = parseJobrightSubject_(subject);
      if (!parsed.role) parsed.role = fromSubject.role;
      if (!parsed.company) parsed.company = fromSubject.company;
    }
    jobs[jobId] = {
      jobId: jobId,
      url: "https://jobright.ai/jobs/info/" + jobId,
      role: parsed.role,
      company: parsed.company,
      location: parsed.location,
      easyApply: false,
      source: "Jobright"
    };
  }
  return Object.keys(jobs).map(function (id) {
    return jobs[id];
  });
}

function parseJobrightCardText_(chunk) {
  var lines = String(chunk || "")
    .split("\n")
    .map(function (line) {
      return line.replace(/\u00a0/g, " ").replace(/\u034F/g, "").replace(/\s+/g, " ").trim();
    })
    .filter(function (line) {
      if (!line) return false;
      if (/^apply now/i.test(line)) return false;
      if (/^jobright/i.test(line)) return false;
      if (/instant alert/i.test(line)) return false;
      if (/always be the first/i.test(line)) return false;
      if (/early applicant/i.test(line)) return false;
      if (/minutes? ago|hours? ago|days? ago/i.test(line)) return false;
      if (/referrals$/i.test(line)) return false;
      if (/more great matches/i.test(line)) return false;
      if (/^https?:\/\//i.test(line)) return false;
      if (/^[\d.]+\s*%$/.test(line)) return false;
      if (/·/.test(line) && line.length < 60) return false;
      return true;
    });
  var role = "";
  var company = "";
  var location = "";
  if (lines.length >= 3) {
    company = lines[lines.length - 3];
    role = lines[lines.length - 2];
    location = lines[lines.length - 1];
    if (/remote|toronto|canada|vancouver|montreal|calgary|ottawa|,\s*[A-Z]{2}\b/i.test(role) && !/remote|toronto|canada/i.test(location)) {
      location = role;
      role = company;
      company = lines.length >= 4 ? lines[lines.length - 4] : "";
    }
  } else if (lines.length === 2) {
    company = lines[0];
    role = lines[1];
  } else if (lines.length === 1) {
    role = lines[0];
  }
  if (/^\$|CA\$|\/yr/i.test(location)) location = "";
  return { role: role, company: company, location: location };
}

function parseJobrightSubject_(subject) {
  var posted = String(subject || "").match(/^(.+?)\s+just posted a\s+\d+%\s+match\s+(.+?)\s+role\b/i);
  if (posted) return { company: posted[1].trim(), role: posted[2].trim() };
  var quoted = String(subject || "").match(/[“"](.+?)[”"]/);
  if (quoted) return { company: "", role: quoted[1].trim() };
  return { role: "", company: "" };
}

function parseIndeedJobs_(html, subject) {
  var text = decodeLinkedInEntities_(String(html || "").replace(/\r/g, "\n"));
  var jobs = {};
  var re = /https?:\/\/(?:[\w.-]+\.)?indeed\.com\/[^\s"'<>]*[?&]jk=([a-f0-9]+)/gi;
  var match;
  while ((match = re.exec(text))) {
    var jk = match[1];
    var jobId = "indeed_" + jk;
    if (jobs[jobId]) continue;
    var raw = match[0].toLowerCase();
    if (raw.indexOf("/jobs?") !== -1 && raw.indexOf("jk=") === -1) continue;
    var before = text.slice(Math.max(0, match.index - 500), match.index);
    var parsed = parseIndeedCardText_(before);
    if (!parsed.role) {
      var fromSubject = parseAtCompanySubject_(subject);
      parsed.role = fromSubject.role;
      if (!parsed.company) parsed.company = fromSubject.company;
    }
    jobs[jobId] = {
      jobId: jobId,
      url: "https://www.indeed.com/viewjob?jk=" + jk,
      role: parsed.role,
      company: parsed.company,
      location: parsed.location,
      easyApply: /easily apply/i.test(before),
      source: "Indeed"
    };
  }
  return Object.keys(jobs).map(function (id) {
    return jobs[id];
  });
}

function parseIndeedCardText_(chunk) {
  var lines = String(chunk || "")
    .split("\n")
    .map(function (line) {
      return line.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
    })
    .filter(function (line) {
      if (!line) return false;
      if (/^https?:\/\//i.test(line)) return false;
      if (/^easily apply$/i.test(line)) return false;
      if (/just posted|days? ago|responsive employer/i.test(line)) return false;
      if (/^\$|CA\$|₹|a year|an hour/i.test(line)) return false;
      if (/indeed job alert/i.test(line)) return false;
      if (/see matching results/i.test(line)) return false;
      if (line.length > 220) return false;
      return true;
    });
  var role = "";
  var company = "";
  var location = "";
  if (lines.length >= 2) {
    role = lines[lines.length - 2];
    var coLoc = lines[lines.length - 1].match(/^(.*?)\s+[-–]\s+(.+)$/);
    if (coLoc) {
      company = coLoc[1].trim();
      location = coLoc[2].trim();
    } else {
      company = lines[lines.length - 1];
    }
  } else if (lines.length === 1) {
    role = lines[0];
  }
  return { role: role, company: company, location: location };
}

function parseGlassdoorJobs_(html, subject) {
  var text = decodeLinkedInEntities_(String(html || "").replace(/\r/g, "\n"));
  var jobs = {};
  var re = /https?:\/\/(?:www\.)?glassdoor\.com\/partner\/jobListing\.htm\?[^\s"'<>]*jobListingId=(\d+)/gi;
  var match;
  while ((match = re.exec(text))) {
    var listingId = match[1];
    var jobId = "gd_" + listingId;
    if (jobs[jobId]) continue;
    var before = text.slice(Math.max(0, match.index - 600), match.index);
    var parsed = parseGlassdoorCardText_(before);
    if (!parsed.role) {
      var fromSubject = parseAtCompanySubject_(subject);
      parsed.role = fromSubject.role;
      if (!parsed.company) parsed.company = fromSubject.company;
    }
    jobs[jobId] = {
      jobId: jobId,
      url: "https://www.glassdoor.com/job-listing/job.htm?jl=" + listingId,
      role: parsed.role,
      company: parsed.company,
      location: parsed.location,
      easyApply: /easy apply/i.test(before),
      source: "Glassdoor"
    };
  }
  return Object.keys(jobs).map(function (id) {
    return jobs[id];
  });
}

function parseGlassdoorCardText_(chunk) {
  var lines = String(chunk || "")
    .split("\n")
    .map(function (line) {
      return line.replace(/\u00a0/g, " ").replace(/\ufeff/g, "").replace(/\|/g, " ").replace(/\s+/g, " ").trim();
    })
    .filter(function (line) {
      if (!line) return false;
      if (/^https?:\/\//i.test(line)) return false;
      if (/^easy apply$/i.test(line)) return false;
      if (/^job alert/i.test(line)) return false;
      if (/your job listings/i.test(line)) return false;
      if (/glassdoor est|employer est/i.test(line)) return false;
      if (/^CA\$|^\$|★/.test(line) && line.length < 40) return false;
      if (/^\d+[dh]$|^just posted$/i.test(line)) return false;
      if (line.length > 160) return false;
      return true;
    });
  var role = "";
  var company = "";
  var location = "";
  if (lines.length >= 3) {
    company = lines[lines.length - 3].replace(/\s*\d+(\.\d+)?\s*$/, "").trim();
    role = lines[lines.length - 2];
    location = lines[lines.length - 1];
  } else if (lines.length === 2) {
    company = lines[0];
    role = lines[1];
  } else if (lines.length === 1) {
    role = lines[0];
  }
  return { role: role, company: company, location: location };
}

function parseAtCompanySubject_(subject) {
  var text = String(subject || "").replace(/\s+and \d+ more.*$/i, "").trim();
  var at = text.match(/^(.*?)\s+at\s+(.+?)(?:\s+in\s+(.+))?$/i);
  if (at) return { role: at[1].trim(), company: at[2].replace(/\s*@\s*/g, "").trim(), location: (at[3] || "").trim() };
  return { role: text, company: "" };
}

function linkedinJobIdFromUrl_(url) {
  var match = String(url || "").match(/\/jobs\/view\/(\d+)/i);
  return match ? match[1] : "";
}

function alertJobIdFromUrl_(url) {
  var u = String(url || "");
  var linkedin = linkedinJobIdFromUrl_(u);
  if (linkedin) return linkedin;
  var jobright = u.match(/jobright\.ai\/jobs\/info\/([a-f0-9]+)/i);
  if (jobright) return jobright[1];
  var indeed = u.match(/[?&]jk=([a-f0-9]+)/i);
  if (indeed && /indeed\.com/i.test(u)) return "indeed_" + indeed[1];
  var gd = u.match(/jobListingId=(\d+)/i) || u.match(/[?&]jl=(\d+)/i);
  if (gd && /glassdoor\.com/i.test(u)) return "gd_" + gd[1];
  return "";
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
