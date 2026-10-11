function importLinkedInJobsFromGmail(optSilent) {
  var digest = optSilent === true;
  var tabs = ensureTrackerTabs_();
  var sheet = tabs.linkedin;
  var known = linkedinKnownRows_(sheet);
  if (!digest) backfillAlertSource_(sheet);
  var feedbackSkips = { companies: {}, roles: {} };
  try {
    feedbackSkips = collectFeedbackSkips_();
  } catch (eSkip) {}
  var days = digest ? CFG.DIGEST_GMAIL_LOOKBACK_DAYS || 2 : CFG.LINKEDIN_LOOKBACK_DAYS || 21;
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
  var pageSize = 40;
  var maxThreads = digest ? CFG.DIGEST_GMAIL_MAX_THREADS || 40 : CFG.LINKEDIN_MAX_THREADS || 120;

  while (start < maxThreads) {
    var threads = GmailApp.search(query, start, pageSize);
    if (!threads.length) break;
    threads.forEach(function (thread) {
      var messages = digest ? [thread.getMessages().pop()] : thread.getMessages();
      messages.forEach(function (msg) {
        if (!msg) return;
        var subject = String(msg.getSubject() || "");
        var html = "";
        try {
          html = String(msg.getBody() || "");
        } catch (eHtml) {}
        var body = digest ? htmlToAlertText_(html) : linkedinMessageText_(msg);
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
            if (!digest) refreshAlertRow_(sheet, known[job.jobId], job, subject);
            return;
          }
          if (seenThisRun[job.jobId]) return;
          if (isSkippedByFeedback_(job.company, job.role, feedbackSkips, job.url)) return;
          seenThisRun[job.jobId] = true;
          var stamp = torontoParts_(msg.getDate() || new Date());
          sheet.appendRow([
            stamp.date,
            stamp.time,
            cleanAlertField_(job.company, "company"),
            cleanAlertField_(job.role, "role"),
            cleanAlertField_(job.location, "location"),
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
      role: cleanAlertField_(parsed.role, "role"),
      company: cleanAlertField_(parsed.company, "company"),
      location: cleanAlertField_(parsed.location, "location"),
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
    var parsed = parseJobrightHtmlCard_(raw.slice(match.index, match.index + 4000));
    if (!parsed.role) parsed = parseJobrightCardText_(htmlToAlertText_(raw.slice(match.index, match.index + 4000)));
    if (isDigestHeadline_(parsed.role, subject)) parsed.role = "";
    if (!parsed.role || !parsed.company) {
      var fromSubject = parseJobrightSubject_(subject);
      if (!parsed.role && fromSubject.role && !isDigestHeadline_(fromSubject.role, subject)) parsed.role = fromSubject.role;
      if (!parsed.company && looksLikeCompanyName_(fromSubject.company)) parsed.company = fromSubject.company;
    }
    jobs[jobId] = {
      jobId: jobId,
      url: "https://jobright.ai/jobs/info/" + jobId,
      role: cleanAlertField_(parsed.role, "role"),
      company: cleanAlertField_(parsed.company, "company"),
      location: cleanAlertField_(parsed.location, "location"),
      easyApply: false,
      source: "Jobright"
    };
  }
  return Object.keys(jobs).map(function (id) {
    return jobs[id];
  });
}

function parseJobrightHtmlCard_(chunk) {
  return parseJobrightCardText_(htmlToAlertText_(chunk));
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
      if (/instant alert|always be the first|early applicant/i.test(line)) return false;
      if (/minutes? ago|hours? ago|days? ago/i.test(line)) return false;
      if (/referrals$/i.test(line)) return false;
      if (/more great matches|your top job matches|explore this today/i.test(line)) return false;
      if (/[{}]|font-family|margin:|padding:|!important|style=/i.test(line)) return false;
      if (/^https?:\/\//i.test(line)) return false;
      if (/^[\d.]+\s*%$/.test(line)) return false;
      if (/·/.test(line)) return false;
      if (/^\$|CA\$|\/yr/i.test(line)) return false;
      if (looksLikeCssOrJd_(line)) return false;
      return true;
    });
  var company = "";
  var role = "";
  var location = "";
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (/^(remote|hybrid)\b/i.test(line) || /,\s*[A-Z]{2}\b/.test(line) || /\b(toronto|montreal|vancouver|calgary|ottawa|canada)\b/i.test(line)) {
      if (!location) location = line;
      continue;
    }
    if (!company && looksLikeCompanyName_(line)) {
      company = line;
      continue;
    }
    if (company && !role) {
      role = line;
      continue;
    }
  }
  return { role: role, company: company, location: location };
}

function parseJobrightSubject_(subject) {
  var posted = String(subject || "").match(/^(.+?)\s+just posted a\s+\d+%\s+match\s+(.+?)\s+role\b/i);
  if (posted) return { company: posted[1].trim(), role: posted[2].trim() };
  var quoted = String(subject || "").match(/[“"](.+?)[”"]/);
  if (quoted) return { company: "", role: quoted[1].trim() };
  return { role: "", company: "" };
}

function htmlToAlertText_(html) {
  return decodeLinkedInEntities_(String(html || "").replace(/\r/g, "\n"))
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|td|h[1-6]|li|a)>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\n[ \t]+/g, "\n");
}

function indeedTitlesFromHtml_(html) {
  var titles = {};
  var re = /jk=([a-f0-9]+)[^>]*>\s*([^<]{3,140})\s*</gi;
  var match;
  while ((match = re.exec(String(html || "")))) {
    var title = decodeLinkedInEntities_(match[2]).replace(/\s+/g, " ").trim();
    if (!title || /^https?:/i.test(title)) continue;
    if (!titles[match[1]]) titles[match[1]] = title;
  }
  return titles;
}

function isDigestHeadline_(text, subject) {
  var t = String(text || "");
  var s = String(subject || "");
  if (/your top job matches|you would be a great fit|and \d+ more/i.test(t)) return true;
  if (!/more .+jobs|you would be a great fit|top job matches/i.test(s)) return false;
  var sub = parseAtCompanySubject_(s);
  return Boolean(t && sub.role && t.toLowerCase() === String(sub.role).toLowerCase());
}

function looksLikeCssOrJd_(text) {
  var t = String(text || "").trim();
  if (!t) return true;
  if (/[{}]|font-family|font-weight|margin:|padding:|rgba?\(|!important|display:|text-decoration|border-radius/i.test(t)) {
    return true;
  }
  if (t.length > 80) return true;
  if (/[.!?]/.test(t) && t.split(/\s+/).length > 8) return true;
  if (
    /^(the|a|an|this|you|we|build|bachelor|undergraduate|analyze|recommend|ownership|high-quality|continuous)\b/i.test(t) &&
    t.split(/\s+/).length > 6
  ) {
    return true;
  }
  return false;
}

function looksLikeCompanyName_(text) {
  var t = String(text || "").trim();
  if (looksLikeCssOrJd_(t)) return false;
  if (t.split(/\s+/).length > 14) return false;
  return t.length >= 2 && t.length <= 100;
}

function cleanAlertField_(value, kind) {
  var t = String(value || "").replace(/\s+/g, " ").trim();
  if (!t) return "";
  if (kind === "company" && !looksLikeCompanyName_(t)) return "";
  if (kind !== "company" && looksLikeCssOrJd_(t)) return "";
  return t;
}

function applySubjectFallback_(parsed, subject, onlyIfEmpty) {
  if (onlyIfEmpty && (parsed.role || parsed.company)) return parsed;
  if (/more .+jobs|you would be a great fit|top job matches/i.test(String(subject || ""))) return parsed;
  if (!parsed.role) {
    var fromSubject = parseAtCompanySubject_(subject);
    parsed.role = fromSubject.role;
    if (!parsed.company) parsed.company = fromSubject.company;
    if (!parsed.location) parsed.location = fromSubject.location || "";
  }
  return parsed;
}

function refreshAlertRow_(sheet, row, job, subject) {
  var values = sheet.getRange(row, 1, 1, 11).getValues()[0];
  if (job.easyApply && String(values[8] || "") !== "yes") {
    sheet.getRange(row, 9).setValue("yes");
  }
  var oldRole = String(values[3] || "").trim();
  var oldCompany = String(values[2] || "").trim();
  var oldLoc = String(values[4] || "").trim();
  var copied = isDigestHeadline_(oldRole, subject) || isDigestHeadline_(oldCompany, subject);
  var role = cleanAlertField_(job.role, "role");
  var company = cleanAlertField_(job.company, "company");
  var loc = cleanAlertField_(job.location, "location");
  if (role && (!oldRole || copied || looksLikeCssOrJd_(oldRole))) sheet.getRange(row, 4).setValue(role);
  if (company && (!oldCompany || copied || looksLikeCssOrJd_(oldCompany))) sheet.getRange(row, 3).setValue(company);
  if (loc && (!oldLoc || copied || looksLikeCssOrJd_(oldLoc))) sheet.getRange(row, 5).setValue(loc);
}

function parseIndeedJobs_(html, subject) {
  var raw = String(html || "");
  var titles = indeedTitlesFromHtml_(raw);
  var text = htmlToAlertText_(raw);
  var jobs = {};
  var re = /https?:\/\/(?:[\w.-]+\.)?indeed\.com\/[^\s"'<>]*[?&]jk=([a-f0-9]+)/gi;
  var match;
  while ((match = re.exec(text))) {
    var jk = match[1];
    var jobId = "indeed_" + jk;
    if (jobs[jobId]) continue;
    var href = match[0].toLowerCase();
    if (href.indexOf("/jobs?") !== -1 && href.indexOf("jk=") === -1) continue;
    var before = text.slice(Math.max(0, match.index - 500), match.index);
    var parsed = parseIndeedCardText_(before);
    if (titles[jk]) parsed.role = titles[jk];
    if (isDigestHeadline_(parsed.role, subject)) parsed.role = titles[jk] || "";
    applySubjectFallback_(parsed, subject, true);
    jobs[jobId] = {
      jobId: jobId,
      url: "https://www.indeed.com/viewjob?jk=" + jk,
      role: cleanAlertField_(parsed.role, "role"),
      company: cleanAlertField_(parsed.company, "company"),
      location: cleanAlertField_(parsed.location, "location"),
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
      if (looksLikeCssOrJd_(line) && !/\s+[-–]\s+/.test(line)) return false;
      return true;
    });
  var role = "";
  var company = "";
  var location = "";
  for (var i = lines.length - 1; i >= 0; i--) {
    var coLoc = lines[i].match(/^(.{2,70})\s+[-–]\s+(.{2,70})$/);
    if (coLoc && looksLikeCompanyName_(coLoc[1])) {
      company = coLoc[1].trim();
      location = coLoc[2].trim();
      if (i >= 1) role = lines[i - 1];
      break;
    }
  }
  if (!company && lines.length === 1) role = lines[0];
  return { role: role, company: company, location: location };
}

function parseGlassdoorJobs_(html, subject) {
  var raw = String(html || "");
  var jobs = {};
  var re = /jobListing\.htm\?[^"'<>]*jobListingId=(\d+)/gi;
  var match;
  var foundHtm = false;
  while ((match = re.exec(raw))) {
    foundHtm = true;
    addGlassdoorJob_(jobs, match[1], raw.slice(match.index, match.index + 6000), subject);
  }
  if (!foundHtm) {
    re = /jobListingId=(\d+)/gi;
    while ((match = re.exec(raw))) {
      addGlassdoorJob_(jobs, match[1], raw.slice(match.index, match.index + 6000), subject);
    }
  }
  return Object.keys(jobs).map(function (id) {
    return jobs[id];
  });
}

function addGlassdoorJob_(jobs, listingId, chunk, subject) {
  var jobId = "gd_" + listingId;
  if (jobs[jobId]) return;
  var parsed = parseGlassdoorHtmlCard_(chunk);
  if (!parsed.role || !parsed.company) {
    var fromText = parseGlassdoorCardText_(htmlToAlertText_(chunk));
    if (!parsed.company) parsed.company = fromText.company;
    if (!parsed.role) parsed.role = fromText.role;
    if (!parsed.location) parsed.location = fromText.location;
  }
  if (isDigestHeadline_(parsed.role, subject)) parsed.role = "";
  applySubjectFallback_(parsed, subject, true);
  jobs[jobId] = {
    jobId: jobId,
    url: "https://www.glassdoor.com/job-listing/job.htm?jl=" + listingId,
    role: cleanAlertField_(parsed.role, "role"),
    company: cleanAlertField_(parsed.company, "company"),
    location: cleanAlertField_(parsed.location, "location"),
    easyApply: /easy apply/i.test(chunk),
    source: "Glassdoor"
  };
}

function parseGlassdoorHtmlCard_(chunk) {
  var html = String(chunk || "");
  var texts = [];
  var re = />([^<]{2,120})</g;
  var m;
  while ((m = re.exec(html))) {
    var t = decodeLinkedInEntities_(m[1]).replace(/\s+/g, " ").trim();
    if (!t) continue;
    if (/★|easy apply|glassdoor est|employer est|apply now|^ca\$|^\$|^\d+[dh]$|^\($|^\)/i.test(t)) continue;
    if (looksLikeCssOrJd_(t) && !looksLikeCompanyName_(t)) continue;
    texts.push(t);
  }
  var company = "";
  var role = "";
  var location = "";
  for (var i = 0; i < texts.length; i++) {
    var line = texts[i];
    if (/^(remote|hybrid)\b/i.test(line) || /,\s*[A-Z]{2}\b/.test(line) || /^(toronto|ottawa|edmonton|calgary|vancouver|montreal|alberta|ontario|canada|hicksville|findlay|parker|birmingham)\b/i.test(line)) {
      if (!location) location = line;
      continue;
    }
    if (!company && looksLikeCompanyName_(line)) {
      company = line.replace(/\s*\|\s*.+$/, function (rest) {
        return rest.length > 40 ? "" : rest;
      }).replace(/\s*\|\s*$/, "").trim() || line.split("|")[0].trim();
      continue;
    }
    if (company && !role && line !== company) {
      role = line;
      continue;
    }
  }
  return { role: role, company: company, location: location };
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
      if (/^apply now/i.test(line)) return false;
      if (/^job alert/i.test(line)) return false;
      if (/your job listings|jobs for /i.test(line)) return false;
      if (/glassdoor est|employer est/i.test(line)) return false;
      if (/^CA\$|^\$|★/.test(line) && line.length < 40) return false;
      if (/\d+\.\d+\s*★/.test(line)) return false;
      if (/^\d+[dh]$|^just posted$/i.test(line)) return false;
      if (/^joblistingid=/i.test(line)) return false;
      if (/utm_|rdforyou|jrtk=/i.test(line)) return false;
      if (line.length > 160) return false;
      return true;
    });
  var role = "";
  var company = "";
  var location = "";
  if (lines.length >= 3) {
    company = lines[0].replace(/\s*\d+(\.\d+)?\s*$/, "").trim();
    role = lines[1];
    location = lines[2];
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
