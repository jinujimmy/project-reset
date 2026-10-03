function fetchJson_(url) {
  var res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { Accept: "application/json" },
  });
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("HTTP " + code + " " + url);
  }
  return JSON.parse(res.getContentText());
}

function boardUrl_(source, token, offset) {
  var t = encodeURIComponent(token);
  if (source === "greenhouse") {
    return "https://boards-api.greenhouse.io/v1/boards/" + t + "/jobs?content=true";
  }
  if (source === "lever") return "https://api.lever.co/v0/postings/" + t + "?mode=json";
  if (source === "ashby") return "https://api.ashbyhq.com/posting-api/job-board/" + t;
  return (
    "https://api.smartrecruiters.com/v1/companies/" +
    t +
    "/postings?limit=100&offset=" +
    (offset || 0)
  );
}

function asJobsGreenhouse_(board, payload) {
  var jobs = [];
  (payload.jobs || []).forEach(function (row) {
    if (!row || !row.title || !row.absolute_url) return;
    jobs.push({
      title: row.title,
      company: row.company_name || board.company,
      location: (row.location && row.location.name) || "",
      url: row.absolute_url,
      description: htmlToPlain_(row.content || ""),
      source: "greenhouse",
    });
  });
  return jobs;
}

function asJobsLever_(board, payload) {
  if (!payload || !payload.length) return [];
  return payload
    .map(function (row) {
      if (!row || !row.text) return null;
      var cats = row.categories || {};
      var locs = (cats.allLocations || []).concat([cats.location, row.country]).filter(Boolean);
      return {
        title: row.text,
        company: board.company,
        location: locs.join(", "),
        url: row.hostedUrl || row.applyUrl,
        description: row.descriptionPlain || htmlToPlain_(row.description || ""),
        source: "lever",
      };
    })
    .filter(Boolean);
}

function asJobsAshby_(board, payload) {
  return (payload.jobs || [])
    .map(function (row) {
      if (!row || !row.title) return null;
      return {
        title: row.title,
        company: board.company,
        location: row.location || "",
        url: row.jobUrl || row.applyUrl,
        description: row.descriptionPlain || htmlToPlain_(row.descriptionHtml || ""),
        source: "ashby",
      };
    })
    .filter(Boolean);
}

function enrichSmartRecruiters_(board, jobs) {
  jobs
    .filter(function (job) {
      return shouldFetchDescription_(job.title);
    })
    .slice(0, 40)
    .forEach(function (job) {
      var id = String(job.url || "").split("/").pop();
      if (!id) return;
      try {
        var payload = fetchJson_(
          "https://api.smartrecruiters.com/v1/companies/" +
            encodeURIComponent(board.token) +
            "/postings/" +
            encodeURIComponent(id),
        );
        var sections = (((payload.jobAd || {}).sections) || {});
        var chunks = [];
        ["jobDescription", "qualifications", "additionalInformation"].forEach(function (key) {
          var section = sections[key] || {};
          chunks.push(htmlToPlain_(section.text || section.html || ""));
        });
        job.description = chunks.join(" ");
      } catch (e) {
        // keep listing
      }
    });
}

function fetchBoardJobs_(board) {
  try {
    if (board.source === "smartrecruiters") {
      var payload = fetchJson_(boardUrl_(board.source, board.token, 0));
      var jobs = (payload.content || []).map(function (row) {
        var loc = row.location || {};
        return {
          title: row.name,
          company: (row.company && row.company.name) || board.company,
          location: loc.fullLocation || [loc.city, loc.region, loc.country].filter(Boolean).join(", "),
          url: row.ref || "",
          description: "",
          source: "smartrecruiters",
        };
      });
      enrichSmartRecruiters_(board, jobs);
      return { jobs: jobs };
    }
    var data = fetchJson_(boardUrl_(board.source, board.token));
    var parsed =
      board.source === "greenhouse"
        ? asJobsGreenhouse_(board, data)
        : board.source === "lever"
          ? asJobsLever_(board, data)
          : asJobsAshby_(board, data);
    return { jobs: parsed };
  } catch (error) {
    return { jobs: [], error: String(error) };
  }
}
