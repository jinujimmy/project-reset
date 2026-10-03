function htmlToPlain_(input) {
  if (!input) return "";
  return String(input)
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isAlreadyAppliedCompany_(company) {
  var c = String(company || "").toLowerCase();
  return APPLIED_COMPANY_SKIP.some(function (name) {
    return c.indexOf(name) !== -1;
  });
}

var CANADA_PLACE_ =
  /\b(canada|canadian|toronto|vancouver|montreal|montréal|calgary|ottawa|edmonton|winnipeg|halifax|mississauga|brampton|kitchener|waterloo|london,\s*on|ontario|quebec|québec|british columbia|\bbc\b|alberta|manitoba|saskatchewan|nova scotia|new brunswick)\b/i;

function mentionsCanada_(text) {
  return CANADA_PLACE_.test(String(text || ""));
}

function placeAfterRemote_(location) {
  var loc = String(location || "").replace(/[–—]/g, "-");
  if (!/\bremote\b/i.test(loc)) return "";
  return loc
    .replace(/\bremote\b/gi, " ")
    .replace(
      /\b(or|and|hybrid|onsite|on-site|flexible|worldwide|global|anywhere|north america|americas|multiple locations|various|tbd|wfh|work from home)\b/gi,
      " "
    )
    .replace(/[^a-z\u00C0-\u024F\s]/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isRemoteOtherCountry_(location) {
  if (!/\bremote\b/i.test(String(location || ""))) return false;
  if (mentionsCanada_(location)) return false;
  return placeAfterRemote_(location).length > 0;
}

function basedInOtherCountry_(description) {
  var match = String(description || "").match(/\bbased in\s+([^.\n]{2,80})/i);
  if (!match) return false;
  if (mentionsCanada_(match[1])) return false;
  var rest = match[1]
    .replace(/\b(the|a|an|our|your|location|office|offices|home|region|area)\b/gi, " ")
    .replace(/[^a-z\u00C0-\u024F\s]/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
  return rest.length > 0;
}

function isCanadaEligible_(location, description) {
  var loc = String(location || "");
  var blob = loc + " " + String(description || "");
  if (isRemoteOtherCountry_(loc)) return false;
  if (/\bremote\b/i.test(loc) && !mentionsCanada_(loc) && basedInOtherCountry_(description)) return false;
  if (mentionsCanada_(blob)) return true;
  if (/\bremote\b/i.test(loc)) return true;
  return false;
}

function bucketHit_(text, signals) {
  return signals.some(function (signal) {
    return text.indexOf(signal) !== -1;
  });
}

function scoreAgainstTargetJd_(description) {
  var text = String(description || "").toLowerCase();
  var reasons = [];
  var asked = SKILL_BUCKETS.filter(function (bucket) {
    if (!bucketHit_(text, bucket.signals)) return false;
    if (bucket.name === "SQL" && /\b(no|not|without)\s+sql\b(?!\s+server)/.test(text)) {
      return false;
    }
    return true;
  });
  var covered = asked.filter(function (bucket) {
    return TARGET_COVERED[bucket.name];
  });
  var askedNames = asked.map(function (b) {
    return b.name;
  });
  var coverage = 0;
  if (asked.length === 0) coverage = 0;
  else if (asked.length < CFG.MIN_ASKED_SKILLS) {
    coverage = Math.round((covered.length / CFG.MIN_ASKED_SKILLS) * 100);
  } else {
    coverage = Math.round((covered.length / asked.length) * 100);
  }

  var askedSql = asked.some(function (b) {
    return b.name === "SQL";
  });
  var disqualifiers = DISQUALIFYING_SIGNALS.filter(function (s) {
    return text.indexOf(s) !== -1;
  });
  var softHits = SOFT_NEGATIVE_SIGNALS.filter(function (s) {
    return text.indexOf(s) !== -1;
  });
  var uatWithoutSql = !askedSql && softHits.indexOf("user acceptance testing") !== -1;
  var penalty = 12 * (disqualifiers.length + softHits.length);
  var percent = Math.max(0, Math.min(100, coverage - penalty));
  if (disqualifiers.length > 0 || uatWithoutSql) percent = Math.min(percent, 49);
  if (asked.length === 0) reasons.push("asked 0 skills");
  else reasons.push("coverage " + covered.length + "/" + asked.length + " asked (" + askedNames.join(", ") + ")");
  return { percent: percent, reasons: reasons };
}

function isAnalyticsIcJd_(description) {
  var text = String(description || "").toLowerCase();
  var family =
    /\b(data analysts?|product analysts?|business data analysts?|insights analysts?|insight analysts?|reporting analysts?|bi analysts?|business intelligence analysts?|analytics engineers?|analytics interns?|growth analysts?|marketing analysts?|staff analysts?)\b/i;
  if (!family.test(text)) return false;
  if (/\bhire, manage, and develop\b/.test(text)) return false;
  if (/\bowning their performance, growth\b/.test(text)) return false;
  if (
    /\b(this role is|looking for)[\s\S]{0,80}\b(director|vice president|head of analytics|head of data)\b/.test(
      text
    )
  ) {
    return false;
  }
  var hireNear =
    /\b(?:looking for|we(?:'re| are) (?:hiring|seeking)|this role|join us as)[\s\S]{0,200}\b(?:data analysts?|product analysts?|business data analysts?|insights analysts?|insight analysts?|reporting analysts?|bi analysts?|business intelligence analysts?|analytics engineers?|analytics interns?|growth analysts?|marketing analysts?|staff analysts?)\b/i;
  if (hireNear.test(text)) return true;
  var seatRole =
    /\b(?:data analysts?|product analysts?|business data analysts?|insights analysts?|insight analysts?|reporting analysts?|bi analysts?|business intelligence analysts?|analytics engineers?|analytics interns?|growth analysts?|marketing analysts?|staff analysts?)[\s\S]{0,60}(?:roles?|positions?|intern(?:ship)?s?)\b/i;
  if (seatRole.test(text)) return true;
  return (
    /\bas an intern\b/.test(text) &&
    /\benrolled in\b/.test(text) &&
    /\bdata analysts?\b/.test(text) &&
    /\b(sql|dashboard|experiment)\b/.test(text) &&
    !/\bphd\b/.test(text)
  );
}

function shouldFetchDescription_(title) {
  var tooSenior =
    /\b(principal|director|vice president|\bvp\b|chief|head of|distinguished|fellow|general manager|\bgm\b|c-level|svp)\b/i;
  var skip =
    /\b(software engineer|engineer|architect|swe\b|frontend|front-end|backend|back-end|full[- ]?stack|account executive|sales development|\bsdr\b|product manager|program manager|data scientist|data science)\b/i;
  var hints =
    /\b(analyst|analytics|insight|business|product|data|strategy|operations|research|intelligence|reporting|metric|growth|bizops|decision science|program)\b/i;
  if (tooSenior.test(title)) return false;
  if (skip.test(title) && !/\b(analyst|analytics|insights?)\b/i.test(title)) return false;
  return hints.test(title);
}

function recommendedApplyDay_() {
  var tz = "America/Toronto";
  var now = new Date();
  var hour = Number(Utilities.formatDate(now, tz, "H"));
  var base = new Date(now.getTime());
  if (hour >= 12) base.setDate(base.getDate() + 1);
  while ([0, 6].indexOf(base.getDay()) !== -1) base.setDate(base.getDate() + 1);
  return Utilities.formatDate(base, tz, "yyyy-MM-dd");
}
