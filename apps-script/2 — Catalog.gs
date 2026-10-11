function defaultCatalog_() {
  return [
    ["greenhouse", "stripe", "Stripe"],
    ["greenhouse", "databricks", "Databricks"],
    ["greenhouse", "airbnb", "Airbnb"],
    ["greenhouse", "cloudflare", "Cloudflare"],
    ["greenhouse", "gitlab", "GitLab"],
    ["greenhouse", "figma", "Figma"],
    ["greenhouse", "janestreet", "Jane Street"],
    ["greenhouse", "point72", "Point72"],
    ["ashby", "openai", "OpenAI"],
    ["ashby", "snowflake", "Snowflake"],
    ["ashby", "notion", "Notion"],
    ["ashby", "ramp", "Ramp"],
    ["ashby", "cursor", "Cursor"],
    ["ashby", "perplexity", "Perplexity"],
    ["ashby", "linear", "Linear"],
    ["ashby", "posthog", "PostHog"],
    ["ashby", "motherduck", "MotherDuck"],
    ["ashby", "hex", "Hex"],
    ["lever", "palantir", "Palantir"],
    ["lever", "wealthfront", "Wealthfront"],
    ["smartrecruiters", "NielsenIQ", "NielsenIQ"],
    ["lever", "caseware", "Caseware"],
    ["ashby", "wealthsimple", "Wealthsimple"],
    ["greenhouse", "instacart", "Instacart"],
    ["ashby", "1password", "1Password"],
    ["greenhouse", "hootsuite", "Hootsuite"],
    ["ashby", "hopper", "Hopper"],
    ["ashby", "lightspeed", "Lightspeed"],
    ["greenhouse", "d2l", "D2L"],
    ["greenhouse", "stackadapt", "StackAdapt"],
    ["greenhouse", "ritual", "Ritual"],
    ["ashby", "thinkific", "Thinkific"],
    ["ashby", "koho", "KOHO"],
    ["greenhouse", "datadog", "Datadog"],
    ["greenhouse", "hubspot", "HubSpot"],
    ["greenhouse", "mongodb", "MongoDB"],
    ["greenhouse", "intercom", "Intercom"],
    ["greenhouse", "asana", "Asana"],
    ["greenhouse", "dropbox", "Dropbox"],
    ["greenhouse", "coinbase", "Coinbase"],
    ["greenhouse", "twilio", "Twilio"],
  ].map(function (row) {
    return { source: row[0], token: row[1], company: row[2], enabled: true };
  });
}

function catalogHeaders_() {
  return ["source", "token", "company", "enabled"];
}

function ensureCatalogTab_(ss) {
  var name = CFG.CATALOG_TAB || "Catalog";
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  ensureHeaderRow_(sheet, catalogHeaders_());
  if (sheet.getLastRow() < 2) {
    var seed = defaultCatalog_().map(function (b) {
      return [b.source, b.token, b.company, b.enabled === false ? "no" : "yes"];
    });
    if (seed.length) sheet.getRange(2, 1, seed.length, 4).setValues(seed);
  }
  ensureCatalogDropdowns_(sheet);
  return sheet;
}

function ensureCatalogDropdowns_(sheet) {
  var last = Math.max(sheet.getLastRow(), 20);
  var sources = SpreadsheetApp.newDataValidation()
    .requireValueInList(["greenhouse", "lever", "ashby", "smartrecruiters"], true)
    .setAllowInvalid(false)
    .setHelpText("ATS board type. Not LinkedIn or Jobright.")
    .build();
  var enabled = SpreadsheetApp.newDataValidation()
    .requireValueInList(["yes", "no"], true)
    .setAllowInvalid(true)
    .setHelpText("yes = scan this company on the digest. no = skip.")
    .build();
  sheet.getRange(2, 1, last, 1).setDataValidation(sources);
  sheet.getRange(2, 4, last, 1).setDataValidation(enabled);
}

function activeCatalog_() {
  var fallback = defaultCatalog_().filter(function (b) {
    return b.enabled;
  });
  try {
    var sheet = ensureTrackerTabs_().catalog;
    if (!sheet) return fallback;
    var values = sheet.getDataRange().getValues();
    var boards = [];
    var seen = {};
    var sawRow = false;
    for (var i = 1; i < values.length; i++) {
      var source = String(values[i][0] || "").toLowerCase().trim();
      var token = String(values[i][1] || "").trim();
      var company = String(values[i][2] || token).trim();
      var flag = String(values[i][3] || "yes").toLowerCase().trim();
      if (!token && !source && !String(values[i][2] || "").trim()) continue;
      sawRow = true;
      if (!token) continue;
      if (
        source !== "greenhouse" &&
        source !== "lever" &&
        source !== "ashby" &&
        source !== "smartrecruiters"
      ) {
        continue;
      }
      if (flag === "no" || flag === "false" || flag === "0") continue;
      var key = source + "\t" + token.toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;
      boards.push({ source: source, token: token, company: company, enabled: true });
    }
    if (sawRow) return boards;
  } catch (e) {}
  return fallback;
}
