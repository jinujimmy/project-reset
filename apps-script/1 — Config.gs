/** Bind this project to the Canada ATS tracker. Timezone: America/Toronto. */
var CFG = {
  APPLICATIONS_FOLDER_ID: "YOUR_APPLICATIONS_FOLDER_ID",
  TRACKER_SHEET_ID: "YOUR_TRACKER_SHEET_ID",
  EMAIL: "you@example.com",
  MATCH_THRESHOLD: 70,
  MIN_ASKED_SKILLS: 3,
  BOARD_BATCH: 12,
  MAX_RUNTIME_MS: 270000,
  SUGGESTIONS_TAB: "Suggestions",
  STATUS_TAB: "Application status",
  LINKEDIN_TAB: "Suggestions - Linkedin",
  LINKEDIN_LOOKBACK_DAYS: 21
};

var SKILL_BUCKETS = [
  { name: "SQL", signals: ["sql"] },
  { name: "Python", signals: ["python"] },
  {
    name: "Dashboards / BI",
    signals: [
      "dashboard",
      "looker",
      "tableau",
      "power bi",
      "quicksight",
      "thoughtspot",
      "sigma",
      "metabase",
      "preset"
    ]
  },
  { name: "Metrics / reporting", signals: ["kpi", "metric", "reporting", "scorecard", "insights"] },
  {
    name: "Stakeholders",
    signals: [
      "stakeholder",
      "product manager",
      "cross-functional",
      "cross-functionally",
      "senior leadership",
      "business partner"
    ]
  },
  {
    name: "Warehouse",
    signals: ["warehouse", "snowflake", "dbt", "bigquery", "big query", "redshift", "databricks"]
  },
  {
    name: "Experiments / funnels",
    signals: [
      "experiment",
      "a/b test",
      "ab test",
      "a/b",
      "funnel",
      "retention",
      "conversion",
      "customer journey",
      "user journey"
    ]
  }
];

var TARGET_COVERED = {
  SQL: true,
  Python: true,
  "Dashboards / BI": true,
  "Metrics / reporting": true,
  Stakeholders: true,
  Warehouse: true,
  "Experiments / funnels": true
};

var DISQUALIFYING_SIGNALS = [
  "credit specialist",
  "collections and recoveries",
  "plant analyst",
  "manufacturing execution",
  "manufacturing floor",
  "people manager",
  "direct reports",
  "head of data",
  "lead and manage a team",
  "in a leadership position",
  "merchandising"
];

var SOFT_NEGATIVE_SIGNALS = [
  "user acceptance testing",
  "user stories and acceptance criteria"
];

var APPLIED_COMPANY_SKIP = [
  "caseware",
  "mckesson",
  "eldis",
  "high liner",
  "highliner",
  "affirm",
  "roofr",
  "clutch",
  "opentable",
  "open table",
  "ebay",
  "tjx",
  "pacific smoke",
  "suzuki",
  "optimum",
  "manulife",
  "hellofresh",
  "hello fresh",
  "kroll",
  "konrad",
  "vitasoy",
  "niche baker",
  "capital one",
  "capone",
  "moneris",
  "childrens aid",
  "children's aid",
  "fat guys",
  "dare foods",
  "form3",
  "cognizant",
  "babylist",
  "upstart",
  "spinwheel",
  "nearsource",
  "isg",
  "covet",
  "alphasense",
  "stackadapt"
];
