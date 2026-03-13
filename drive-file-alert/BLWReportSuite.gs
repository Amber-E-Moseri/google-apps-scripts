/***************************************
 * BLW COMBINED REPORT SUITE v4.0
 *
 * Changes from v3.2:
 *   1. Elvanto Import source removed throughout
 *      (Cell FTs + Service FTs only)
 *   2. New top-level menu: 📦 Comprehensive Reports
 *      → "✨ Generate All Reports" button
 *      → Prompts label, grade, % colour
 *      → Validates FT sheets (graceful fallback)
 *      → Generates all 6 combined subgroup sheets in one run
 *
 * ── MENUS ──────────────────────────────────────────────────────
 *   📦 Comprehensive Reports
 *     ✨ Generate All Reports   ← all 6 combined sheets at once
 *     📧 Email Reports as PDF
 *
 *   📊 Follow-Up Reports
 *     1) Regional Overview — No Grade
 *     2) Regional Overview — With Grade
 *     3) Subgroup Report — No Grade       ← combined sheet (follow-up + FT)
 *     4) Subgroup Report — With Grade     ← combined sheet (follow-up + FT)
 *     📝 Leader Note Counter
 *     🗑 Delete Report Sheets
 *
 *   📋 First-Timer Tools
 *     1) FT Regional Report               ← FT-only, one sheet
 *     2) FT Subgroup Reports              ← combined sheets (follow-up + FT)
 *     3) FT Regional + All Subgroups
 *     ⚙️ Setup FT Import Sheets
 *     🔍 Diagnose FT Data
 ***************************************/


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FOLLOW-UP CONFIG ─────────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

const CFG = {
  NOTES_SHEET: "ELVANTO_NOTES_CLEAN",
  GROUP_SHEET: "GROUP",

  PROTECTED_SHEETS: [
    "How to Use",
    "ELVANTO_NOTES_RAW",
    "ELVANTO_NOTES_CLEAN",
    "GROUP",
    "Cell FTs Import",
    "Service FTs Import",
  ],

  SUBGROUPS: [
    "BLW Canada Central-East SGA",
    "BLW Canada Central-East SGB",
    "BLW Canada Central SGA",
    "BLW Canada Central SGB",
    "BLW Canada West SGA",
    "BLW Canada West SGB",
  ],

  CATEGORY_MAP: {
    "BLW Canada Central-East SGA": "BLW Canada Central-East SGA",
    "BLW Canada Central-East SGB": "BLW Canada Central-East SGB",
    "BLW Canada Central SGA":      "BLW Canada Central SGA",
    "BLW Canada Central SGB":      "BLW Canada Central SGB",
    "BLW Canada West SGA":         "BLW Canada West SGA",
    "BLW Canada West SGB":         "BLW Canada West SGB",
  },

  SUBGROUP_CROSSWALK: {
    "BLW Canada Central-East SGA": "Central East Subgroup A",
    "BLW Canada Central-East SGB": "Central East Subgroup B",
    "BLW Canada Central SGA":      "Central Subgroup A",
    "BLW Canada Central SGB":      "Central Subgroup B",
    "BLW Canada West SGA":         "West Subgroup A",
    "BLW Canada West SGB":         "West Subgroup B",
  },

  TIMEZONE: "America/Toronto",
};

// ─── FOLLOW-UP THEME ──────────────────────────────────────────────────────────

const T = {
  titleBg:    "#0D2540",
  titleFg:    "#FFFFFF",
  accentBg:   "#1F3A5F",
  accentFg:   "#FFFFFF",
  periodBg:   "#1F3A5F",
  periodFg:   "#A8C4E0",
  headerBg:   "#2A5298",
  headerFg:   "#FFFFFF",
  rowBase:    "#FFFFFF",
  rowAlt:     "#F0F5FC",
  totalBg:    "#1F3A5F",
  totalFg:    "#FFFFFF",
  sectionBg:  "#0D2540",
  sectionFg:  "#FFFFFF",
  dividerBg:  "#D0E0F4",
  statCardBg: "#EBF2FB",
  statCardFg: "#0D2540",
  statValFg:  "#1F3A5F",
  mutedFg:    "#6B7A8A",
  font:       "Google Sans",
  fontSize:   10,
};


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER CONFIG ───────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

const FT_IMPORT = {
  CELL:    "Cell FTs Import",
  SERVICE: "Service FTs Import",
};

const FT_COL_FULL_NAME    = "full name";
const FT_COL_GROUPS       = "groups";
const FT_COL_DEMOGRAPHICS = "demographics";
const FT_COL_DATE_ADDED   = "date added";

const FT_KNOWN_SUBGROUPS = [
  "Central East Subgroup A",
  "Central East Subgroup B",
  "Central Subgroup A",
  "Central Subgroup B",
  "West Subgroup A",
  "West Subgroup B",
];

// ─── FIRST-TIMER THEME ────────────────────────────────────────────────────────

const FT = {
  titleBg:         "#1F3A5F",
  titleFg:         "#FFFFFF",
  subtitleFg:      "#A8C4E0",
  rowBase:         "#FFFFFF",
  rowAlt:          "#F8FAFD",
  nameFg:          "#10243F",
  accentFg:        "#1F3A5F",
  mutedFg:         "#6B7A8A",
  border:          "#B0C4DE",

  cellBannerBg:     "#1F3A5F",
  serviceBannerBg:  "#2E5FA3",
  combinedBannerBg: "#0D2540",
  bannerFg:         "#FFFFFF",

  groupHdrBg:      "#2A5298",
  groupHdrFg:      "#FFFFFF",
  subRowBg:        "#EBF2FB",
  subRowAltBg:     "#FFFFFF",
  subNameFg:       "#10243F",
  subCountFg:      "#1F3A5F",

  subtotalBg:      "#CCDDF5",
  subtotalFg:      "#0D2540",
  grandTotalBg:    "#0D2540",
  grandTotalFg:    "#FFFFFF",
  grandCountFg:    "#A8C8F0",

  colHdrBg:        "#E4EEF9",
  colHdrFg:        "#0D2540",

  font:            "Google Sans",
  fontSize:        10,
};

const COMBINED_COLS  = 6;
const FT_DETAIL_COLS = 3;
const FT_TALLY_COLS  = 2;


// ═══════════════════════════════════════════════════════════════════════════════
// ─── UNIFIED MENU ─────────────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("📦 Comprehensive Reports")
    .addItem("✨  Generate All Reports", "btn_comprehensiveAll")
    .addSeparator()
    .addItem("📧  Email Reports as PDF",  "btn_emailReports")
    .addToUi();

  ui.createMenu("📊 Follow-Up Reports")
    .addItem("1)  Regional Overview — No Grade",   "btn_regionalNoGrade")
    .addItem("2)  Regional Overview — With Grade", "btn_regionalWithGrade")
    .addSeparator()
    .addItem("3)  Subgroup Report — No Grade",     "btn_subgroupNoGrade")
    .addItem("4)  Subgroup Report — With Grade",   "btn_subgroupWithGrade")
    .addSeparator()
    .addItem("📝  Leader Note Counter",            "btn_leaderNoteCounter")
    .addSeparator()
    .addItem("🗑  Delete Report Sheets",           "btn_deleteSheets")
    .addToUi();

  ui.createMenu("📋 First-Timer Tools")
    .addItem("1)  FT Regional Report",          "btn_ftRegional")
    .addItem("2)  FT Subgroup Reports",         "btn_ftSubgroup")
    .addItem("3)  FT Regional + All Subgroups", "btn_ftBoth")
    .addSeparator()
    .addItem("⚙️  Setup FT Import Sheets",      "btn_setupFTSheets")
    .addSeparator()
    .addItem("🔍  Diagnose FT Data",            "btn_ftDiagnose")
    .addToUi();
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FOLLOW-UP: HEALTH BANDS ──────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function healthBand_(score) {
  if (score >= 75) return { bg: "#D9F2E3", fg: "#1B5E3C" };
  if (score >= 50) return { bg: "#FFF4CC", fg: "#7A5A00" };
  if (score >= 25) return { bg: "#FFE3CC", fg: "#8A3C00" };
  return               { bg: "#F8D7DA", fg: "#7A1C24" };
}

function applyHealthColour_(sh, r, col, score) {
  const { bg, fg } = healthBand_(score);
  sh.getRange(r, col).setBackground(bg).setFontColor(fg).setFontWeight("bold");
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── COMPREHENSIVE REPORTS: BUTTON + RUNNER ───────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function btn_comprehensiveAll() { runComprehensiveReports_(); }

function runComprehensiveReports_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const labelResp = ui.prompt(
    "Report Label",
    "Enter a label for this report\n(e.g. Feb 2026 Pulse Check):",
    ui.ButtonSet.OK_CANCEL
  );
  if (labelResp.getSelectedButton() !== ui.Button.OK) return;
  const label = (labelResp.getResponseText() || "").trim() || "Comprehensive Report";

  const gradeResp = ui.alert(
    "Include Grade?",
    "Include a grade (A–F) in the follow-up section of each report?",
    ui.ButtonSet.YES_NO
  );
  if (gradeResp === ui.Button.CANCEL) return;
  const grade = (gradeResp === ui.Button.YES);

  let colourPct = false;
  if (grade) {
    const pctResp = ui.alert(
      "Colour Code % Reached?",
      "Apply health-band colours to % Reached cells as well?\n\nGrade cells are always colour coded.",
      ui.ButtonSet.YES_NO
    );
    colourPct = (pctResp === ui.Button.YES);
  }

  const notesSh = ss.getSheetByName(CFG.NOTES_SHEET);
  const groupSh = ss.getSheetByName(CFG.GROUP_SHEET);
  if (!notesSh || !groupSh) {
    ui.alert("❌  Missing sheet: " + (!notesSh ? CFG.NOTES_SHEET : CFG.GROUP_SHEET));
    return;
  }

  const ftSheetErrors = [];
  Object.values(FT_IMPORT).forEach(name => {
    if (!ss.getSheetByName(name)) {
      ftSheetErrors.push(`"${name}" — sheet not found`);
      return;
    }
    const v = validateFTSheetHeaders_(name);
    if (!v.ok) ftSheetErrors.push(`"${name}" — missing columns: ${v.missing.join(", ")}`);
  });

  let ftAllData = null;
  if (ftSheetErrors.length > 0) {
    const proceed = ui.alert(
      "First-Timer Sheet Issues",
      "Some First-Timer import sheets are missing or have incorrect headers:\n\n• " +
      ftSheetErrors.join("\n• ") +
      "\n\nContinue generating reports with Follow-Up data only?",
      ui.ButtonSet.YES_NO
    );
    if (proceed !== ui.Button.YES) return;
  } else {
    ftAllData = loadAllFTData_();
  }

  const groupDir  = loadGroupDirectory_(groupSh);
  const notesData = aggregateNotes_(notesSh, groupDir.byName);

  const created = [];
  for (const sub of CFG.SUBGROUPS) {
    const sheetName = `${label} — ${sub}`;
    const sh = getOrCreate_(ss, sheetName);
    renderCombinedSubgroupSheet_(sh, {
      label, subgroup: sub, grade, colourPct, groupDir, notesData, ftAllData,
    });
    created.push(sheetName);
  }

  ss.setActiveSheet(ss.getSheetByName(created[0]));
  ui.alert(
    `✅  All Comprehensive Reports generated!\n\n` +
    `${created.length} sheets created:\n` +
    created.map(n => `• ${n}`).join("\n")
  );
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FOLLOW-UP: BUTTON HANDLERS ───────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function btn_regionalNoGrade()   { runReport_({ grade: false, perSubgroup: false }); }
function btn_regionalWithGrade() { runReport_({ grade: true,  perSubgroup: false }); }
function btn_subgroupNoGrade()   { runReport_({ grade: false, perSubgroup: true  }); }
function btn_subgroupWithGrade() { runReport_({ grade: true,  perSubgroup: true  }); }
function btn_leaderNoteCounter() { runLeaderNoteCounter_(); }


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FOLLOW-UP: MAIN RUNNER ───────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function runReport_({ grade, perSubgroup }) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const labelResp = ui.prompt(
    "Report Label",
    "Enter a label for this report\n(e.g. Feb 2026 Pulse Check):",
    ui.ButtonSet.OK_CANCEL
  );
  if (labelResp.getSelectedButton() !== ui.Button.OK) return;
  const label = (labelResp.getResponseText() || "").trim() || "Follow-Up Report";

  let colourPct = false;
  if (grade) {
    const pctResp = ui.alert(
      "Colour Code % Reached?",
      "Apply health-band colours to % Reached cells as well?\n\nGrade cells are always colour coded.",
      ui.ButtonSet.YES_NO
    );
    colourPct = (pctResp === ui.Button.YES);
  }

  const notesSh = ss.getSheetByName(CFG.NOTES_SHEET);
  const groupSh = ss.getSheetByName(CFG.GROUP_SHEET);
  if (!notesSh || !groupSh) {
    ui.alert("❌  Missing sheet: " + (!notesSh ? CFG.NOTES_SHEET : CFG.GROUP_SHEET));
    return;
  }

  const groupDir  = loadGroupDirectory_(groupSh);
  const notesData = aggregateNotes_(notesSh, groupDir.byName);

  if (!perSubgroup) {
    const sh = getOrCreate_(ss, label);
    renderRegionalReport_(sh, { label, grade, colourPct, groupDir, notesData });
    ss.setActiveSheet(sh);
    ui.alert(`✅  Report created: "${label}"`);
    return;
  }

  const ftSheetErrors = [];
  Object.values(FT_IMPORT).forEach(name => {
    if (!ss.getSheetByName(name)) {
      ftSheetErrors.push(`"${name}" — sheet not found`);
      return;
    }
    const v = validateFTSheetHeaders_(name);
    if (!v.ok) ftSheetErrors.push(`"${name}" — missing columns: ${v.missing.join(", ")}`);
  });

  if (ftSheetErrors.length > 0) {
    const proceed = ui.alert(
      "First-Timer Sheet Issues",
      "Some First-Timer import sheets are missing or have incorrect headers:\n\n• " +
      ftSheetErrors.join("\n• ") +
      "\n\nContinue generating subgroup sheets with Follow-Up data only?",
      ui.ButtonSet.YES_NO
    );
    if (proceed !== ui.Button.YES) return;
  }

  const ftAllData = ftSheetErrors.length ? null : loadAllFTData_();

  const scopeResp = ui.prompt(
    "Subgroup Scope",
    "Type a number or name for ONE subgroup, or leave blank for ALL 6:\n\n" +
    CFG.SUBGROUPS.map((s, i) => `${i + 1}. ${s}`).join("\n"),
    ui.ButtonSet.OK_CANCEL
  );
  if (scopeResp.getSelectedButton() !== ui.Button.OK) return;
  const typed = (scopeResp.getResponseText() || "").trim();

  if (typed) {
    const match = resolveSubgroup_(typed);
    if (!match) {
      ui.alert(`"${typed}" didn't match any subgroup. Check spelling or use the number.`);
      return;
    }
    const sheetName = `${label} — ${match}`;
    const sh = getOrCreate_(ss, sheetName);
    renderCombinedSubgroupSheet_(sh, { label, subgroup: match, grade, colourPct, groupDir, notesData, ftAllData });
    ss.setActiveSheet(sh);
    ui.alert(`✅  Report created: "${sheetName}"`);
  } else {
    const created = [];
    for (const sub of CFG.SUBGROUPS) {
      const sheetName = `${label} — ${sub}`;
      const sh = getOrCreate_(ss, sheetName);
      renderCombinedSubgroupSheet_(sh, { label, subgroup: sub, grade, colourPct, groupDir, notesData, ftAllData });
      created.push(sheetName);
    }
    ss.setActiveSheet(ss.getSheetByName(created[0]));
    ui.alert(`✅  Generated ${created.length} reports:\n` + created.map(n => `• ${n}`).join("\n"));
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FOLLOW-UP: LOAD GROUP DIRECTORY ──────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function loadGroupDirectory_(sheet) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { byName: new Map(), bySubgroup: new Map() };

  const hdrs = rows[0].map(h => h.toString().trim().toLowerCase());
  const ci = {
    name:    Math.max(hdrs.indexOf("name"), 0),
    cat:     Math.max(hdrs.indexOf("categories"), 1),
    members: Math.max(hdrs.indexOf("total members"), 2),
  };

  const byName     = new Map();
  const bySubgroup = new Map();

  for (let i = 1; i < rows.length; i++) {
    const r   = rows[i];
    const nm  = str_(r[ci.name]);
    const cat = str_(r[ci.cat]);
    const mem = toNum_(r[ci.members]);
    if (!nm) continue;

    const subgroup        = CFG.CATEGORY_MAP[cat] || null;
    const adjustedMembers = Math.max(0, mem - 1);

    byName.set(nm, { subgroup, members: mem, adjustedMembers });

    if (subgroup) {
      if (!bySubgroup.has(subgroup)) {
        bySubgroup.set(subgroup, { groups: [], totalMembers: 0, adjustedMembers: 0 });
      }
      const sb = bySubgroup.get(subgroup);
      sb.groups.push(nm);
      sb.totalMembers    += mem;
      sb.adjustedMembers += adjustedMembers;
    }
  }

  return { byName, bySubgroup };
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FOLLOW-UP: AGGREGATE NOTES ───────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function aggregateNotes_(sheet, groupByName) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return emptyAgg_();

  const hdrs = rows[0].map(h => h.toString().trim().toLowerCase());
  const ci = {
    date:   findCol_(hdrs, ["notedate", "note date", "date"], 0),
    leader: findCol_(hdrs, ["leader"], 1),
    person: findCol_(hdrs, ["person"], 2),
    group:  findCol_(hdrs, ["groupname", "group name", "group"], 3),
  };

  const bySubgroup = new Map();
  const byGroup    = new Map();
  const byLeader   = new Map();
  const unmapped   = new Map();
  let minDate = null, maxDate = null;

  for (let i = 1; i < rows.length; i++) {
    const r      = rows[i];
    const person = str_(r[ci.person]);
    const group  = str_(r[ci.group]);
    const leader = str_(r[ci.leader]);
    const d      = coerceDate_(r[ci.date]);

    if (!person || !group) continue;
    if (d) {
      if (!minDate || d < minDate) minDate = d;
      if (!maxDate || d > maxDate) maxDate = d;
    }

    let resolvedGroup = group;
    let groupInfo     = groupByName.get(group);

    if (!groupInfo && group.includes(",")) {
      const parts = group.split(",").map(p => p.trim()).filter(Boolean);
      resolvedGroup = parts[parts.length - 1];
      groupInfo     = groupByName.get(resolvedGroup);
      if (!groupInfo) {
        resolvedGroup = parts[0];
        groupInfo     = groupByName.get(resolvedGroup);
      }
    }

    if (!byGroup.has(resolvedGroup)) {
      byGroup.set(resolvedGroup, { totalNotes: 0, people: new Set(), leaders: new Set() });
    }
    const grp = byGroup.get(resolvedGroup);
    grp.totalNotes++;
    grp.people.add(person);
    if (leader) grp.leaders.add(leader);

    if (!groupInfo || !groupInfo.subgroup) {
      unmapped.set(resolvedGroup, (unmapped.get(resolvedGroup) || 0) + 1);
      continue;
    }

    const sub = groupInfo.subgroup;
    if (!bySubgroup.has(sub)) {
      bySubgroup.set(sub, { totalNotes: 0, people: new Set(), groupsActive: new Set(), leaders: new Set() });
    }
    const sb = bySubgroup.get(sub);
    sb.totalNotes++;
    sb.people.add(person);
    sb.groupsActive.add(resolvedGroup);
    if (leader) sb.leaders.add(leader);

    if (leader) {
      if (!byLeader.has(sub)) byLeader.set(sub, new Map());
      const subLeaderMap = byLeader.get(sub);
      if (!subLeaderMap.has(leader)) subLeaderMap.set(leader, { notes: 0, groups: new Set() });
      const ld = subLeaderMap.get(leader);
      ld.notes++;
      ld.groups.add(resolvedGroup);
    }
  }

  const tz  = CFG.TIMEZONE;
  const fmt = d => Utilities.formatDate(d, tz, "MMMM d, yyyy");
  const rangeText = (minDate && maxDate)
    ? (minDate.toDateString() === maxDate.toDateString()
        ? fmt(minDate)
        : `${fmt(minDate)} – ${fmt(maxDate)}`)
    : "(date range unavailable)";

  const allLeaders = new Set();
  bySubgroup.forEach(sb => sb.leaders.forEach(l => allLeaders.add(l)));

  return { rangeText, bySubgroup, byGroup, byLeader, unmapped, allLeaderCount: allLeaders.size };
}

function emptyAgg_() {
  return {
    rangeText: "(no data)",
    bySubgroup: new Map(),
    byGroup: new Map(),
    byLeader: new Map(),
    unmapped: new Map(),
    allLeaderCount: 0,
  };
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FOLLOW-UP: RENDER REGIONAL REPORT ────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function renderRegionalReport_(sh, { label, grade, colourPct, groupDir, notesData }) {
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearFormats();
  sh.getBandings().forEach(b => b.remove());

  const { rangeText, bySubgroup, byGroup, unmapped, allLeaderCount } = notesData;
  const { bySubgroup: grpBySub, byName } = groupDir;

  const NUM_COLS = grade ? 7 : 5;
  let r = 1;
  paintSheet_(sh, 600, NUM_COLS);

  r = titleBlock_(sh, r, label, rangeText,
    grade ? "Regional Overview — With Grade" : "Regional Overview — No Grade", NUM_COLS);
  r = sectionLabel_(sh, r, "SUBGROUP SUMMARY", NUM_COLS);

  const subCols = grade
    ? ["Subgroup", "Total Notes", "Unique People", "Total Members", "Groups Active", "% Reached", "Grade"]
    : ["Subgroup", "Total Notes", "Unique People", "Total Members", "Groups Active"];

  writeHeader_(sh, r, subCols, NUM_COLS);
  r++;

  let gNotes = 0, gUnique = 0, gMembers = 0, gAdjMembers = 0, gActive = 0, gTotal = 0;

  CFG.SUBGROUPS.forEach((sub, i) => {
    const st = bySubgroup.get(sub) || { totalNotes: 0, people: new Set(), groupsActive: new Set() };
    const gd = grpBySub.get(sub)   || { groups: [], totalMembers: 0, adjustedMembers: 0 };

    const notes     = st.totalNotes;
    const unique    = st.people.size;
    const members   = gd.totalMembers;
    const adjMem    = gd.adjustedMembers;
    const active    = st.groupsActive.size;
    const totalGrps = gd.groups.length;

    gNotes      += notes;
    gUnique     += unique;
    gMembers    += members;
    gAdjMembers += adjMem;
    gActive     += active;
    gTotal      += totalGrps;

    const bg = i % 2 === 0 ? T.rowBase : T.rowAlt;
    if (grade) {
      const pct      = safeRatio_(unique, adjMem);
      const pctScore = pct * 100;
      const score    = gradeScore_(pct, active, totalGrps);
      writeDataRow_(sh, r, [sub, notes, unique, members, `${active} / ${totalGrps}`, pctText_(pct), gradeLabel_(score)], bg, NUM_COLS);
      applyHealthColour_(sh, r, 7, score);
      if (colourPct) applyHealthColour_(sh, r, 6, pctScore);
    } else {
      writeDataRow_(sh, r, [sub, notes, unique, members, `${active} / ${totalGrps}`], bg, NUM_COLS);
    }
    r++;
  });

  const totalPct   = safeRatio_(gUnique, gAdjMembers);
  const totalScore = gradeScore_(totalPct, gActive, gTotal);
  if (grade) {
    writeTotalRow_(sh, r, ["TOTAL", gNotes, gUnique, gMembers, `${gActive} / ${gTotal}`, pctText_(totalPct), gradeLabel_(totalScore)], NUM_COLS);
    applyHealthColour_(sh, r, 7, totalScore);
    if (colourPct) applyHealthColour_(sh, r, 6, totalPct * 100);
  } else {
    writeTotalRow_(sh, r, ["TOTAL", gNotes, gUnique, gMembers, `${gActive} / ${gTotal}`], NUM_COLS);
  }
  r += 2;

  if (grade) {
    r = sectionLabel_(sh, r, "CALCULATION NOTE", NUM_COLS);
    const note =
      "Unique People reached ÷ Adjusted Members\n" +
      "(Adjusted Members = Total Members − 1 per group, to account for the group leader)";
    mergeWrite_(sh, r, 1, 1, NUM_COLS, note, { bg: T.rowAlt, fg: "#000000", size: 9, align: "left", height: 36 });
    r += 2;
  }

  r = sectionLabel_(sh, r, "FELLOWSHIP BREAKDOWN", NUM_COLS);

  const fellowCols = grade
    ? ["Group", "Subgroup", "Total Notes", "Unique People", "Total Members", "% Reached", "Grade"]
    : ["Group", "Subgroup", "Total Notes", "Unique People", "Total Members"];

  writeHeader_(sh, r, fellowCols, NUM_COLS);
  r++;

  const fellowRows = [];
  byGroup.forEach((v, g) => {
    const info   = byName.get(g) || {};
    const sub    = info.subgroup || "—";
    const mem    = info.members  || 0;
    const adjMem = info.adjustedMembers || 0;
    const notes  = v.totalNotes;
    const unique = v.people.size;
    const pct    = safeRatio_(unique, adjMem);
    if (grade) {
      const score = pct * 100;
      fellowRows.push({ row: [g, sub, notes, unique, mem, pctText_(pct), gradeLabel_(score)], score });
    } else {
      fellowRows.push({ row: [g, sub, notes, unique, mem], score: 0 });
    }
  });

  fellowRows.sort((a, b) => a.row[1].localeCompare(b.row[1]) || a.row[0].localeCompare(b.row[0]));
  fellowRows.forEach(({ row, score }, i) => {
    writeDataRow_(sh, r, row, i % 2 === 0 ? T.rowBase : T.rowAlt, NUM_COLS);
    if (grade) {
      applyHealthColour_(sh, r, 7, score);
      if (colourPct) applyHealthColour_(sh, r, 6, score);
    }
    r++;
  });
  r++;

  footerRow_(sh, r, `Leaders Participating: ${allLeaderCount}`, NUM_COLS);
  r += 2;

  if (unmapped.size > 0) {
    r = sectionLabel_(sh, r, "⚠  UNMAPPED GROUPS — Not Found in GROUP Sheet", NUM_COLS);
    writeHeader_(sh, r, ["Group Name", "Note Count"], NUM_COLS);
    r++;
    unmapped.forEach((cnt, g) => {
      writeDataRow_(sh, r, [g, cnt], T.rowAlt, NUM_COLS);
      r++;
    });
  }

  finalise_(sh, NUM_COLS);
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── COMBINED SUBGROUP SHEET ───────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function renderCombinedSubgroupSheet_(sh, { label, subgroup, grade, colourPct, groupDir, notesData, ftAllData }) {
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearFormats();
  sh.getBandings().forEach(b => b.remove());

  const ftSubgroupName = CFG.SUBGROUP_CROSSWALK[subgroup] || null;
  const { rangeText, bySubgroup, byGroup, byLeader } = notesData;
  const { bySubgroup: grpBySub, byName }   = groupDir;
  const NC = COMBINED_COLS;
  let r = 1;

  sh.getRange(1, 1, 1000, NC)
    .setBackground(T.rowBase).setFontFamily(T.font).setFontSize(T.fontSize)
    .setFontColor("#000000").setFontWeight("normal");

  sh.getRange(r, 1, 1, NC).merge()
    .setValue(label.toUpperCase())
    .setBackground(T.titleBg).setFontColor(T.titleFg)
    .setFontWeight("bold").setFontSize(18).setFontFamily(T.font)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(r, 52);
  r++;

  sh.getRange(r, 1, 1, NC).merge()
    .setValue(subgroup)
    .setBackground(T.titleBg).setFontColor("#A8C4E0")
    .setFontSize(11).setFontFamily(T.font)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(r, 26);
  r++;

  sh.getRange(r, 1, 1, NC).merge()
    .setValue("📅  " + rangeText)
    .setBackground(T.periodBg).setFontColor(T.periodFg)
    .setFontSize(9).setFontFamily(T.font)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(r, 22);
  r++;

  sh.getRange(r, 1, 1, NC).setBackground(T.rowBase);
  sh.setRowHeight(r, 8);
  r++;

  const st = bySubgroup.get(subgroup) || { totalNotes: 0, people: new Set(), groupsActive: new Set(), leaders: new Set() };
  const gd = grpBySub.get(subgroup)   || { groups: [], totalMembers: 0, adjustedMembers: 0 };

  const notes          = st.totalNotes;
  const unique         = st.people.size;
  const members        = gd.totalMembers;
  const adjMem         = gd.adjustedMembers;
  const active         = st.groupsActive.size;
  const total          = gd.groups.length;
  const pct            = safeRatio_(unique, adjMem);
  const subLeaderCount = st.leaders ? st.leaders.size : 0;
  const gradeScore     = gradeScore_(pct, active, total);

  sh.getRange(r, 1, 1, NC).merge()
    .setValue("FOLLOW-UP SUMMARY")
    .setBackground(T.sectionBg).setFontColor(T.sectionFg)
    .setFontWeight("bold").setFontSize(10).setFontFamily(T.font)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.getRange(r, 1).setValue("  FOLLOW-UP SUMMARY");
  sh.setRowHeight(r, 28);
  r++;

  const stats1 = [
    { label: "TOTAL NOTES",   value: notes },
    { label: "UNIQUE PEOPLE", value: unique },
    { label: "TOTAL MEMBERS", value: members },
  ];
  const stats2 = [
    { label: "% REACHED",     value: pctText_(pct) },
    { label: "GROUPS ACTIVE", value: `${active} / ${total}` },
    { label: "LEADERS",       value: subLeaderCount },
  ];

  function writeStatRow_(stats, rowNum) {
    stats.forEach((stat, i) => {
      const col = i * 2 + 1;
      sh.getRange(rowNum, col, 1, 2).merge()
        .setValue(stat.label)
        .setBackground(T.headerBg).setFontColor(T.headerFg)
        .setFontFamily(T.font).setFontSize(8).setFontWeight("bold")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
      sh.setRowHeight(rowNum, 18);
    });
    const valRow = rowNum + 1;
    stats.forEach((stat, i) => {
      const col = i * 2 + 1;
      sh.getRange(valRow, col, 1, 2).merge()
        .setValue(stat.value)
        .setBackground(T.statCardBg).setFontColor(T.statValFg)
        .setFontFamily(T.font).setFontSize(15).setFontWeight("bold")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
      sh.setRowHeight(valRow, 36);
      if (grade && colourPct && stat.label === "% REACHED") {
        applyHealthColour_(sh, valRow, col, pct * 100);
        sh.getRange(valRow, col, 1, 2).merge();
      }
    });
  }

  writeStatRow_(stats1, r);
  r += 2;

  sh.getRange(r, 1, 1, NC).setBackground(T.dividerBg);
  sh.setRowHeight(r, 2);
  r++;

  writeStatRow_(stats2, r);
  r += 2;

  if (grade) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue("GRADE")
      .setBackground(T.headerBg).setFontColor(T.headerFg)
      .setFontFamily(T.font).setFontSize(8).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sh.setRowHeight(r, 18);
    r++;
    const gradeRange = sh.getRange(r, 1, 1, NC).merge()
      .setValue(gradeLabel_(gradeScore))
      .setFontFamily(T.font).setFontSize(20).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    applyHealthColour_(sh, r, 1, gradeScore);
    gradeRange.merge();
    sh.setRowHeight(r, 40);
    r++;
  }

  sh.getRange(r, 1, 1, NC).setBackground(T.rowBase);
  sh.setRowHeight(r, 10);
  r++;

  sh.getRange(r, 1, 1, NC).merge()
    .setValue("  FELLOWSHIP DETAIL")
    .setBackground(T.sectionBg).setFontColor(T.sectionFg)
    .setFontWeight("bold").setFontSize(10).setFontFamily(T.font)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.setRowHeight(r, 28);
  r++;

  const fellowCols = grade
    ? ["Group", "Notes", "People", "Members", "% Reached", "Grade"]
    : ["Group", "Notes", "People", "Members", "", ""];
  sh.getRange(r, 1, 1, NC)
    .setValues([fellowCols])
    .setBackground(T.headerBg).setFontColor(T.headerFg)
    .setFontWeight("bold").setFontFamily(T.font).setFontSize(T.fontSize)
    .setHorizontalAlignment("center");
  sh.getRange(r, 1).setHorizontalAlignment("left");
  sh.setRowHeight(r, 24);
  r++;

  const allowedGroups = new Set(gd.groups);
  const fellowRows = [];
  byGroup.forEach((v, g) => {
    if (!allowedGroups.has(g)) return;
    const info  = byName.get(g) || {};
    const mem   = info.members || 0;
    const adjM  = info.adjustedMembers || 0;
    const n     = v.totalNotes;
    const u     = v.people.size;
    const p     = safeRatio_(u, adjM);
    const sc    = p * 100;
    fellowRows.push({ row: grade ? [g, n, u, mem, pctText_(p), gradeLabel_(sc)] : [g, n, u, mem, "", ""], score: sc });
  });

  fellowRows.sort((a, b) => a.row[0].localeCompare(b.row[0]));
  if (fellowRows.length) {
    fellowRows.forEach(({ row, score }, i) => {
      const bg = i % 2 === 0 ? T.rowBase : T.rowAlt;
      sh.getRange(r, 1, 1, NC)
        .setValues([row])
        .setBackground(bg).setFontFamily(T.font).setFontSize(T.fontSize)
        .setFontColor("#000000").setFontWeight("normal");
      sh.getRange(r, 1).setFontWeight("bold").setFontColor(T.statValFg);
      sh.getRange(r, 2, 1, NC - 1).setHorizontalAlignment("center");
      sh.getRange(r, 1).setHorizontalAlignment("left");
      if (grade) {
        applyHealthColour_(sh, r, 6, score);
        if (colourPct) applyHealthColour_(sh, r, 5, score);
      }
      sh.setRowHeight(r, 22);
      r++;
    });

    const totalNotes  = fellowRows.reduce((s, x) => s + x.row[1], 0);
    const totalPeople = fellowRows.reduce((s, x) => s + x.row[2], 0);
    const totalMem    = fellowRows.reduce((s, x) => s + x.row[3], 0);
    sh.getRange(r, 1, 1, NC)
      .setValues([["TOTAL", totalNotes, totalPeople, totalMem, grade ? pctText_(pct) : "", grade ? gradeLabel_(gradeScore) : ""]])
      .setBackground(T.totalBg).setFontColor(T.totalFg)
      .setFontWeight("bold").setFontFamily(T.font).setFontSize(T.fontSize);
    sh.getRange(r, 2, 1, NC - 1).setHorizontalAlignment("center");
    sh.getRange(r, 1).setHorizontalAlignment("left");
    if (grade) applyHealthColour_(sh, r, 6, gradeScore);
    sh.setRowHeight(r, 24);
    r++;
  } else {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue("(No fellowship data for this subgroup)")
      .setBackground(T.rowAlt).setFontColor(T.mutedFg)
      .setFontStyle("italic").setHorizontalAlignment("center");
    sh.setRowHeight(r, 22);
    r++;
  }

  sh.getRange(r, 1, 1, NC).setBackground(T.rowBase);
  sh.setRowHeight(r, 10);
  r++;

  const lpHdrRange = sh.getRange(r, 1, 1, NC);
  lpHdrRange.merge();
  lpHdrRange.setValue("  LEADERS PARTICIPATING")
    .setBackground(T.sectionBg).setFontColor(T.sectionFg)
    .setFontWeight("bold").setFontSize(10).setFontFamily(T.font)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.setRowHeight(r, 28);
  r++;

  const subLeaderMap = byLeader.get(subgroup) || new Map();
  if (subLeaderMap.size === 0) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue("(No leader data for this subgroup)")
      .setBackground(T.rowAlt).setFontColor(T.mutedFg)
      .setFontStyle("italic").setHorizontalAlignment("center")
      .setFontFamily(T.font).setFontSize(T.fontSize);
    sh.setRowHeight(r, 22);
    r++;
  } else {
    const lpColHdr = sh.getRange(r, 1, 1, NC - 1);
    lpColHdr.merge();
    lpColHdr.setValue("Leader").setBackground(T.headerBg).setFontColor(T.sectionFg)
      .setFontWeight("bold").setFontFamily(T.font).setFontSize(T.fontSize)
      .setHorizontalAlignment("left");
    sh.getRange(r, NC).setValue("Notes").setBackground(T.headerBg).setFontColor(T.sectionFg)
      .setFontWeight("bold").setFontFamily(T.font).setFontSize(T.fontSize)
      .setHorizontalAlignment("center");
    sh.setRowHeight(r, 22);
    r++;

    const leaderRows = [];
    subLeaderMap.forEach((v, name) => leaderRows.push({ name, notes: v.notes }));
    leaderRows.sort((a, b) => b.notes - a.notes || a.name.localeCompare(b.name));

    leaderRows.forEach(({ name, notes }, i) => {
      const bg = i % 2 === 0 ? T.rowBase : T.rowAlt;
      const nameRange = sh.getRange(r, 1, 1, NC - 1);
      nameRange.merge();
      nameRange.setValue(name).setBackground(bg).setFontFamily(T.font)
        .setFontSize(T.fontSize).setFontWeight("bold")
        .setFontColor(T.statValFg).setHorizontalAlignment("left");
      sh.getRange(r, NC).setValue(notes).setBackground(bg).setFontFamily(T.font)
        .setFontSize(T.fontSize).setFontWeight("bold")
        .setFontColor(T.statValFg).setHorizontalAlignment("center");
      sh.setRowHeight(r, 22);
      r++;
    });
  }

  sh.getRange(r, 1, 1, NC).setBackground(T.rowBase);
  sh.setRowHeight(r, 10);
  r++;

  if (!ftSubgroupName || !ftAllData) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue("ℹ️  No First-Timer data configured for this subgroup.")
      .setBackground(T.rowAlt).setFontColor(T.mutedFg)
      .setFontSize(9).setHorizontalAlignment("left");
    sh.setRowHeight(r, 22);
    r++;
  } else {
    const { cellData, serviceData } = ftAllData;
    const cellRows    = (cellData.rows    || []).filter(row => row.subgroup === ftSubgroupName);
    const serviceRows = (serviceData.rows || []).filter(row => row.subgroup === ftSubgroupName);
    const cellSubData    = cellData.hierarchyMap[ftSubgroupName]    || null;
    const serviceSubData = serviceData.hierarchyMap[ftSubgroupName] || null;
    const combinedSubTotal = cellRows.length + serviceRows.length;

    sh.getRange(r, 1, 1, NC).merge()
      .setValue("  FIRST-TIMERS")
      .setBackground(T.sectionBg).setFontColor(T.sectionFg)
      .setFontWeight("bold").setFontSize(10).setFontFamily(T.font)
      .setHorizontalAlignment("left").setVerticalAlignment("middle");
    sh.setRowHeight(r, 28);
    r++;

    r = writeFTTallySectionInWideSheet_(sh, r, "CELL FIRST-TIMERS", FT.cellBannerBg,
      cellSubData ? { [ftSubgroupName]: cellSubData } : {}, cellRows.length, NC);
    r++;

    r = writeFTTallySectionInWideSheet_(sh, r, "SERVICE FIRST-TIMERS", FT.serviceBannerBg,
      serviceSubData ? { [ftSubgroupName]: serviceSubData } : {}, serviceRows.length, NC);
    r++;

    r = writeFTGrandTotalRow_(sh, r, "TOTAL FIRST-TIMERS", FT.combinedBannerBg, combinedSubTotal, NC);
  }

  sh.getRange(r, 1, 1, NC).setBackground(T.rowBase);
  sh.setRowHeight(r, 16);

  sh.setFrozenRows(3);
  sh.autoResizeColumns(1, NC);
  sh.setColumnWidth(1, Math.max(sh.getColumnWidth(1), 240));
  for (let c = 2; c <= NC; c++) sh.setColumnWidth(c, Math.max(sh.getColumnWidth(c), 90));
  SpreadsheetApp.flush();
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FOLLOW-UP: LEADER NOTE COUNTER ───────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function runLeaderNoteCounter_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const notesSh = ss.getSheetByName(CFG.NOTES_SHEET);
  const groupSh = ss.getSheetByName(CFG.GROUP_SHEET);
  if (!notesSh || !groupSh) {
    ui.alert("❌  Missing sheet: " + (!notesSh ? CFG.NOTES_SHEET : CFG.GROUP_SHEET));
    return;
  }

  const groupDir = loadGroupDirectory_(groupSh);
  const rows     = notesSh.getDataRange().getValues();
  if (rows.length < 2) { ui.alert("No data in " + CFG.NOTES_SHEET); return; }

  const hdrs = rows[0].map(h => h.toString().trim().toLowerCase());
  const ci = {
    leader: findCol_(hdrs, ["leader"], 1),
    group:  findCol_(hdrs, ["groupname", "group name", "group"], 3),
  };

  const leaderStats = new Map();

  for (let i = 1; i < rows.length; i++) {
    const r      = rows[i];
    const leader = str_(r[ci.leader]);
    const group  = str_(r[ci.group]);
    if (!leader) continue;

    if (!leaderStats.has(leader)) {
      leaderStats.set(leader, { notes: 0, groups: new Set(), subgroups: new Set() });
    }

    const ls = leaderStats.get(leader);
    ls.notes++;

    let resolvedGroup = group;
    let groupInfo     = groupDir.byName.get(group);
    if (!groupInfo && group.includes(",")) {
      const parts = group.split(",").map(p => p.trim()).filter(Boolean);
      resolvedGroup = parts[parts.length - 1];
      groupInfo     = groupDir.byName.get(resolvedGroup);
      if (!groupInfo) { resolvedGroup = parts[0]; groupInfo = groupDir.byName.get(resolvedGroup); }
    }

    if (resolvedGroup) ls.groups.add(resolvedGroup);
    if (groupInfo && groupInfo.subgroup) ls.subgroups.add(groupInfo.subgroup);
  }

  const ranked = [];
  leaderStats.forEach((v, leader) => {
    const subgroup = v.subgroups.size === 1
      ? Array.from(v.subgroups)[0]
      : v.subgroups.size > 1
        ? Array.from(v.subgroups).join(" / ")
        : "(Unassigned)";
    ranked.push({ leader, notes: v.notes, subgroup });
  });
  ranked.sort((a, b) => b.notes - a.notes || a.leader.localeCompare(b.leader));

  const sheetName = "Leader Note Counter";
  let sh = ss.getSheetByName(sheetName);
  if (sh) {
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
    sh.clear(); sh.clearFormats();
    sh.getBandings().forEach(b => b.remove());
  } else {
    sh = ss.insertSheet(sheetName);
  }
  ss.setActiveSheet(sh);

  const LC = {
    titleBg: "#1F3A5F", titleFg: "#FFFFFF",
    subhdrBg: "#F5F7FA", subhdrFg: "#4A5A6A",
    colHdrBg: "#E9EFF7", colHdrFg: "#243447",
    rowBase: "#FFFFFF", rowAlt: "#F8FAFD",
    subtotalBg: "#DCE6F4", subtotalFg: "#10243F",
    grandBg: "#E9EFF7", grandFg: "#243447",
    subgroupFg: "#1F3A5F", border: "#6B7A8A",
  };

  const SOLID        = SpreadsheetApp.BorderStyle.SOLID;
  const SOLID_MEDIUM = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  const addBorders_  = (range, style) =>
    range.setBorder(true, true, true, true, true, true, LC.border, style || SOLID);

  sh.getRange(1, 1, 600, 4)
    .setBackground(LC.rowBase).setFontFamily("Google Sans").setFontSize(10).setFontColor("#000000");

  let r = 1;
  sh.getRange(r, 1, 1, 4).merge().setValue("LEADER NOTE COUNTER")
    .setBackground(LC.titleBg).setFontColor(LC.titleFg).setFontWeight("bold")
    .setFontSize(18).setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontFamily("Google Sans");
  sh.setRowHeight(r, 52);
  r++;

  sh.getRange(r, 1, 1, 4).merge().setValue("Notes logged per leader — ranked highest to lowest")
    .setBackground(LC.titleBg).setFontColor(LC.titleFg).setFontSize(9)
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontFamily("Google Sans");
  sh.setRowHeight(r, 20);
  r += 2;

  function renderBlock_(subRows, blockTitle) {
    sh.getRange(r, 1, 1, 4).merge().setValue(blockTitle)
      .setBackground(LC.subhdrBg).setFontColor(LC.subhdrFg).setFontWeight("bold")
      .setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setFontFamily("Google Sans");
    sh.setRowHeight(r, 26);
    r++;

    const hdrRange = sh.getRange(r, 1, 1, 4);
    hdrRange.setValues([["Rank", "Leader", "Notes Logged", "Subgroup"]])
      .setBackground(LC.colHdrBg).setFontColor(LC.colHdrFg).setFontWeight("bold")
      .setFontSize(10).setFontFamily("Google Sans").setHorizontalAlignment("center");
    addBorders_(hdrRange, SOLID_MEDIUM);
    sh.setRowHeight(r, 22);
    r++;

    subRows.forEach((row, i) => {
      const rowRange = sh.getRange(r, 1, 1, 4);
      rowRange.setValues([[i + 1, row.leader, row.notes, row.subgroup]])
        .setBackground(i % 2 === 0 ? LC.rowBase : LC.rowAlt)
        .setFontFamily("Google Sans").setFontSize(10).setFontColor("#000000");
      sh.getRange(r, 1).setHorizontalAlignment("center");
      sh.getRange(r, 2).setHorizontalAlignment("left").setFontWeight("bold");
      sh.getRange(r, 3).setHorizontalAlignment("right").setFontWeight("bold");
      sh.getRange(r, 4).setHorizontalAlignment("center").setFontColor(LC.subgroupFg)
        .setFontSize(9).setFontWeight("normal");
      addBorders_(rowRange, SOLID);
      r++;
    });

    const subTotal = subRows.reduce((sum, x) => sum + x.notes, 0);
    const stRange  = sh.getRange(r, 1, 1, 4);
    stRange.setValues([["SUBTOTAL", "", subTotal, ""]])
      .setBackground(LC.subtotalBg).setFontColor(LC.subtotalFg).setFontWeight("bold")
      .setFontSize(10).setFontFamily("Google Sans");
    sh.getRange(r, 1).setHorizontalAlignment("left");
    sh.getRange(r, 3).setHorizontalAlignment("right");
    stRange.setBorder(true, true, true, true, false, false, LC.border, SOLID_MEDIUM);
    sh.setRowHeight(r, 22);
    r += 2;
  }

  for (const sub of CFG.SUBGROUPS) {
    const subRows = ranked.filter(x => x.subgroup === sub);
    if (!subRows.length) continue;
    renderBlock_(subRows, sub);
  }

  const unassigned = ranked.filter(x => x.subgroup === "(Unassigned)" || x.subgroup.includes(" / "));
  if (unassigned.length) renderBlock_(unassigned, "OTHER / UNASSIGNED");

  r++;
  const grandTotal = ranked.reduce((sum, x) => sum + x.notes, 0);
  const grandRange = sh.getRange(r, 1, 1, 4);
  grandRange.setValues([["GRAND TOTAL", "", grandTotal, ""]])
    .setBackground(LC.grandBg).setFontColor(LC.grandFg).setFontWeight("bold")
    .setFontSize(13).setFontFamily("Google Sans");
  sh.getRange(r, 1).setHorizontalAlignment("left");
  sh.getRange(r, 3).setHorizontalAlignment("right");
  grandRange.setBorder(true, true, true, true, false, false, LC.border, SOLID_MEDIUM);
  sh.setRowHeight(r, 30);

  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 60); sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 110); sh.setColumnWidth(4, 220);
  SpreadsheetApp.flush();
  ui.alert('✅  Leader Note Counter generated in sheet: "' + sheetName + '"');
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FOLLOW-UP: DELETE SHEETS ──────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function btn_deleteSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const deletable = ss.getSheets()
    .map(s => s.getName())
    .filter(n => !CFG.PROTECTED_SHEETS.includes(n));

  if (deletable.length === 0) {
    ui.alert("No report sheets found to delete.\nAll existing sheets are protected.");
    return;
  }

  const resp = ui.prompt(
    "Delete Report Sheets",
    "The following sheets can be deleted:\n\n" +
    deletable.map((n, i) => `${i + 1}. ${n}`).join("\n") +
    "\n\nType numbers to delete (e.g. 1,3,5), or ALL.\nCancel to exit.",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const input = (resp.getResponseText() || "").trim();
  if (!input) { ui.alert("Nothing entered. No sheets deleted."); return; }

  let toDelete = [];
  if (input.toLowerCase() === "all") {
    toDelete = deletable;
  } else {
    const nums = input.split(",").map(s => parseInt(s.trim(), 10)).filter(n => !isNaN(n));
    toDelete = nums.filter(n => n >= 1 && n <= deletable.length).map(n => deletable[n - 1]);
  }

  if (toDelete.length === 0) { ui.alert("No valid selections. No sheets deleted."); return; }

  const confirm = ui.alert(
    "Confirm Delete",
    `Delete ${toDelete.length} sheet(s)?\n\n• ` + toDelete.join("\n• "),
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  toDelete.forEach(name => { const sh = ss.getSheetByName(name); if (sh) ss.deleteSheet(sh); });
  ui.alert(`✅  Deleted ${toDelete.length} sheet(s).`);
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: BUTTON HANDLERS ─────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function btn_ftRegional()    { runFTReports_(true,  false); }
function btn_ftSubgroup()    { runFTReports_(false, true);  }
function btn_ftBoth()        { runFTReports_(true,  true);  }
function btn_setupFTSheets() { setupFTSheets_(); }
function btn_ftDiagnose()    { runFTDiagnose_(); }


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: DIAGNOSTICS ──────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function runFTDiagnose_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const lines = [];

  Object.values(FT_IMPORT).forEach(sheetName => {
    lines.push(`\n── ${sheetName} ──`);
    const sh = ss.getSheetByName(sheetName);
    if (!sh) { lines.push("  ❌ Sheet not found"); return; }

    const data = sh.getDataRange().getValues();
    if (data.length < 2) { lines.push("  ⚠️ No data rows"); return; }

    const headers = data[0].map(h => String(h).trim());
    lines.push(`  Headers: ${headers.join(" | ")}`);

    const lowerHeaders = headers.map(h => h.toLowerCase());
    const demoIdx = lowerHeaders.indexOf(FT_COL_DEMOGRAPHICS);
    const nameIdx = lowerHeaders.indexOf(FT_COL_FULL_NAME);

    if (demoIdx === -1) { lines.push(`  ❌ "${FT_COL_DEMOGRAPHICS}" column not found`); return; }

    let shown = 0;
    for (let i = 1; i < data.length && shown < 5; i++) {
      const name  = nameIdx >= 0 ? String(data[i][nameIdx] || "").trim() : `row ${i+1}`;
      const demo  = String(data[i][demoIdx] || "").trim();
      if (!demo) continue;
      const { subgroup } = parseFTDemographics_(demo);
      lines.push(`  • ${name}: "${demo}" → ${subgroup || "❌ NO MATCH"}`);
      shown++;
    }
    if (shown === 0) lines.push("  ⚠️ All Demographics cells are empty");

    const unmatched = [];
    for (let i = 1; i < data.length; i++) {
      const demo = String(data[i][demoIdx] || "").trim();
      if (!demo) continue;
      const { subgroup } = parseFTDemographics_(demo);
      if (!subgroup) unmatched.push(demo);
    }
    if (unmatched.length) {
      const unique = [...new Set(unmatched)].slice(0, 5);
      lines.push(`  ⚠️ ${unmatched.length} unmatched rows. Sample unmatched values:`);
      unique.forEach(v => lines.push(`    → "${v}"`));
    } else {
      lines.push(`  ✅ All demographics matched successfully`);
    }
  });

  ui.alert("FT Data Diagnostics", lines.join("\n"), ui.ButtonSet.OK);
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: MAIN RUNNER ──────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function runFTReports_(doRegional, doSubgroup) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const labelResp = ui.prompt(
    "Report Label",
    "Enter a label for this report\n(e.g. Feb 2026 Pulse Check):",
    ui.ButtonSet.OK_CANCEL
  );
  if (labelResp.getSelectedButton() !== ui.Button.OK) return;
  const label = (labelResp.getResponseText() || "").trim() || "FT Report";

  const sheetErrors = [];
  Object.values(FT_IMPORT).forEach(name => {
    if (!ss.getSheetByName(name)) { sheetErrors.push(`"${name}" — sheet not found`); return; }
    const v = validateFTSheetHeaders_(name);
    if (!v.ok) sheetErrors.push(`"${name}" — missing columns: ${v.missing.join(", ")}`);
  });

  if (sheetErrors.length > 0) {
    ui.alert(
      "❌  Import sheet issues:\n\n• " + sheetErrors.join("\n• ") +
      "\n\nRun '⚙️ Setup FT Import Sheets' to create/refresh headers, then re-paste your data."
    );
    return;
  }

  const ftAllData    = loadAllFTData_();
  const totalRecords = ftAllData.cellData.grandTotal + ftAllData.serviceData.grandTotal;

  if (totalRecords === 0) {
    ui.alert("⚠️  No data rows found across the import sheets.\nPlease paste your data and try again.");
    return;
  }

  const dateRange = formatFTDateRange_(ftAllData.globalMin, ftAllData.globalMax);
  const created   = [];

  if (doRegional) {
    const sheetName = `${label} — FT Regional`;
    const sh = getOrCreate_(ss, sheetName);
    renderFTRegionalSheet_(sh, label, dateRange, ftAllData);
    created.push(sheetName);
    ss.setActiveSheet(sh);
  }

  if (doSubgroup) {
    const notesSh = ss.getSheetByName(CFG.NOTES_SHEET);
    const groupSh = ss.getSheetByName(CFG.GROUP_SHEET);
    const groupDir  = (notesSh && groupSh) ? loadGroupDirectory_(groupSh) : { byName: new Map(), bySubgroup: new Map() };
    const notesData = (notesSh && groupSh) ? aggregateNotes_(notesSh, groupDir.byName) : emptyAgg_();

    CFG.SUBGROUPS.forEach(sub => {
      const sheetName = `${label} — ${sub}`;
      const sh = getOrCreate_(ss, sheetName);
      renderCombinedSubgroupSheet_(sh, {
        label, subgroup: sub, grade: false, colourPct: false, groupDir, notesData, ftAllData,
      });
      created.push(sheetName);
    });

    if (!doRegional) ss.setActiveSheet(ss.getSheetByName(created[0]));
  }

  ui.alert(
    `✅  First-Timer Reports generated!\n\n` +
    `${created.length} sheet(s) created:\n` +
    created.map(n => `• ${n}`).join("\n")
  );
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: HEADER VALIDATOR ────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function validateFTSheetHeaders_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return { ok: false, missing: ["Sheet not found"] };
  const data = sh.getDataRange().getValues();
  if (!data.length) return { ok: false, missing: ["No header row found"] };
  const headers  = data[0].map(h => String(h).trim().toLowerCase());
  const required = [FT_COL_FULL_NAME, FT_COL_GROUPS, FT_COL_DEMOGRAPHICS, FT_COL_DATE_ADDED];
  const missing  = required.filter(h => !headers.includes(h));
  return { ok: missing.length === 0, missing };
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: LOAD ALL FT DATA ────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function loadAllFTData_() {
  const cellData    = parseFTSheet_(FT_IMPORT.CELL);
  const serviceData = parseFTSheet_(FT_IMPORT.SERVICE);
  const combinedMap   = mergeFTHierarchy_(cellData.hierarchyMap, serviceData.hierarchyMap);
  const combinedTotal = cellData.grandTotal + serviceData.grandTotal;
  const combinedRows  = [...cellData.rows, ...serviceData.rows];
  const allMins = [cellData.minDate, serviceData.minDate].filter(Boolean);
  const allMaxs = [cellData.maxDate, serviceData.maxDate].filter(Boolean);
  return {
    cellData, serviceData,
    combinedData: { hierarchyMap: combinedMap, grandTotal: combinedTotal, rows: combinedRows },
    globalMin: allMins.length ? allMins.reduce((a, b) => a < b ? a : b) : null,
    globalMax: allMaxs.length ? allMaxs.reduce((a, b) => a > b ? a : b) : null,
  };
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: PARSE IMPORT SHEET ──────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function parseFTSheet_(sheetName) {
  const empty = { hierarchyMap: {}, grandTotal: 0, minDate: null, maxDate: null, rows: [] };
  const sh    = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return empty;

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return empty;

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const ci = {
    fullName:     headers.indexOf(FT_COL_FULL_NAME),
    groups:       headers.indexOf(FT_COL_GROUPS),
    demographics: headers.indexOf(FT_COL_DEMOGRAPHICS),
    dateAdded:    headers.indexOf(FT_COL_DATE_ADDED),
  };

  if (Object.values(ci).includes(-1)) return empty;

  const hierarchyMap = {};
  const rows = [];
  let grandTotal = 0, minDate = null, maxDate = null;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(cell => String(cell).trim() === "")) continue;

    const fullName     = String(row[ci.fullName] || "").trim();
    const rawGroups    = String(row[ci.groups] || "").trim();
    const demographics = String(row[ci.demographics] || "").trim();
    const rawDate      = row[ci.dateAdded];

    if (!fullName) continue;

    if (rawDate) {
      const d = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      if (!isNaN(d)) {
        if (!minDate || d < minDate) minDate = d;
        if (!maxDate || d > maxDate) maxDate = d;
      }
    }

    const { subgroup, fellowship } = parseFTDemographics_(demographics, rawGroups);
    const group  = rawGroups || "(No Group)";
    const subKey = subgroup || "(No Subgroup)";

    if (!hierarchyMap[subKey]) hierarchyMap[subKey] = { total: 0, groups: {} };
    hierarchyMap[subKey].total++;
    hierarchyMap[subKey].groups[group] = (hierarchyMap[subKey].groups[group] || 0) + 1;

    rows.push({ fullName, group, subgroup: subKey, fellowship });
    grandTotal++;
  }

  return { hierarchyMap, grandTotal, minDate, maxDate, rows };
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: MERGE HIERARCHY MAPS ────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function mergeFTHierarchy_(mapA, mapB) {
  const merged = JSON.parse(JSON.stringify(mapA));
  Object.entries(mapB).forEach(([sub, subData]) => {
    if (!merged[sub]) merged[sub] = { total: 0, groups: {} };
    merged[sub].total += subData.total;
    Object.entries(subData.groups).forEach(([g, count]) => {
      merged[sub].groups[g] = (merged[sub].groups[g] || 0) + count;
    });
  });
  return merged;
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: FT REGIONAL SHEET ───────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function renderFTRegionalSheet_(sh, label, dateRange, ftAllData) {
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear(); sh.clearFormats();
  sh.getBandings().forEach(b => b.remove());

  const { cellData, serviceData, combinedData } = ftAllData;
  const NC = FT_TALLY_COLS;
  let r = 1;

  sh.getRange(1, 1, 900, NC)
    .setBackground(FT.rowBase).setFontFamily(FT.font).setFontSize(FT.fontSize);

  sh.getRange(r, 1, 1, NC).merge()
    .setValue(label.toUpperCase())
    .setBackground(FT.titleBg).setFontColor(FT.titleFg)
    .setFontWeight("bold").setFontSize(15).setFontFamily(FT.font)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(r, 44); r++;

  sh.getRange(r, 1, 1, NC).merge()
    .setValue("First-Timer Regional Overview")
    .setBackground(FT.titleBg).setFontColor(FT.subtitleFg)
    .setFontSize(11).setFontFamily(FT.font)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(r, 26); r++;

  sh.getRange(r, 1, 1, NC).merge()
    .setValue("Reporting Period:  " + dateRange)
    .setBackground("#F5F7FA").setFontColor("#4A5A6A")
    .setFontSize(9).setFontFamily(FT.font)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(r, 20); r += 2;

  r = writeFTTallySection_(sh, r, "CELL FIRST-TIMERS", FT.cellBannerBg, cellData.hierarchyMap, cellData.grandTotal, NC);
  r += 2;
  r = writeFTTallySection_(sh, r, "SERVICE FIRST-TIMERS", FT.serviceBannerBg, serviceData.hierarchyMap, serviceData.grandTotal, NC);
  r += 2;
  r = writeFTTallySection_(sh, r, "COMBINED (ALL SOURCES)", FT.combinedBannerBg, combinedData.hierarchyMap, combinedData.grandTotal, NC);

  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 100);
  sh.setFrozenRows(1);
  SpreadsheetApp.flush();
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: TALLY SECTION (standalone 2-col sheet) ──────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function writeFTTallySection_(sh, startRow, sectionTitle, bannerBg, hierarchyMap, grandTotal, numCols) {
  const SOLID     = SpreadsheetApp.BorderStyle.SOLID;
  const SOLID_MED = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  let r = startRow;

  sh.getRange(r, 1, 1, numCols).merge()
    .setValue(sectionTitle)
    .setBackground(bannerBg).setFontColor(FT.bannerFg)
    .setFontFamily(FT.font).setFontSize(12).setFontWeight("bold")
    .setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, FT.border, SOLID_MED);
  sh.setRowHeight(r, 34); r++;

  if (grandTotal === 0 || Object.keys(hierarchyMap).length === 0) {
    sh.getRange(r, 1, 1, numCols).merge()
      .setValue("(No data)")
      .setBackground(FT.rowAlt).setFontColor(FT.mutedFg)
      .setFontFamily(FT.font).setFontSize(FT.fontSize)
      .setFontStyle("italic").setHorizontalAlignment("center");
    sh.setRowHeight(r, 22); r++;
    return r;
  }

  const knownOrder    = FT_KNOWN_SUBGROUPS.map(s => s.toLowerCase());
  const sortedEntries = Object.entries(hierarchyMap).sort((a, b) => {
    const ai = knownOrder.indexOf(a[0].toLowerCase());
    const bi = knownOrder.indexOf(b[0].toLowerCase());
    if (ai !== -1 && bi !== -1) return ai - bi;
    if (ai !== -1) return -1;
    if (bi !== -1) return 1;
    return a[0].localeCompare(b[0]);
  });

  sortedEntries.forEach(([subgroupName, subgroupData]) => {
    sh.getRange(r, 1, 1, numCols).merge()
      .setValue(subgroupName)
      .setBackground(FT.groupHdrBg).setFontColor(FT.groupHdrFg)
      .setFontFamily(FT.font).setFontSize(FT.fontSize).setFontWeight("bold")
      .setHorizontalAlignment("left").setVerticalAlignment("middle");
    sh.setRowHeight(r, 26); r++;

    Object.entries(subgroupData.groups)
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .forEach(([groupName, count], idx) => {
        const bg = idx % 2 === 0 ? FT.subRowBg : FT.subRowAltBg;
        sh.getRange(r, 1, 1, numCols - 1).merge()
          .setValue("     └  " + groupName)
          .setBackground(bg).setFontColor(FT.subNameFg)
          .setFontFamily(FT.font).setFontSize(FT.fontSize).setFontWeight("normal")
          .setHorizontalAlignment("left").setVerticalAlignment("middle");
        sh.getRange(r, numCols)
          .setValue(count)
          .setBackground(bg).setFontColor(FT.subCountFg)
          .setFontFamily(FT.font).setFontSize(FT.fontSize).setFontWeight("bold")
          .setHorizontalAlignment("center").setVerticalAlignment("middle");
        sh.setRowHeight(r, 22); r++;
      });

    sh.getRange(r, 1, 1, numCols - 1).merge()
      .setValue(subgroupName + "  —  Subtotal")
      .setBackground(FT.subtotalBg).setFontColor(FT.subtotalFg)
      .setFontFamily(FT.font).setFontSize(FT.fontSize).setFontWeight("bold")
      .setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBorder(true, true, true, false, false, false, FT.border, SOLID);
    sh.getRange(r, numCols)
      .setValue(subgroupData.total)
      .setBackground(FT.subtotalBg).setFontColor(FT.subtotalFg)
      .setFontFamily(FT.font).setFontSize(FT.fontSize).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBorder(true, false, true, true, false, false, FT.border, SOLID);
    sh.setRowHeight(r, 24); r++;
    sh.setRowHeight(r, 6); r++;
  });

  sh.getRange(r, 1, 1, numCols - 1).merge()
    .setValue("GRAND TOTAL")
    .setBackground(FT.grandTotalBg).setFontColor(FT.grandTotalFg)
    .setFontFamily(FT.font).setFontSize(11).setFontWeight("bold")
    .setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setBorder(true, true, true, false, false, false, FT.border, SOLID_MED);
  sh.getRange(r, numCols)
    .setValue(grandTotal)
    .setBackground(FT.grandTotalBg).setFontColor(FT.grandCountFg)
    .setFontFamily(FT.font).setFontSize(11).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, false, true, true, false, false, FT.border, SOLID_MED);
  sh.setRowHeight(r, 30); r++;
  return r;
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: GRAND TOTAL ROW (wide sheet) ────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function writeFTGrandTotalRow_(sh, startRow, label, bannerBg, total, nc) {
  const SOLID_MED = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  let r = startRow;
  sh.getRange(r, 1, 1, nc - 1).merge()
    .setValue(label)
    .setBackground(bannerBg).setFontColor("#FFFFFF")
    .setFontFamily(FT.font).setFontSize(13).setFontWeight("bold")
    .setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setBorder(true, true, true, false, false, false, FT.border, SOLID_MED);
  sh.getRange(r, nc)
    .setValue(total)
    .setBackground(bannerBg).setFontColor("#FFFFFF")
    .setFontFamily(FT.font).setFontSize(13).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, false, true, true, false, false, FT.border, SOLID_MED);
  sh.setRowHeight(r, 36); r++;
  return r;
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: FT TALLY SECTION (wide sheet) ───────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function writeFTTallySectionInWideSheet_(sh, startRow, sectionTitle, bannerBg, hierarchyMap, grandTotal, nc) {
  const SOLID     = SpreadsheetApp.BorderStyle.SOLID;
  const SOLID_MED = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  const TC        = FT_TALLY_COLS;
  let r = startRow;

  sh.getRange(r, 1, 1, nc).merge()
    .setValue(sectionTitle)
    .setBackground(bannerBg).setFontColor(FT.bannerFg)
    .setFontFamily(FT.font).setFontSize(11).setFontWeight("bold")
    .setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, FT.border, SOLID_MED);
  sh.setRowHeight(r, 28); r++;

  if (grandTotal === 0 || Object.keys(hierarchyMap).length === 0) {
    sh.getRange(r, 1, 1, nc).merge()
      .setValue("(No data)")
      .setBackground(FT.rowAlt).setFontColor(FT.mutedFg)
      .setFontFamily(FT.font).setFontSize(FT.fontSize)
      .setFontStyle("italic").setHorizontalAlignment("center");
    sh.setRowHeight(r, 22); r++;
    return r;
  }

  const allGroups = {};
  Object.values(hierarchyMap).forEach(subData => {
    Object.entries(subData.groups).forEach(([g, count]) => {
      allGroups[g] = (allGroups[g] || 0) + count;
    });
  });

  Object.entries(allGroups)
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .forEach(([groupName, count], idx) => {
      const bg = idx % 2 === 0 ? FT.subRowBg : FT.subRowAltBg;
      sh.getRange(r, 1)
        .setValue("     └  " + groupName)
        .setBackground(bg).setFontColor(FT.subNameFg)
        .setFontFamily(FT.font).setFontSize(FT.fontSize).setFontWeight("normal")
        .setHorizontalAlignment("left").setVerticalAlignment("middle");
      sh.getRange(r, 2)
        .setValue(count)
        .setBackground(bg).setFontColor(FT.subCountFg)
        .setFontFamily(FT.font).setFontSize(FT.fontSize).setFontWeight("bold")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
      if (nc > TC) sh.getRange(r, TC + 1, 1, nc - TC).setBackground(bg);
      sh.setRowHeight(r, 22); r++;
    });

  sh.getRange(r, 1)
    .setValue("Section Total")
    .setBackground(FT.subtotalBg).setFontColor(FT.subtotalFg)
    .setFontFamily(FT.font).setFontSize(FT.fontSize).setFontWeight("bold")
    .setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setBorder(true, true, true, false, false, false, FT.border, SOLID);
  sh.getRange(r, 2)
    .setValue(grandTotal)
    .setBackground(FT.subtotalBg).setFontColor(FT.subtotalFg)
    .setFontFamily(FT.font).setFontSize(FT.fontSize).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, false, true, true, false, false, FT.border, SOLID);
  if (nc > TC) sh.getRange(r, TC + 1, 1, nc - TC).setBackground(FT.subtotalBg);
  sh.setRowHeight(r, 24); r++;
  return r;
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: DEMOGRAPHICS PARSER ─────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

const FT_SUBGROUP_ALIASES = {
  "central east subgroup a":     "Central East Subgroup A",
  "central east sga":            "Central East Subgroup A",
  "blw canada central-east sga": "Central East Subgroup A",
  "blw canada central east sga": "Central East Subgroup A",
  "canada central-east sga":     "Central East Subgroup A",
  "canada central east sga":     "Central East Subgroup A",
  "central-east sga":            "Central East Subgroup A",
  "zz | central-east sga":       "Central East Subgroup A",
  "central east subgroup b":     "Central East Subgroup B",
  "central east sgb":            "Central East Subgroup B",
  "blw canada central-east sgb": "Central East Subgroup B",
  "blw canada central east sgb": "Central East Subgroup B",
  "canada central-east sgb":     "Central East Subgroup B",
  "canada central east sgb":     "Central East Subgroup B",
  "central-east sgb":            "Central East Subgroup B",
  "zz | central-east sgb":       "Central East Subgroup B",
  "central subgroup a":          "Central Subgroup A",
  "central sga":                 "Central Subgroup A",
  "blw canada central sga":      "Central Subgroup A",
  "canada central sga":          "Central Subgroup A",
  "zz | central sga":            "Central Subgroup A",
  "central subgroup b":          "Central Subgroup B",
  "central sgb":                 "Central Subgroup B",
  "blw canada central sgb":      "Central Subgroup B",
  "canada central sgb":          "Central Subgroup B",
  "zz | central sgb":            "Central Subgroup B",
  "west subgroup a":             "West Subgroup A",
  "west sga":                    "West Subgroup A",
  "blw canada west sga":         "West Subgroup A",
  "canada west sga":             "West Subgroup A",
  "zz | west sga":               "West Subgroup A",
  "west subgroup b":             "West Subgroup B",
  "west sgb":                    "West Subgroup B",
  "blw canada west sgb":         "West Subgroup B",
  "canada west sgb":             "West Subgroup B",
  "zz | west sgb":               "West Subgroup B",
};

function parseFTDemographics_(raw, groupsRaw) {
  var sources = [raw, groupsRaw].filter(function(v) { return v && String(v).trim(); });

  for (var si = 0; si < sources.length; si++) {
    var trimmed = String(sources[si]).trim();
    var lower   = trimmed.toLowerCase();

    for (var ki = 0; ki < FT_KNOWN_SUBGROUPS.length; ki++) {
      var known = FT_KNOWN_SUBGROUPS[ki];
      if (lower.indexOf(known.toLowerCase()) !== -1) {
        var esc1 = known.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        var rem1 = trimmed.replace(new RegExp(esc1, 'i'), '').replace(/^[\s,]+|[\s,]+$/g, '').trim();
        return { subgroup: known, fellowship: rem1 };
      }
    }

    var aliasKeys = Object.keys(FT_SUBGROUP_ALIASES);
    for (var ai = 0; ai < aliasKeys.length; ai++) {
      var alias     = aliasKeys[ai];
      var canonical = FT_SUBGROUP_ALIASES[alias];
      if (lower.indexOf(alias) !== -1) {
        var esc2 = alias.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        var rem2 = trimmed.replace(new RegExp(esc2, 'i'), '').replace(/^[\s,]+|[\s,]+$/g, '').trim();
        return { subgroup: canonical, fellowship: rem2 };
      }
    }

    if (/central.?east/i.test(lower) && /sga|subgroup.?a/i.test(lower))
      return { subgroup: "Central East Subgroup A", fellowship: trimmed };
    if (/central.?east/i.test(lower) && /sgb|subgroup.?b/i.test(lower))
      return { subgroup: "Central East Subgroup B", fellowship: trimmed };
    if (/central/i.test(lower) && !/east/i.test(lower) && /sga|subgroup.?a/i.test(lower))
      return { subgroup: "Central Subgroup A", fellowship: trimmed };
    if (/central/i.test(lower) && !/east/i.test(lower) && /sgb|subgroup.?b/i.test(lower))
      return { subgroup: "Central Subgroup B", fellowship: trimmed };
    if (/west/i.test(lower) && /sga|subgroup.?a/i.test(lower))
      return { subgroup: "West Subgroup A", fellowship: trimmed };
    if (/west/i.test(lower) && /sgb|subgroup.?b/i.test(lower))
      return { subgroup: "West Subgroup B", fellowship: trimmed };
  }

  return { subgroup: "", fellowship: raw ? String(raw).trim() : "" };
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── FIRST-TIMER: SETUP SHEETS ────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function setupFTSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const created = [];

  [
    { name: FT_IMPORT.CELL,    label: "Cell FTs" },
    { name: FT_IMPORT.SERVICE, label: "Service FTs" },
  ].forEach(({ name, label }) => {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); created.push(`"${name}"`); }
    addFTImportGuide_(sh, label);
  });

  const firstSheet = ss.getSheetByName(FT_IMPORT.CELL);
  if (firstSheet) ss.setActiveSheet(firstSheet);

  if (created.length > 0) {
    ui.alert(
      `✅  Setup complete!\n\nCreated: ${created.join(", ")}\n\n` +
      `Next steps:\n` +
      `1. Paste your Elvanto CSV (headers + data) starting from row 1\n` +
      `2. Ensure columns: Full Name | Groups | Demographics | Date Added\n` +
      `3. Run reports from the 📋 First-Timer Tools menu`
    );
  } else {
    ui.alert(`ℹ️  Both import sheets already exist — header rows refreshed.\n\nPaste your data into each tab and generate reports from the menu.`);
  }
}

function addFTImportGuide_(sheet, label) {
  const headers = ["Date Added", "Full Name", "Groups", "Demographics"];
  sheet.getRange(1, 1, 1, 6).clearContent().clearFormat();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(FT.titleBg).setFontColor(FT.titleFg)
    .setFontWeight("bold").setFontFamily(FT.font).setFontSize(FT.fontSize)
    .setHorizontalAlignment("center");
  sheet.getRange(1, 6)
    .setValue(`← Paste your ${label} Elvanto CSV here (row 1 = headers, data from row 2)`)
    .setBackground("#F5F7FA").setFontColor(FT.mutedFg)
    .setFontStyle("italic").setFontFamily(FT.font).setFontSize(9);
  sheet.setColumnWidth(1, 180); sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 200); sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(6, 420);
  sheet.setFrozenRows(1); sheet.setRowHeight(1, 34);
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── SHARED RENDERING HELPERS ──────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function paintSheet_(sh, numRows, numCols) {
  sh.getRange(1, 1, numRows, numCols)
    .setBackground(T.rowBase).setFontFamily(T.font).setFontSize(T.fontSize);
}

function titleBlock_(sh, r, label, rangeText, subtitle, numCols) {
  mergeWrite_(sh, r, 1, 1, numCols, label.toUpperCase(),
    { bg: T.titleBg, fg: T.titleFg, bold: true, size: 15, align: "center", height: 44 });
  r++;
  mergeWrite_(sh, r, 1, 1, numCols, subtitle,
    { bg: T.titleBg, fg: T.titleFg, bold: false, size: 11, align: "center", height: 26 });
  r++;
  mergeWrite_(sh, r, 1, 1, numCols, `Reporting Period:  ${rangeText}`,
    { bg: T.periodBg, fg: T.periodFg, size: 9, align: "center", height: 20 });
  r += 2;
  return r;
}

function sectionLabel_(sh, r, text, numCols) {
  mergeWrite_(sh, r, 1, 1, numCols, text,
    { bg: T.sectionBg, fg: T.sectionFg, bold: true, size: 10, height: 24, align: "center" });
  return r + 1;
}

function writeHeader_(sh, r, cols, numCols) {
  sh.getRange(r, 1, 1, numCols)
    .setBackground(T.headerBg).setFontColor(T.headerFg)
    .setFontWeight("bold").setFontFamily(T.font).setFontSize(T.fontSize)
    .setHorizontalAlignment("center");
  sh.getRange(r, 1, 1, cols.length).setValues([cols]);
}

function writeDataRow_(sh, r, vals, bg, numCols) {
  sh.getRange(r, 1, 1, numCols)
    .setBackground(bg).setFontFamily(T.font).setFontSize(T.fontSize)
    .setFontColor("#000000").setFontWeight("normal");
  sh.getRange(r, 1, 1, vals.length).setValues([vals]);
  sh.getRange(r, 1).setHorizontalAlignment("left");
  if (vals.length > 1) sh.getRange(r, 2, 1, vals.length - 1).setHorizontalAlignment("center");
}

function writeTotalRow_(sh, r, vals, numCols) {
  sh.getRange(r, 1, 1, numCols)
    .setBackground(T.totalBg).setFontColor(T.totalFg)
    .setFontWeight("bold").setFontFamily(T.font).setFontSize(T.fontSize);
  sh.getRange(r, 1, 1, vals.length).setValues([vals]);
  sh.getRange(r, 1).setHorizontalAlignment("left");
  if (vals.length > 1) sh.getRange(r, 2, 1, vals.length - 1).setHorizontalAlignment("center");
}

function footerRow_(sh, r, text, numCols) {
  sh.getRange(r, 1, 1, numCols).setBackground(T.rowBase);
  sh.getRange(r, 1).setValue(text)
    .setFontWeight("bold").setFontColor("#000000")
    .setFontFamily(T.font).setFontSize(T.fontSize);
}

function mergeWrite_(sh, r, c, rs, cs, val, fmt) {
  const range = sh.getRange(r, c, rs, cs);
  if (rs > 1 || cs > 1) range.merge();
  range.setValue(val).setVerticalAlignment("middle").setFontFamily(T.font);
  if (fmt.bg) range.setBackground(fmt.bg);
  if (fmt.fg) range.setFontColor(fmt.fg);
  if (fmt.bold !== undefined) range.setFontWeight(fmt.bold ? "bold" : "normal");
  if (fmt.size) range.setFontSize(fmt.size);
  if (fmt.align) range.setHorizontalAlignment(fmt.align);
  if (fmt.height) sh.setRowHeight(r, fmt.height);
}

function finalise_(sh, numCols) {
  sh.setFrozenRows(4);
  sh.autoResizeColumns(1, numCols);
  sh.setColumnWidth(1, Math.max(sh.getColumnWidth(1), 220));
  SpreadsheetApp.flush();
}

function getOrCreate_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (sh) {
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
    sh.clear(); sh.clearFormats();
    sh.getBandings().forEach(b => b.remove());
    return sh;
  }
  return ss.insertSheet(name);
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── GRADE LOGIC ───────────────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function gradeScore_(reachRatio, groupsActive, totalGroups) {
  const reach = Math.max(0, Math.min(1, reachRatio || 0)) * 80;
  const mob   = totalGroups > 0
    ? Math.max(0, Math.min(1, (groupsActive || 0) / totalGroups)) * 20
    : 0;
  return reach + mob;
}

function gradeLabel_(score) {
  const s = Math.max(0, Math.min(100, Math.round((score || 0) * 10) / 10));
  if (s >= 90) return "A+";
  if (s >= 85) return "A";
  if (s >= 80) return "A-";
  if (s >= 77) return "B+";
  if (s >= 73) return "B";
  if (s >= 70) return "B-";
  if (s >= 67) return "C+";
  if (s >= 63) return "C";
  if (s >= 60) return "C-";
  if (s >= 57) return "D+";
  if (s >= 53) return "D";
  if (s >= 50) return "D-";
  return "F";
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── UTILITIES ─────────────────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function formatFTDateRange_(minDate, maxDate) {
  if (!minDate) return "(date range unavailable)";
  const tz  = CFG.TIMEZONE;
  const fmt = d => Utilities.formatDate(d, tz, "MMMM d, yyyy");
  if (!maxDate || minDate.toDateString() === maxDate.toDateString()) return fmt(minDate);
  return `${fmt(minDate)} – ${fmt(maxDate)}`;
}

function resolveSubgroup_(input) {
  const n = parseInt(input, 10);
  if (!isNaN(n) && n >= 1 && n <= CFG.SUBGROUPS.length) return CFG.SUBGROUPS[n - 1];
  return CFG.SUBGROUPS.find(s => s.toLowerCase().includes(input.toLowerCase())) || null;
}

function findCol_(hdrs, names, fallback) {
  for (const n of names) {
    const i = hdrs.indexOf(n.toLowerCase());
    if (i >= 0) return i;
  }
  return fallback;
}

function str_(v)   { return (v === null || v === undefined) ? "" : String(v).trim(); }
function toNum_(v) {
  if (typeof v === "number") return v;
  const n = parseFloat(String(v || "").replace(/[, ]+/g, ""));
  return isNaN(n) ? 0 : n;
}
function safeRatio_(num, den) {
  const n = toNum_(num), d = toNum_(den);
  return d ? Math.max(0, Math.min(1, n / d)) : 0;
}
function pctText_(r) {
  return `${Math.round(Math.max(0, Math.min(1, r || 0)) * 1000) / 10}%`;
}
function coerceDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const s = String(v || "").trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── EMAIL REPORTS AS PDF ──────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function btn_emailReports() { runEmailReports_(); }

function runEmailReports_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const sendable = ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return !CFG.PROTECTED_SHEETS.includes(n); });

  if (sendable.length === 0) {
    ui.alert("No report sheets found to send.\nGenerate reports first.");
    return;
  }

  const listText = sendable.map(function(n, i) { return (i + 1) + ".  " + n; }).join("\n");
  const selResp = ui.prompt(
    "Select Sheets to Email",
    "Available report sheets:\n\n" + listText + "\n\nEnter numbers to send (e.g. 1,3,5) or ALL:",
    ui.ButtonSet.OK_CANCEL
  );
  if (selResp.getSelectedButton() !== ui.Button.OK) return;

  const input = (selResp.getResponseText() || "").trim();
  if (!input) { ui.alert("Nothing entered. No emails sent."); return; }

  let toSend = [];
  if (input.toLowerCase() === "all") {
    toSend = sendable.slice();
  } else {
    const nums = input.split(",").map(function(s) { return parseInt(s.trim(), 10); });
    toSend = nums
      .filter(function(n) { return !isNaN(n) && n >= 1 && n <= sendable.length; })
      .map(function(n) { return sendable[n - 1]; });
  }

  if (toSend.length === 0) { ui.alert("No valid selections. No emails sent."); return; }

  const sent = [], skipped = [];

  for (let i = 0; i < toSend.length; i++) {
    const name = toSend[i];
    const emailResp = ui.prompt(
      "Email for sheet " + (i + 1) + " of " + toSend.length,
      "Recipient email for:\n\"" + name + "\"\n\n(Leave blank or Cancel to skip)",
      ui.ButtonSet.OK_CANCEL
    );

    if (emailResp.getSelectedButton() !== ui.Button.OK) { skipped.push(name); continue; }
    const email = (emailResp.getResponseText() || "").trim();
    if (!email) { skipped.push(name); continue; }
    if (email.indexOf("@") === -1 || email.indexOf(".") === -1) {
      ui.alert("\"" + email + "\" doesn't look like a valid email.\nSkipping: \"" + name + "\"");
      skipped.push(name); continue;
    }

    try {
      const pdf     = exportSheetAsPDF_(ss, name);
      const subject = name;
      const body    = "Please find attached your report: " + name + ".\n\nGenerated by BLW Report Suite.";
      GmailApp.sendEmail(email, subject, body, {
        attachments: [pdf.setName(name + ".pdf")],
        name: "BLW Report Suite",
      });
      sent.push({ name, email });
    } catch (e) {
      ui.alert("Error sending \"" + name + "\" to " + email + ":\n" + e.message);
      skipped.push(name);
    }
  }

  let summary = "";
  if (sent.length > 0) {
    summary += "Sent " + sent.length + " report(s):\n";
    for (let j = 0; j < sent.length; j++) {
      summary += "  - " + sent[j].name + "\n    to: " + sent[j].email + "\n";
    }
  }
  if (skipped.length > 0) {
    if (summary) summary += "\n";
    summary += "Skipped " + skipped.length + " report(s):\n";
    for (let k = 0; k < skipped.length; k++) summary += "  - " + skipped[k] + "\n";
  }
  ui.alert("Email Summary", summary.trim(), ui.ButtonSet.OK);
}


// ═══════════════════════════════════════════════════════════════════════════════
// ─── PDF EXPORT HELPER ─────────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════════

function exportSheetAsPDF_(ss, sheetName) {
  const sh      = ss.getSheetByName(sheetName);
  const sheetId = sh.getSheetId();
  const ssId    = ss.getId();

  const url = "https://docs.google.com/spreadsheets/d/" + ssId +
    "/export?exportFormat=pdf&format=pdf&size=A4&portrait=false" +
    "&fitw=true&sheetnames=false&printtitle=false" +
    "&pagenumbers=false&gridlines=false&fzr=false&gid=" + sheetId;

  const token    = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("PDF export failed (HTTP " + response.getResponseCode() + ")");
  }

  return response.getBlob().setContentType("application/pdf");
}
