/**
 * SIMC RAD QUEST v3 — gamified staff development
 * Google Apps Script backend (bound to one Google Sheet).
 *
 * The app itself (the "app shell") is a static site in /docs, hosted on GitHub Pages.
 * It talks to this script over HTTPS: every action is a POST to this web app's /exec URL.
 *
 * Sheet tabs (created by Rad Quest ▸ Run setup):
 *   Config, Staff, Quests, Submissions, Ledger, Shop, Redemptions, Pets, Boss
 *
 * How it flows:
 *   1. Staff open their personal link, pick a quest and submit it inside the app (text + photos).
 *   2. The submission lands in "Submissions" as PENDING. Photos are saved in a private Drive folder.
 *   3. An admin opens Admin ▸ Review in the app, looks at the proof and approves or rejects.
 *      Approval writes coins + XP to the Ledger.
 *   4. Next time the staff member opens the app, they see what happened while they were away.
 *   5. Buying a real reward deducts coins and creates a row in "Redemptions" for the admin to fulfil.
 *
 * The Ledger is append-only: never edit or delete rows. Corrections = new rows (Type ADJUST).
 */

var VERSION = '3.0';

// ───────────────────────────── constants ─────────────────────────────

var TAB = {
  CONFIG: 'Config', STAFF: 'Staff', QUESTS: 'Quests', SUBS: 'Submissions', LEDGER: 'Ledger',
  SHOP: 'Shop', REDEEM: 'Redemptions', PETS: 'Pets', BOSS: 'Boss'
};

// New columns are always added at the END so an existing sheet keeps working after re-running setup.
var HEADERS = {
  Config: ['Key', 'Value', 'Notes'],
  Staff: ['Name', 'Email', 'Role', 'Team', 'Active', 'OnLeave', 'PetName', 'Token', 'PersonalLink', 'IsAdmin'],
  Quests: ['QuestID', 'Title', 'Stat', 'Roles', 'Coins', 'XP', 'Proof', 'Repeat', 'Active', 'Description', 'AdminNotes',
    'Link', 'StartDate', 'EndDate', 'MaxClaims'],
  Submissions: ['Timestamp', 'Email', 'Name', 'QuestID', 'QuestTitle', 'WhatIDid', 'Evidence', 'Flag',
    'Approve', 'Reject', 'Status', 'ReviewedBy', 'ReviewedAt', 'ReviewNote', 'CoinsAwarded', 'XPAwarded'],
  Ledger: ['Timestamp', 'Email', 'Type', 'Ref', 'Coins', 'XP', 'Stat', 'Note'],
  Shop: ['ItemID', 'Name', 'Category', 'Price', 'Effect', 'Description', 'Active'],
  Redemptions: ['Timestamp', 'Email', 'Name', 'ItemID', 'ItemName', 'Price', 'Fulfilled', 'Refund',
    'Status', 'HandledBy', 'HandledAt', 'Note'],
  Pets: ['Email', 'Food', 'FoodUpdatedAt', 'Cosmetics', 'LastPetAt', 'Room', 'LastSeenAt', 'Hatched'],
  Boss: ['BossID', 'Name', 'MaxHP', 'HP', 'Stats', 'RewardCoins', 'StartDate', 'EndDate', 'Status', 'Description']
};

var STATS = [
  { key: 'SKILL', name: 'Skill', color: '#ff9f43', about: 'Clinical & technical' },
  { key: 'SHIELD', name: 'Shield', color: '#4dabff', about: 'Radiation safety' },
  { key: 'WISDOM', name: 'Wisdom', color: '#b48cff', about: 'Learning & teaching' },
  { key: 'HEART', name: 'Heart', color: '#ff6fa3', about: 'Patient care & teamwork' },
  { key: 'VITAL', name: 'Vitality', color: '#7cff6b', about: 'Wellness' }
];

var REPEATS = ['once', 'daily', 'weekly', 'monthly', 'unlimited'];
var MAX_PHOTOS = 3;
var MAX_PHOTO_CHARS = 1500000; // base64 length, about 1.1 MB per photo after the app shrinks it

// Tabs that rarely change are kept in the script cache for a few minutes so the app opens faster.
var CACHED_TABS = { Config: 1, Staff: 1, Quests: 1, Shop: 1, Boss: 1 };
var CACHE_SECONDS = 300;

// ───────────────────────────── menu ─────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Rad Quest')
    .addItem('1. Run setup (safe to re-run)', 'setup')
    .addItem('2. Generate personal links', 'generatePersonalLinks')
    .addItem('3. Email personal links to staff', 'emailPersonalLinks')
    .addSeparator()
    .addItem('Show balances summary', 'showBalances')
    .addItem('Show admin PIN', 'showAdminPin')
    .addToUi();
}

// ───────────────────────────── sheet helpers ─────────────────────────────

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) { return ss_().getSheetByName(name); }

var MEMO_ = {};      // per-run memory of tabs already read
var CFG_CACHE_ = null;

function cache_() { try { return CacheService.getScriptCache(); } catch (err) { return null; } }

/** Forget what we know about a tab after writing to it. */
function touch_(name) {
  delete MEMO_[name];
  if (name === TAB.CONFIG) CFG_CACHE_ = null;
  if (CACHED_TABS[name]) { var c = cache_(); if (c) { try { c.remove('tab_' + name); } catch (err) { } } }
}

function clearAllCaches_() {
  MEMO_ = {}; CFG_CACHE_ = null;
  var c = cache_();
  if (c) { try { c.removeAll(Object.keys(CACHED_TABS).map(function (n) { return 'tab_' + n; })); } catch (err) { } }
}

function readSheet_(name) {
  var sh = sheet_(name);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var empty = true;
    for (var k = 0; k < row.length; k++) { if (row[k] !== '' && row[k] !== false && row[k] !== null) { empty = false; break; } }
    if (empty) continue;
    var o = { _row: r + 1 };
    for (var c = 0; c < head.length; c++) o[head[c]] = row[c];
    out.push(o);
  }
  return out;
}

/**
 * Read a tab as an array of objects keyed by header. Each object gets _row (1-based sheet row).
 * Note: dates may come back as Date objects (fresh read) or ISO strings (from cache): always wrap in new Date().
 */
function readTable_(name) {
  if (MEMO_[name]) return MEMO_[name];
  var rows = null;
  var c = CACHED_TABS[name] ? cache_() : null;
  if (c) {
    try { var hit = c.get('tab_' + name); if (hit) rows = JSON.parse(hit); } catch (err) { rows = null; }
  }
  if (!rows) {
    rows = readSheet_(name);
    if (c) {
      try { var s = JSON.stringify(rows); if (s.length < 90000) c.put('tab_' + name, s, CACHE_SECONDS); } catch (err) { }
    }
  }
  MEMO_[name] = rows;
  return rows;
}

/** Last row that really has data in column A (ignores empty checkbox rows). */
function lastDataRow_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var col = sh.getRange(1, 1, last, 1).getValues();
  for (var r = col.length - 1; r >= 1; r--) { if (col[r][0] !== '' && col[r][0] !== null) return r + 1; }
  return 1;
}

function appendObj_(name, obj) {
  var sh = sheet_(name);
  var head = HEADERS[name];
  var row = lastDataRow_(sh) + 1;
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200);
  sh.getRange(row, 1, 1, head.length).setValues([head.map(function (h) { return obj[h] === undefined ? '' : obj[h]; })]);
  touch_(name);
  return row;
}

/** Update several cells of one row in a single write. */
function setCells_(name, row, obj) {
  var sh = sheet_(name);
  var head = HEADERS[name];
  var rng = sh.getRange(row, 1, 1, head.length);
  var vals = rng.getValues()[0];
  Object.keys(obj).forEach(function (k) {
    var i = head.indexOf(k);
    if (i < 0) throw new Error('No column ' + k + ' in ' + name);
    vals[i] = obj[k];
  });
  rng.setValues([vals]);
  touch_(name);
}

function setCell_(name, row, header, value) {
  var col = HEADERS[name].indexOf(header) + 1;
  if (col < 1) throw new Error('No column ' + header + ' in ' + name);
  sheet_(name).getRange(row, col).setValue(value);
  touch_(name);
}

function cfg_(key, fallback) {
  if (!CFG_CACHE_) {
    CFG_CACHE_ = {};
    readTable_(TAB.CONFIG).forEach(function (r) { CFG_CACHE_[String(r.Key)] = r.Value; });
  }
  var v = CFG_CACHE_[key];
  return (v === undefined || v === '' || v === null) ? fallback : v;
}

function setCfg_(key, value, notes) {
  var rows = readSheet_(TAB.CONFIG);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].Key === key) { setCell_(TAB.CONFIG, rows[i]._row, 'Value', value); return; }
  }
  appendObj_(TAB.CONFIG, { Key: key, Value: value, Notes: notes || '' });
}

function norm_(email) { return String(email || '').trim().toLowerCase(); }
function truthy_(v) { return v === true || String(v).toUpperCase() === 'TRUE' || String(v).toUpperCase() === 'YES'; }
function num_(v) { var n = Number(v); return isNaN(n) ? 0 : n; }
function tz_() { return ss_().getSpreadsheetTimeZone() || 'Asia/Manila'; }

/** Run fn while holding the script lock, reading fresh data (so two taps at once can't double-spend). */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  MEMO_ = {}; CFG_CACHE_ = null;
  try { return fn(); } finally { lock.releaseLock(); }
}

// ───────────────────────────── pure game rules (no Google services) ─────────────────────────────

/** Total XP needed to reach a level. Lv1 = 0, Lv2 = 100, Lv3 = 300, Lv4 = 600, Lv5 = 1000 ... */
function xpForLevel_(level) { return 100 * (level - 1) * level / 2; }

function levelFromXp_(xp) {
  var lv = 1;
  while (xp >= xpForLevel_(lv + 1)) lv++;
  return lv;
}

/** Stat points 0–10 from XP earned in that stat. */
function statPoints_(statXp, xpPerPoint) {
  return Math.max(0, Math.min(10, Math.floor(statXp / xpPerPoint)));
}

/**
 * Pet form depends on level AND balance across the five stats,
 * so a well-rounded person gets the best evolution.
 */
function petForm_(level, statVals) {
  var minStat = Math.min.apply(null, statVals);
  if (level >= 9 && minStat >= 7) return 'RADIANT GUARDIAN';
  if (level >= 6 && minStat >= 4) return 'GUARDIAN';
  if (level >= 6) return 'PHANTOM+';
  if (level >= 3) return 'PHANTOM';
  return 'PIXEL PUP';
}

/**
 * Hunger decays lazily: 1 point every `decayHours`. On leave = no decay.
 * Returns the current food value (0..max).
 */
function currentFood_(storedFood, updatedAtMs, nowMs, decayHours, onLeave) {
  if (onLeave || !updatedAtMs) return Math.max(0, storedFood);
  var steps = Math.floor((nowMs - updatedAtMs) / (decayHours * 3600 * 1000));
  return Math.max(0, storedFood - Math.max(0, steps));
}

function moodFor_(food, max) {
  if (food <= 0) return 'Asleep. Needs a snack.';
  if (food < max * 0.4) return 'Hungry...';
  if (food < max * 0.8) return 'Content';
  return 'Full & happy!';
}

/** Parse a shop Effect like "food:2", "cosmetic:cape", "room:ct". */
function parseEffect_(effect) {
  var s = String(effect || '').trim();
  if (!s) return { kind: 'none', value: '' };
  var i = s.indexOf(':');
  if (i < 0) return { kind: s.toLowerCase(), value: '' };
  return { kind: s.slice(0, i).trim().toLowerCase(), value: s.slice(i + 1).trim() };
}

/** Period bucket so a weekly quest can be earned once per calendar week, etc. */
function periodKey_(repeat, date, tz) {
  var r = String(repeat || 'unlimited').toLowerCase();
  if (r === 'once') return 'once';
  if (r === 'daily') return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  if (r === 'weekly') return Utilities.formatDate(date, tz, "YYYY-'W'ww");
  if (r === 'monthly') return Utilities.formatDate(date, tz, 'yyyy-MM');
  return null; // unlimited
}

function questFitsRole_(rolesCell, role) {
  var roles = String(rolesCell || 'ALL').split(',').map(function (s) { return s.trim().toLowerCase(); });
  return roles.indexOf('all') >= 0 || roles.indexOf(String(role || '').trim().toLowerCase()) >= 0;
}

function startOfDay_(d) { var x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function endOfDay_(d) { var x = new Date(d); x.setHours(23, 59, 59, 999); return x; }

/** A quest is live when Active is ticked and today is inside its optional StartDate–EndDate window. */
function questIsLive_(q, now) {
  if (!truthy_(q.Active)) return false;
  if (q.StartDate && now < startOfDay_(new Date(q.StartDate))) return false;
  if (q.EndDate && now > endOfDay_(new Date(q.EndDate))) return false;
  return true;
}

/** Only http(s) links are allowed in quests and proof. */
function safeUrl_(u) {
  var s = String(u || '').trim();
  return /^https?:\/\/[^\s<>"']+$/i.test(s) ? s : '';
}

/** How many times a quest has been claimed (pending + approved), for MaxClaims limits. */
function claimsFor_(questId, subs) {
  return subs.filter(function (s) { return s.QuestID === questId && (s.Status === 'PENDING' || s.Status === 'APPROVED'); }).length;
}

/** Self-reported quests don't need a photo. */
function needsProof_(quest) { return !/self|honou?r/i.test(String(quest.Proof || '')); }

/** Pet names: letters, numbers, spaces, - and ' only, max 12 characters. */
function cleanPetName_(name) {
  var s = String(name || '').replace(/[^A-Za-z0-9 '\-]/g, '').replace(/\s+/g, ' ').trim().slice(0, 12);
  return s;
}

function randomToken_() {
  var chars = 'abcdefghijkmnpqrstuvwxyz23456789';
  var t = '';
  for (var i = 0; i < 12; i++) t += chars.charAt(Math.floor(Math.random() * chars.length));
  return t;
}

function randomPin_() { return String(100000 + Math.floor(Math.random() * 900000)); }

function driveId_(url) {
  var m = String(url).match(/[?&]id=([\w-]{10,})/) || String(url).match(/\/d\/([\w-]{10,})/);
  return m ? m[1] : '';
}

// ───────────────────────────── setup ─────────────────────────────

function setup() {
  var ss = ss_();
  clearAllCaches_();
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var head = HEADERS[name];
    if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#1b2450').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });

  // v1 kept Google Form links in this column. v2 submits inside the app, so clear them.
  var qs = sheet_(TAB.QUESTS);
  var notesCol = HEADERS.Quests.indexOf('AdminNotes') + 1;
  if (qs.getLastRow() > 1) {
    var rng = qs.getRange(2, notesCol, qs.getLastRow() - 1, 1);
    rng.setValues(rng.getValues().map(function (r) { return [/^https:\/\/docs\.google\.com\/forms/.test(String(r[0])) ? '' : r[0]]; }));
  }

  // Checkboxes
  var cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  function boxes(tab, names) {
    names.forEach(function (h) { sheet_(tab).getRange(2, HEADERS[tab].indexOf(h) + 1, 500, 1).setDataValidation(cb); });
  }
  boxes(TAB.SUBS, ['Approve', 'Reject']);
  boxes(TAB.REDEEM, ['Fulfilled', 'Refund']);
  boxes(TAB.STAFF, ['Active', 'OnLeave', 'IsAdmin']);
  boxes(TAB.QUESTS, ['Active']);
  boxes(TAB.SHOP, ['Active']);

  // Dropdowns that keep data tidy
  var statRule = SpreadsheetApp.newDataValidation().requireValueInList(STATS.map(function (s) { return s.key; }), true).build();
  sheet_(TAB.QUESTS).getRange(2, HEADERS.Quests.indexOf('Stat') + 1, 500, 1).setDataValidation(statRule);
  var repRule = SpreadsheetApp.newDataValidation().requireValueInList(REPEATS, true).build();
  sheet_(TAB.QUESTS).getRange(2, HEADERS.Quests.indexOf('Repeat') + 1, 500, 1).setDataValidation(repRule);
  var roleRule = SpreadsheetApp.newDataValidation().requireValueInList(['Tech', 'Radiologist', 'Admin'], true).build();
  sheet_(TAB.STAFF).getRange(2, HEADERS.Staff.indexOf('Role') + 1, 500, 1).setDataValidation(roleRule);

  seedIfEmpty_();
  clearAllCaches_();

  // Make sure the owner can open the Admin screen
  var owner = norm_(Session.getEffectiveUser().getEmail());
  var hasAdmin = readTable_(TAB.STAFF).some(function (s) { return truthy_(s.IsAdmin); });
  if (!hasAdmin) {
    var mine = findStaffByEmail_(owner);
    if (mine) setCell_(TAB.STAFF, mine._row, 'IsAdmin', true);
  }

  proofFolder_();
  installTriggers_();

  var def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  SpreadsheetApp.getUi().alert(
    'Rad Quest ' + VERSION + ' setup done.\n\n' +
    'Your admin PIN is ' + cfg_('ADMIN_PIN', '(none)') + ' (you can change it in the Config tab).\n\n' +
    'Next:\n' +
    '• Fill in the Staff tab.\n' +
    '• Deploy ▸ New deployment ▸ Web app (Execute as: Me, Access: Anyone). Copy the web app URL into docs/config.js in the GitHub repo.\n' +
    '• Put your GitHub Pages address in Config ▸ APP_URL.\n' +
    '• Then run Rad Quest ▸ 2. Generate personal links.');
}

function seedIfEmpty_() {
  var now = new Date();
  var me = Session.getEffectiveUser().getEmail();
  var have = {};
  readSheet_(TAB.CONFIG).forEach(function (r) { have[r.Key] = true; });
  [
    ['SEASON_NAME', 'Season 1', 'Shown in the app'],
    ['SEASON_START', new Date(now.getFullYear(), now.getMonth(), 1), 'Team leaderboard counts XP from this date'],
    ['FOOD_MAX', 10, 'Pet fullness cap'],
    ['FOOD_DECAY_HOURS', 12, 'Pet loses 1 fullness every N hours (paused while OnLeave is ticked)'],
    ['STAT_XP_PER_POINT', 60, 'XP in a stat needed for 1 stat point (max 10)'],
    ['ADMIN_EMAIL', me, 'Gets an email when someone buys a real reward'],
    ['ADMIN_PIN', randomPin_(), 'Admins type this once per session to open the Admin screen. Blank = no PIN.'],
    ['APPROVERS', me, 'Emails allowed to tick Approve in the Submissions tab (blank = anyone with edit access)'],
    ['NOTIFY_STAFF', 'TRUE', 'Email staff when a quest is approved or rejected'],
    ['DAILY_DIGEST', 'TRUE', 'Email admins each morning if submissions are waiting'],
    ['PROOF_FOLDER_ID', '', 'Filled by setup: the private Drive folder where proof photos are saved'],
    ['APP_URL', '', 'Where the app is hosted, e.g. https://yourname.github.io/rad-quest/ (needed for personal links)']
  ].forEach(function (r) { if (!have[r[0]]) appendObj_(TAB.CONFIG, { Key: r[0], Value: r[1], Notes: r[2] }); });

  if (readSheet_(TAB.STAFF).length === 0) {
    appendObj_(TAB.STAFF, {
      Name: 'Your Name', Email: me, Role: 'Tech', Team: 'Team Alpha',
      Active: true, OnLeave: false, PetName: 'Boney', IsAdmin: true
    });
  }

  if (readSheet_(TAB.QUESTS).length === 0) {
    // ID, Title, Stat, Roles, Coins, XP, Proof, Repeat, Description
    [
      ['Q01', 'Weekly QC phantom test', 'SKILL', 'Tech', 40, 60, 'Photo of QC log', 'weekly', 'Run and log the weekly QC test for your room.'],
      ['Q02', 'Shadow another modality for one shift', 'SKILL', 'Tech', 60, 90, 'Modality lead sign-off', 'monthly', 'CT, MRI, US or mammo. Write 3 things you learned.'],
      ['Q03', 'Update a technique / exposure chart', 'SKILL', 'Tech', 50, 80, 'Photo of chart', 'monthly', 'Review and update one exposure chart with your lead.'],
      ['Q04', 'Positioning self-review (10 images)', 'SKILL', 'Tech', 40, 60, 'Photo of checklist', 'monthly', 'Score 10 of your own images against positioning criteria. No patient identifiers.'],
      ['Q05', 'Learn a system feature and teach it', 'SKILL', 'ALL', 40, 60, 'Peer sign-off', 'monthly', 'PACS, RIS, HIS or modality console. Teach one colleague.'],
      ['Q06', 'Inspect 5 lead aprons', 'SHIELD', 'Tech', 30, 40, 'Photo of inspection log', 'monthly', 'Visual / tactile check, log results, flag damaged aprons.'],
      ['Q07', 'Dose audit: 10 cases', 'SHIELD', 'Tech,Radiologist', 70, 100, 'Photo of audit sheet', 'monthly', 'Compare 10 exams against reference levels. De-identified only.'],
      ['Q08', 'Repeat / reject analysis (1 week)', 'SHIELD', 'Tech', 60, 90, 'Photo of analysis', 'monthly', 'Tally rejects and reasons for your room for one week. Damages the boss.'],
      ['Q09', 'Radiation safety refresher', 'SHIELD', 'ALL', 40, 60, 'Certificate', 'once', 'Complete the department radiation safety refresher module.'],
      ['Q10', 'Dosimeter returned on time', 'SHIELD', 'ALL', 20, 30, 'RPO confirms', 'monthly', 'Badge worn and exchanged on schedule.'],
      ['Q11', 'ALARA spot check of a room', 'SHIELD', 'Tech', 30, 50, 'Photo of checklist', 'weekly', 'Shielding, collimation, signage, door interlocks.'],
      ['Q12', 'Lead a near-miss learning review', 'SHIELD', 'ALL', 60, 90, 'Lead sign-off', 'monthly', 'Blame-free discussion of a near miss and what we change. Not points per report.'],
      ['Q13', 'Summarize a journal article', 'WISDOM', 'ALL', 50, 70, 'Photo or link to summary', 'monthly', 'One-page summary shared with the team.'],
      ['Q14', 'Present a case at a team huddle', 'WISDOM', 'ALL', 60, 90, 'Lead sign-off', 'monthly', '5-minute case: technique, findings, learning point.'],
      ['Q15', 'Finish a CPD / online course', 'WISDOM', 'ALL', 80, 120, 'Certificate', 'unlimited', 'Any approved radiology-related course.'],
      ['Q16', 'Teach a student or intern', 'WISDOM', 'Tech', 40, 60, 'Mentor sign-off', 'weekly', 'Teach one positioning or technique with feedback.'],
      ['Q17', 'Teaching session for techs (15 min)', 'WISDOM', 'Radiologist', 80, 120, 'Photo of attendance list', 'weekly', 'Anatomy, pathology or image-quality topic for the tech team.'],
      ['Q18', 'Interesting case of the week', 'WISDOM', 'Radiologist', 50, 80, 'Shared slide / post', 'weekly', 'Walk the team through one case. De-identified.'],
      ['Q19', 'Write or update an SOP', 'WISDOM', 'Admin,Tech', 60, 90, 'Photo or link to document', 'monthly', 'Document one process clearly so anyone can follow it.'],
      ['Q20', 'Image-quality feedback to a tech', 'HEART', 'Radiologist', 30, 50, 'Tech confirms', 'weekly', 'Specific, kind, useful feedback on a real study.'],
      ['Q21', 'Earn a patient compliment', 'HEART', 'ALL', 30, 50, 'Feedback form / letter', 'unlimited', 'Written or verbal compliment recorded by a lead.'],
      ['Q22', 'Cover a teammate', 'HEART', 'ALL', 40, 50, 'Peer sign-off', 'unlimited', 'Cover a shift, duty or task so a colleague can rest.'],
      ['Q23', 'Help another section during a rush', 'HEART', 'ALL', 30, 40, 'Peer sign-off', 'weekly', 'Step in where the queue is longest.'],
      ['Q24', 'Welcome and orient a newcomer', 'HEART', 'ALL', 40, 60, 'Sign-off', 'unlimited', 'Give a new staff member or student a proper tour and buddy day.'],
      ['Q25', 'Front-desk improvement implemented', 'HEART', 'Admin', 60, 90, 'Lead sign-off', 'monthly', 'One change that improves patient waiting, signage or communication.'],
      ['Q26', 'Booking accuracy check', 'SKILL', 'Admin', 40, 60, 'Photo of check', 'weekly', 'Audit the week\'s bookings for errors and fix them.'],
      ['Q27', 'Organize a team activity', 'HEART', 'ALL', 50, 70, 'Photo', 'monthly', 'Birthday, potluck, outing, or celebration.'],
      ['Q28', '50,000 steps in a week', 'VITAL', 'ALL', 25, 40, 'Self-report', 'weekly', 'Screenshot optional. Honour system.'],
      ['Q29', 'Full lunch break 5 days in a row', 'VITAL', 'ALL', 20, 30, 'Self-report', 'weekly', 'Actually take your break.'],
      ['Q30', 'Join a team exercise session', 'VITAL', 'ALL', 30, 40, 'Photo', 'weekly', 'Walk, badminton, zumba, anything together.'],
      ['Q31', 'Safe patient-handling stretch routine', 'VITAL', 'ALL', 30, 40, 'Peer sign-off', 'once', 'Learn and demo a stretch routine for lifting and transfers.'],
      ['Q32', 'Rest and recharge evening', 'VITAL', 'ALL', 20, 30, 'Self-report', 'weekly', 'A screen-free evening, a hobby, time with family.']
    ].forEach(function (q) {
      appendObj_(TAB.QUESTS, {
        QuestID: q[0], Title: q[1], Stat: q[2], Roles: q[3], Coins: q[4], XP: q[5], Proof: q[6],
        Repeat: q[7], Active: true, Description: q[8]
      });
    });
  }

  if (readSheet_(TAB.SHOP).length === 0) {
    [
      ['R01', 'RPO training seat', 'REAL', 1500, '', 'Radiation Protection Officer course from the training budget. Admin confirms the schedule.'],
      ['R02', 'Short course: CT or MRI', 'REAL', 1200, '', 'Registration for one approved short course.'],
      ['R03', 'Conference registration', 'REAL', 900, '', 'One local radiology conference of your choice.'],
      ['R04', 'Online CPD course', 'REAL', 400, '', 'One approved online course.'],
      ['R05', 'Lunch with the Chief', 'REAL', 200, '', 'Department treat and a career chat.'],
      ['R06', 'First pick of shifts (1 month)', 'REAL', 150, '', 'Choose your preferred schedule first.'],
      ['R07', 'Raffle ticket', 'REAL', 20, '', 'Monthly draw for a surprise prize.'],
      ['P01', 'Rad Snack', 'PET', 20, 'food:2', '+2 fullness.'],
      ['P02', 'Contrast Smoothie', 'PET', 50, 'food:6', '+6 fullness.'],
      ['P03', 'Lead apron cape', 'PET', 120, 'cosmetic:cape', 'Cosmetic. Extra protected.'],
      ['P04', 'Dosimeter badge hat', 'PET', 80, 'cosmetic:hat', 'Cosmetic. Safety first.'],
      ['P05', 'Room: CT gantry', 'PET', 150, 'room:ct', 'New background for your pet.'],
      ['P06', 'Room: Night shift', 'PET', 100, 'room:night', 'Starry background for night owls.']
    ].forEach(function (s) {
      appendObj_(TAB.SHOP, { ItemID: s[0], Name: s[1], Category: s[2], Price: s[3], Effect: s[4], Description: s[5], Active: true });
    });
  }

  if (readSheet_(TAB.BOSS).length === 0) {
    appendObj_(TAB.BOSS, {
      BossID: 'B01', Name: 'The Reject Rate Dragon', MaxHP: 3000, HP: 3000, Stats: 'SHIELD,SKILL', RewardCoins: 100,
      StartDate: new Date(now.getFullYear(), now.getMonth(), 1), EndDate: new Date(now.getFullYear(), now.getMonth() + 1, 0),
      Status: 'ACTIVE',
      Description: 'Approved Shield and Skill quests deal damage equal to their XP. Defeat it and everyone gets the reward.'
    });
  }
}

function installTriggers_() {
  var mine = { onEditHandler: 1, onChangeHandler: 1, dailyDigest: 1, onFormSubmitHandler: 1, dailyRefresh: 1 };
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (mine[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss_()).onEdit().create();
  ScriptApp.newTrigger('onChangeHandler').forSpreadsheet(ss_()).onChange().create();
  ScriptApp.newTrigger('dailyDigest').timeBased().everyDays(1).atHour(7).create();
}

/** Private Drive folder for proof photos. Only the Sheet owner can open it; admins see photos through the app. */
function proofFolder_() {
  var id = cfg_('PROOF_FOLDER_ID', '');
  if (id) { try { return DriveApp.getFolderById(String(id)); } catch (err) { } }
  var f = DriveApp.createFolder('Rad Quest proofs');
  setCfg_('PROOF_FOLDER_ID', f.getId(), 'Filled by setup: the private Drive folder where proof photos are saved');
  return f;
}

// ───────────────────────────── quest status ─────────────────────────────

/** 'open' | 'pending' | 'done' for this person + quest in the quest's current repeat period. */
function questStatusFor_(email, quest, now, subs, tz) {
  var key = periodKey_(quest.Repeat, now, tz);
  var pending = false, done = false;
  for (var i = 0; i < subs.length; i++) {
    var s = subs[i];
    if (s.QuestID !== quest.QuestID || norm_(s.Email) !== email) continue;
    if (s.Status === 'PENDING') pending = true;
    if (s.Status === 'APPROVED' && key !== null && periodKey_(quest.Repeat, new Date(s.Timestamp), tz) === key) done = true;
  }
  if (done) return 'done';
  if (pending) return 'pending';
  return 'open';
}

// ───────────────────────────── submitting a quest (from the app) ─────────────────────────────

/**
 * p = { questId, what, link, photos: [dataUrl...], noPatientInfo: true }
 */
function submitQuest(token, p) {
  var me = requireStaff_(token);
  var email = norm_(me.Email);
  p = p || {};
  var what = String(p.what || '').trim().slice(0, 600);
  var link = String(p.link || '').trim();
  if (!what) throw new Error('Tell us in a sentence what you did.');
  if (link && !safeUrl_(link)) throw new Error('The link must start with http:// or https://');
  if (!p.noPatientInfo) throw new Error('Please confirm your proof has no patient details.');
  var photos = [].concat(p.photos || []);
  if (photos.length > MAX_PHOTOS) throw new Error('Up to ' + MAX_PHOTOS + ' photos, please.');
  var blobs = photos.map(function (d, i) {
    var m = String(d).match(/^data:image\/(jpeg|png);base64,([A-Za-z0-9+\/=]+)$/);
    if (!m) throw new Error('Photo ' + (i + 1) + ' could not be read. Try taking it again.');
    if (m[2].length > MAX_PHOTO_CHARS) throw new Error('Photo ' + (i + 1) + ' is too large.');
    return Utilities.newBlob(Utilities.base64Decode(m[2]), 'image/' + m[1], 'proof.' + (m[1] === 'png' ? 'png' : 'jpg'));
  });

  return withLock_(function () {
    var now = new Date();
    var quest = findQuest_(String(p.questId || ''));
    if (!quest || !questIsLive_(quest, now)) throw new Error('That quest is not open right now.');
    if (!questFitsRole_(quest.Roles, me.Role)) throw new Error('That quest is not for your role.');
    var subs = readTable_(TAB.SUBS);
    var st = questStatusFor_(email, quest, now, subs, tz_());
    if (st === 'pending') throw new Error('You already sent this one. It is waiting for review.');
    if (st === 'done') throw new Error('You already earned this quest for now. It will reopen later.');
    if (num_(quest.MaxClaims) > 0 && claimsFor_(quest.QuestID, subs) >= num_(quest.MaxClaims)) throw new Error('This quest is full.');

    var evidence = [];
    if (link) evidence.push(link);
    if (blobs.length) {
      var folder = proofFolder_();
      var stamp = Utilities.formatDate(now, tz_(), 'yyyy-MM-dd');
      blobs.forEach(function (b, i) {
        b.setName(stamp + ' ' + quest.QuestID + ' ' + me.Name + ' ' + (i + 1) + (b.getContentType() === 'image/png' ? '.png' : '.jpg'));
        evidence.push('https://drive.google.com/open?id=' + folder.createFile(b).getId());
      });
    }
    var flag = (!evidence.length && needsProof_(quest)) ? 'NO PROOF ATTACHED' : '';
    appendObj_(TAB.SUBS, {
      Timestamp: now, Email: email, Name: me.Name, QuestID: quest.QuestID, QuestTitle: quest.Title, WhatIDid: what,
      Evidence: evidence.join('\n'), Flag: flag, Approve: false, Reject: false, Status: 'PENDING'
    });
    return { message: 'Sent! An admin will check it soon.', state: buildState_(me) };
  });
}

// ───────────────────────────── sheet triggers ─────────────────────────────

/** Installable onEdit: keeps caches fresh, and lets approvers tick boxes in the Sheet as a backup to the app. */
function onEditHandler(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  var name = sh.getName();
  touch_(name);
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  var row = e.range.getRow();
  if (row < 2) return;
  var header = sh.getRange(1, e.range.getColumn()).getValue();
  if (name === TAB.STAFF && header === 'OnLeave') { leaveChanged_(row, truthy_(e.range.getValue())); return; }
  if (!truthy_(e.range.getValue())) return;
  var who = (e.user && e.user.getEmail && e.user.getEmail()) || 'owner';

  if (name === TAB.SUBS && (header === 'Approve' || header === 'Reject')) {
    if (!isApprover_(who)) { e.range.setValue(false); SpreadsheetApp.getActive().toast('Only approvers listed in Config can approve.'); return; }
    withLock_(function () {
      try { reviewSubmission_(row, header === 'Approve', who); }
      catch (err) { SpreadsheetApp.getActive().toast(err.message); }
    });
  }
  if (name === TAB.REDEEM && (header === 'Fulfilled' || header === 'Refund')) {
    if (!isApprover_(who)) { e.range.setValue(false); return; }
    withLock_(function () { handleRedemption_(row, header, who); });
  }
}

/**
 * OnLeave ticked: remember how full the pet is right now, so it stays that way.
 * OnLeave unticked: hunger starts counting again from this moment.
 */
function leaveChanged_(staffRow, onLeave) {
  withLock_(function () {
    var s = readSheet_(TAB.STAFF).filter(function (x) { return x._row === staffRow; })[0];
    if (!s || !s.Email) return;
    var pet = getPet_(norm_(s.Email));
    var now = new Date();
    if (onLeave) setCells_(TAB.PETS, pet._row, { Food: foodNow_(pet, { OnLeave: false }, now), FoodUpdatedAt: now });
    else setCell_(TAB.PETS, pet._row, 'FoodUpdatedAt', now);
  });
}

/** Rows inserted or deleted by hand: drop the caches so the app sees it. */
function onChangeHandler() { clearAllCaches_(); }

function isApprover_(who) {
  var list = String(cfg_('APPROVERS', '')).split(',').map(norm_).filter(String);
  if (!list.length || who === 'owner') return true;
  return list.indexOf(norm_(who)) >= 0;
}

/**
 * Approve or reject one submission. Used by the Admin screen and by the Sheet checkboxes.
 * opts (optional): { note, coins, xp } — coins/xp override the quest's default award (partial credit or a bonus).
 */
function reviewSubmission_(row, approve, who, opts) {
  opts = opts || {};
  var sub = readTable_(TAB.SUBS).filter(function (s) { return s._row === row; })[0];
  if (!sub) throw new Error('Submission not found.');
  if (sub.Status !== 'PENDING') {
    // already handled: put the checkboxes back the way they were
    setCells_(TAB.SUBS, row, { Approve: sub.Status === 'APPROVED', Reject: sub.Status === 'REJECTED' });
    throw new Error('Already ' + String(sub.Status).toLowerCase() + '.');
  }
  var now = new Date();
  var quest = findQuest_(sub.QuestID);
  var note = (opts.note !== undefined && opts.note !== null && String(opts.note).trim() !== '')
    ? String(opts.note).trim().slice(0, 300) : String(sub.ReviewNote || '');
  if (approve && !quest) {
    setCells_(TAB.SUBS, row, { Approve: false, ReviewNote: 'Cannot approve: quest ID not found' });
    throw new Error('Quest ID not found.');
  }
  var coins = 0, xp = 0;
  function given(v) { return v !== undefined && v !== null && String(v).trim() !== ''; }
  if (approve) {
    coins = given(opts.coins) ? Math.max(0, Math.round(num_(opts.coins))) : num_(quest.Coins);
    xp = given(opts.xp) ? Math.max(0, Math.round(num_(opts.xp))) : num_(quest.XP);
    appendObj_(TAB.LEDGER, {
      Timestamp: now, Email: norm_(sub.Email), Type: 'EARN', Ref: quest.QuestID,
      Coins: coins, XP: xp, Stat: quest.Stat, Note: quest.Title
    });
    damageBoss_(quest.Stat, xp, now);
  }
  setCells_(TAB.SUBS, row, {
    Approve: !!approve, Reject: !approve, Status: approve ? 'APPROVED' : 'REJECTED',
    ReviewedBy: who, ReviewedAt: now, ReviewNote: note,
    CoinsAwarded: approve ? coins : '', XPAwarded: approve ? xp : ''
  });

  if (truthy_(cfg_('NOTIFY_STAFF', 'TRUE')) && sub.Email) {
    try {
      var title = (quest && quest.Title) || sub.QuestTitle || sub.QuestID;
      MailApp.sendEmail(String(sub.Email),
        'Rad Quest: ' + (approve ? 'quest approved!' : 'quest needs another look'),
        approve
          ? 'Nice work! "' + title + '" was approved: +' + coins + ' coins, +' + xp + ' XP.' + (note ? '\nNote: ' + note : '') + '\nOpen Rad Quest to feed your pet.'
          : 'Your submission for "' + title + '" was not approved this time.' + (note ? '\nNote: ' + note : '') + '\nYou can fix it and send it again from the app.');
    } catch (err) { /* email quota or bad address: ignore */ }
  }
}

function damageBoss_(stat, xp, now) {
  readTable_(TAB.BOSS).forEach(function (b) {
    if (b.Status !== 'ACTIVE') return;
    if (b.StartDate && now < startOfDay_(new Date(b.StartDate))) return;
    if (b.EndDate && now > endOfDay_(new Date(b.EndDate))) return;
    var stats = String(b.Stats || 'ALL').split(',').map(function (s) { return s.trim().toUpperCase(); });
    if (stats.indexOf('ALL') < 0 && stats.indexOf(String(stat).toUpperCase()) < 0) return;
    var hp = Math.max(0, num_(b.HP) - num_(xp));
    if (hp > 0) { setCell_(TAB.BOSS, b._row, 'HP', hp); return; }
    setCells_(TAB.BOSS, b._row, { HP: 0, Status: 'DEFEATED' });
    readTable_(TAB.STAFF).filter(function (s) { return truthy_(s.Active) && s.Email; }).forEach(function (s) {
      appendObj_(TAB.LEDGER, {
        Timestamp: now, Email: norm_(s.Email), Type: 'BONUS', Ref: b.BossID,
        Coins: num_(b.RewardCoins), XP: 0, Stat: '', Note: 'Team defeated ' + b.Name
      });
    });
  });
}

function handleRedemption_(row, header, who) {
  var r = readTable_(TAB.REDEEM).filter(function (x) { return x._row === row; })[0];
  if (!r) return;
  if (r.Status !== 'REQUESTED') { setCell_(TAB.REDEEM, row, header, r.Status === (header === 'Refund' ? 'REFUNDED' : 'FULFILLED')); return; }
  var now = new Date();
  if (header === 'Refund') {
    appendObj_(TAB.LEDGER, {
      Timestamp: now, Email: norm_(r.Email), Type: 'REFUND', Ref: r.ItemID,
      Coins: num_(r.Price), XP: 0, Stat: '', Note: 'Refund: ' + r.ItemName
    });
  }
  setCells_(TAB.REDEEM, row, { Status: header === 'Refund' ? 'REFUNDED' : 'FULFILLED', HandledBy: who, HandledAt: now });
}

// ───────────────────────────── lookups ─────────────────────────────

function findStaffByEmail_(email) {
  email = norm_(email);
  return readTable_(TAB.STAFF).filter(function (s) { return norm_(s.Email) === email; })[0] || null;
}

function findStaffByToken_(token) {
  token = String(token || '').trim();
  if (!token) return null;
  return readTable_(TAB.STAFF).filter(function (s) { return String(s.Token) === token && truthy_(s.Active); })[0] || null;
}

function findQuest_(id) {
  return readTable_(TAB.QUESTS).filter(function (q) { return q.QuestID === id; })[0] || null;
}

/** Coins, XP and per-stat XP for one person. `until` (optional Date) only counts rows up to that moment. */
function totalsFor_(email, ledger, until) {
  var t = { coins: 0, xp: 0, stat: {} };
  STATS.forEach(function (s) { t.stat[s.key] = 0; });
  ledger.forEach(function (l) {
    if (norm_(l.Email) !== email) return;
    if (until && new Date(l.Timestamp) > until) return;
    t.coins += num_(l.Coins);
    t.xp += Math.max(0, num_(l.XP));
    if (l.Stat && t.stat[l.Stat] !== undefined) t.stat[l.Stat] += Math.max(0, num_(l.XP));
  });
  return t;
}

function statVals_(totals) {
  var per = num_(cfg_('STAT_XP_PER_POINT', 60)) || 60;
  return STATS.map(function (s) { return statPoints_(totals.stat[s.key], per); });
}

function getPet_(email) {
  var p = readTable_(TAB.PETS).filter(function (x) { return norm_(x.Email) === email; })[0];
  if (!p) {
    var rec = { Email: email, Food: 5, FoodUpdatedAt: new Date(), Cosmetics: '', LastPetAt: '', Room: '', LastSeenAt: '', Hatched: false };
    rec._row = appendObj_(TAB.PETS, rec);
    p = rec;
  }
  return p;
}

function foodNow_(pet, me, now) {
  return currentFood_(num_(pet.Food), pet.FoodUpdatedAt ? new Date(pet.FoodUpdatedAt).getTime() : 0,
    now.getTime(), num_(cfg_('FOOD_DECAY_HOURS', 12)) || 12, truthy_(me.OnLeave));
}

// ───────────────────────────── web app ─────────────────────────────

/** Opening the /exec URL in a browser just confirms the server is alive and points to the app. */
function doGet() {
  var app = safeUrl_(cfg_('APP_URL', ''));
  var html = '<body style="font-family:sans-serif;padding:24px"><h3>Rad Quest server ' + VERSION + ' is running.</h3>' +
    (app ? '<p>Open the app: <a target="_top" href="' + app + '">' + app + '</a> (use your personal link).</p>'
         : '<p>Set APP_URL in the Config tab to your GitHub Pages address.</p>') + '</body>';
  return HtmlService.createHtmlOutput(html).setTitle('Rad Quest server');
}

/** The only actions the app shell may call. Everything checks the caller's personal token itself. */
function api_() {
  return {
    open: openApp_, getState: getState, feedPet: feedPet, buyItem: buyItem, petPet: petPet, namePet: namePet,
    submitQuest: submitQuest, adminLogin: adminLogin, adminGetQueue: adminGetQueue, adminGetProof: adminGetProof,
    adminReview: adminReview, adminGetQuests: adminGetQuests, adminToggleQuest: adminToggleQuest, adminSaveQuest: adminSaveQuest
  };
}

/**
 * The app shell sends { fn: 'feedPet', args: [token, ...] } as the POST body.
 * Replies with { ok: true, value } or { ok: false, error }.
 */
function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var table = api_();
    if (!Object.prototype.hasOwnProperty.call(table, String(req.fn))) throw new Error('Unknown action.');
    var args = Object.prototype.toString.call(req.args) === '[object Array]' ? req.args : [];
    out = { ok: true, value: table[req.fn].apply(null, args), version: VERSION };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err), version: VERSION };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function requireStaff_(token) {
  var s = findStaffByToken_(token);
  if (!s) throw new Error('This link is not active. Ask the admin for your personal Rad Quest link.');
  return s;
}

/** First load: the state plus what happened since this person last opened the app. */
function openApp_(token) {
  var me = requireStaff_(token);
  var email = norm_(me.Email);
  var now = new Date();
  var pet = getPet_(email);
  var lastSeen = pet.LastSeenAt ? new Date(pet.LastSeenAt) : null;
  var state = buildState_(me);
  var news = [];
  var levelUp = null, evolved = null;

  if (lastSeen) {
    var ledger = readTable_(TAB.LEDGER);
    ledger.forEach(function (l) {
      if (norm_(l.Email) !== email || new Date(l.Timestamp) <= lastSeen) return;
      if (l.Type === 'EARN') news.push({ kind: 'good', text: 'Quest approved: ' + (l.Note || l.Ref), coins: num_(l.Coins), xp: num_(l.XP) });
      else if (l.Type === 'BONUS') news.push({ kind: 'good', text: l.Note || 'Bonus!', coins: num_(l.Coins), xp: num_(l.XP) });
      else if (l.Type === 'REFUND') news.push({ kind: 'good', text: l.Note || 'Refund', coins: num_(l.Coins), xp: 0 });
      else if (l.Type === 'ADJUST') news.push({ kind: 'info', text: l.Note || 'Balance adjusted by admin', coins: num_(l.Coins), xp: num_(l.XP) });
    });
    readTable_(TAB.SUBS).forEach(function (s) {
      if (norm_(s.Email) !== email || s.Status !== 'REJECTED' || !s.ReviewedAt || new Date(s.ReviewedAt) <= lastSeen) return;
      news.push({ kind: 'bad', text: 'Needs another look: ' + (s.QuestTitle || s.QuestID) + (s.ReviewNote ? ' — ' + s.ReviewNote : ''), coins: 0, xp: 0 });
    });
    readTable_(TAB.REDEEM).forEach(function (r) {
      if (norm_(r.Email) !== email || r.Status !== 'FULFILLED' || !r.HandledAt || new Date(r.HandledAt) <= lastSeen) return;
      news.push({ kind: 'good', text: 'Reward confirmed: ' + r.ItemName, coins: 0, xp: 0 });
    });
    var before = totalsFor_(email, ledger, lastSeen);
    var lvBefore = levelFromXp_(before.xp);
    if (state.level > lvBefore) levelUp = { from: lvBefore, to: state.level };
    var formBefore = petForm_(lvBefore, statVals_(before));
    if (formBefore !== state.pet.form) evolved = { from: formBefore, to: state.pet.form };
    var hungerLost = num_(pet.Food) - state.pet.food;
    if (hungerLost >= 3 && !state.me.onLeave) news.push({ kind: 'info', text: state.me.petName + ' got hungry while you were away.', coins: 0, xp: 0 });
  }
  setCell_(TAB.PETS, pet._row, 'LastSeenAt', now);
  return { state: state, news: news.slice(-12), levelUp: levelUp, evolved: evolved, hatch: !truthy_(pet.Hatched) };
}

/** Refresh everything for one person. */
function getState(token) { return buildState_(requireStaff_(token)); }

function buildState_(me) {
  var email = norm_(me.Email);
  var now = new Date();
  var tz = tz_();
  var ledger = readTable_(TAB.LEDGER);
  var subs = readTable_(TAB.SUBS);
  var totals = totalsFor_(email, ledger);
  var level = levelFromXp_(totals.xp);
  var vals = statVals_(totals);
  var stats = STATS.map(function (s, i) {
    return { key: s.key, name: s.name, color: s.color, about: s.about, val: vals[i], xp: totals.stat[s.key] };
  });

  // pet
  var pet = getPet_(email);
  var max = num_(cfg_('FOOD_MAX', 10)) || 10;
  var onLeave = truthy_(me.OnLeave);
  if (onLeave && pet.FoodUpdatedAt && now - new Date(pet.FoodUpdatedAt) > 3600000) {
    setCell_(TAB.PETS, pet._row, 'FoodUpdatedAt', now); // freeze hunger while on leave
  }
  var food = Math.min(max, foodNow_(pet, me, now));
  var day = function (d) { return Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd'); };
  var pettedToday = !!pet.LastPetAt && day(pet.LastPetAt) === day(now);

  // quests
  var claimCount = {};
  subs.forEach(function (s) { if (s.Status === 'PENDING' || s.Status === 'APPROVED') claimCount[s.QuestID] = (claimCount[s.QuestID] || 0) + 1; });
  var quests = readTable_(TAB.QUESTS)
    .filter(function (q) { return q.QuestID && questIsLive_(q, now) && questFitsRole_(q.Roles, me.Role); })
    .map(function (q) {
      var status = questStatusFor_(email, q, now, subs, tz);
      var cap = num_(q.MaxClaims), used = claimCount[q.QuestID] || 0;
      if (status === 'open' && cap > 0 && used >= cap) status = 'full';
      return {
        id: q.QuestID, title: q.Title, stat: q.Stat, coins: num_(q.Coins), xp: num_(q.XP), proof: q.Proof,
        repeat: q.Repeat, description: q.Description, infoLink: safeUrl_(q.Link), status: status,
        needsProof: needsProof_(q),
        endsOn: q.EndDate ? Utilities.formatDate(new Date(q.EndDate), tz, 'MMM d') : '',
        spotsLeft: cap > 0 ? Math.max(0, cap - used) : null
      };
    });

  // my recent submissions
  var mySubs = subs.filter(function (s) { return norm_(s.Email) === email; }).slice(-6).reverse().map(function (s) {
    return {
      title: s.QuestTitle || s.QuestID, status: s.Status, note: s.ReviewNote || '',
      when: s.Timestamp ? Utilities.formatDate(new Date(s.Timestamp), tz, 'MMM d') : '',
      coins: (s.CoinsAwarded === '' || s.CoinsAwarded == null) ? null : num_(s.CoinsAwarded),
      xp: (s.XPAwarded === '' || s.XPAwarded == null) ? null : num_(s.XPAwarded)
    };
  });

  // shop
  var owned = String(pet.Cosmetics || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  var shop = readTable_(TAB.SHOP).filter(function (i) { return i.ItemID && truthy_(i.Active); }).map(function (i) {
    var eff = parseEffect_(i.Effect);
    return {
      id: i.ItemID, name: i.Name, category: String(i.Category).toUpperCase(), price: num_(i.Price),
      description: i.Description, effect: eff.kind, amount: eff.kind === 'food' ? (num_(eff.value) || 1) : 0,
      owned: (eff.kind === 'cosmetic' && owned.indexOf(eff.value) >= 0) || (eff.kind === 'room' && pet.Room === eff.value)
    };
  });

  var myRequests = readTable_(TAB.REDEEM).filter(function (r) { return norm_(r.Email) === email; }).slice(-5).map(function (r) {
    return { name: r.ItemName, status: r.Status, when: Utilities.formatDate(new Date(r.Timestamp), tz, 'MMM d') };
  });

  // boss
  var boss = readTable_(TAB.BOSS).filter(function (b) { return b.Status === 'ACTIVE' || b.Status === 'DEFEATED'; }).slice(-1)[0];
  var bossOut = null;
  if (boss) {
    bossOut = {
      name: boss.Name, hp: num_(boss.HP), maxHp: num_(boss.MaxHP) || 1, status: boss.Status, stats: boss.Stats,
      reward: num_(boss.RewardCoins), description: boss.Description,
      daysLeft: boss.EndDate ? Math.max(0, Math.ceil((endOfDay_(new Date(boss.EndDate)) - now) / 86400000)) : null
    };
  }

  // teams (season XP)
  var seasonStart = cfg_('SEASON_START', '') ? new Date(cfg_('SEASON_START')) : new Date(0);
  var teamOf = {}, teams = {};
  readTable_(TAB.STAFF).filter(function (s) { return truthy_(s.Active) && s.Email; }).forEach(function (s) {
    var t = s.Team || 'No team';
    teamOf[norm_(s.Email)] = t;
    if (!teams[t]) teams[t] = { team: t, xp: 0, members: 0, quests: 0 };
    teams[t].members++;
  });
  ledger.forEach(function (l) {
    if (l.Type !== 'EARN' || new Date(l.Timestamp) < seasonStart) return;
    var t = teamOf[norm_(l.Email)];
    if (!t) return;
    teams[t].xp += num_(l.XP);
    teams[t].quests++;
  });
  var teamList = Object.keys(teams).map(function (k) { return teams[k]; })
    .sort(function (a, b) { return (b.xp / b.members) - (a.xp / a.members); });

  var isAdmin = truthy_(me.IsAdmin);
  return {
    me: {
      name: me.Name, role: me.Role, team: me.Team || 'No team', petName: me.PetName || 'Boney', onLeave: onLeave,
      isAdmin: isAdmin, pinRequired: isAdmin && String(cfg_('ADMIN_PIN', '')).trim() !== ''
    },
    pendingReviews: isAdmin ? subs.filter(function (s) { return s.Status === 'PENDING'; }).length : 0,
    season: cfg_('SEASON_NAME', 'Season 1'),
    coins: totals.coins, xp: totals.xp, level: level,
    levelFloor: xpForLevel_(level), levelNext: xpForLevel_(level + 1),
    stats: stats,
    pet: {
      food: food, max: max, mood: onLeave ? 'On holiday. Hunger paused.' : moodFor_(food, max),
      form: petForm_(level, vals), cosmetics: owned, room: pet.Room || '', pettedToday: pettedToday
    },
    quests: quests, mySubs: mySubs, shop: shop, requests: myRequests, boss: bossOut, teams: teamList,
    maxPhotos: MAX_PHOTOS
  };
}

/** Feed = buy the cheapest active food item. */
function feedPet(token) {
  requireStaff_(token);
  var foods = readTable_(TAB.SHOP).filter(function (i) { return truthy_(i.Active) && parseEffect_(i.Effect).kind === 'food'; })
    .sort(function (a, b) { return num_(a.Price) - num_(b.Price); });
  if (!foods.length) throw new Error('No food in the shop yet.');
  return buyItem(token, foods[0].ItemID);
}

function buyItem(token, itemId) {
  requireStaff_(token);
  return withLock_(function () {
    var me = requireStaff_(token);
    var email = norm_(me.Email);
    var item = readTable_(TAB.SHOP).filter(function (i) { return i.ItemID === itemId && truthy_(i.Active); })[0];
    if (!item) throw new Error('That item is not available.');
    var price = num_(item.Price);
    var coins = totalsFor_(email, readTable_(TAB.LEDGER)).coins;
    if (coins < price) throw new Error('Not enough coins. You need ' + (price - coins) + ' more.');

    var eff = parseEffect_(item.Effect);
    var now = new Date();
    var pet = getPet_(email);
    var petName = me.PetName || 'Your pet';
    var msg = '';

    if (eff.kind === 'food') {
      var max = num_(cfg_('FOOD_MAX', 10)) || 10;
      var food = foodNow_(pet, me, now);
      if (food >= max) throw new Error(petName + ' is full! Try a quest instead.');
      setCells_(TAB.PETS, pet._row, { Food: Math.min(max, food + (num_(eff.value) || 1)), FoodUpdatedAt: now });
      msg = 'NOM NOM! ' + petName + ' ate a ' + item.Name + '.';
    } else if (eff.kind === 'cosmetic') {
      var owned = String(pet.Cosmetics || '').split(',').map(function (s) { return s.trim(); }).filter(String);
      if (owned.indexOf(eff.value) >= 0) throw new Error('You already own that.');
      owned.push(eff.value);
      setCell_(TAB.PETS, pet._row, 'Cosmetics', owned.join(','));
      msg = petName + ' is wearing the ' + item.Name + '!';
    } else if (eff.kind === 'room') {
      if (pet.Room === eff.value) throw new Error('That room is already on.');
      setCell_(TAB.PETS, pet._row, 'Room', eff.value);
      msg = 'New room unlocked: ' + item.Name + '.';
    } else {
      appendObj_(TAB.REDEEM, {
        Timestamp: now, Email: email, Name: me.Name, ItemID: item.ItemID, ItemName: item.Name, Price: price,
        Fulfilled: false, Refund: false, Status: 'REQUESTED'
      });
      var admin = cfg_('ADMIN_EMAIL', '');
      if (admin) {
        try { MailApp.sendEmail(String(admin), 'Rad Quest reward request: ' + item.Name, me.Name + ' spent ' + price + ' coins on "' + item.Name + '". See the Redemptions tab.'); } catch (err) { }
      }
      msg = 'Request sent! The admin will confirm "' + item.Name + '".';
    }

    appendObj_(TAB.LEDGER, {
      Timestamp: now, Email: email, Type: 'SPEND', Ref: item.ItemID, Coins: -price, XP: 0, Stat: '', Note: item.Name
    });
    return { message: msg, state: buildState_(me) };
  });
}

function petPet(token) {
  var me = requireStaff_(token);
  var email = norm_(me.Email);
  var pet = getPet_(email);
  var tz = tz_();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var name = me.PetName || 'Your pet';
  if (pet.LastPetAt && Utilities.formatDate(new Date(pet.LastPetAt), tz, 'yyyy-MM-dd') === today) {
    return { message: name + ' already got pets today. Come back tomorrow!', state: buildState_(me) };
  }
  setCell_(TAB.PETS, pet._row, 'LastPetAt', new Date());
  return { message: name + ' wiggles happily!', state: buildState_(me) };
}

/** Name the pet (also finishes the hatching screen on first open). */
function namePet(token, name) {
  var clean = cleanPetName_(name);
  if (clean.length < 2) throw new Error('Pick a name with at least 2 letters.');
  return withLock_(function () {
    var me = requireStaff_(token);
    setCell_(TAB.STAFF, me._row, 'PetName', clean);
    var pet = getPet_(norm_(me.Email));
    setCell_(TAB.PETS, pet._row, 'Hatched', true);
    me = requireStaff_(token);
    return { message: clean + ' says hi!', state: buildState_(me) };
  });
}

// ───────────────────────────── Admin screen API ─────────────────────────────
// Needs IsAdmin ticked in the Staff tab AND the admin PIN from the Config tab.

function requireAdmin_(token, pin) {
  var s = requireStaff_(token);
  if (!truthy_(s.IsAdmin)) throw new Error('Admins only.');
  var real = String(cfg_('ADMIN_PIN', '')).trim();
  if (!real) return s;
  var c = cache_();
  var key = 'pinfail_' + s.Token;
  var fails = c ? Number(c.get(key) || 0) : 0;
  if (fails >= 5) throw new Error('PIN locked: too many wrong tries. Wait 15 minutes.');
  if (String(pin === undefined || pin === null ? '' : pin).trim() !== real) {
    if (c) c.put(key, String(fails + 1), 900);
    throw new Error('PIN wrong. Check the Config tab or ask the owner.');
  }
  if (c && fails) c.remove(key);
  return s;
}

function adminLogin(token, pin) {
  requireAdmin_(token, pin);
  return adminGetQueue(token, pin, 'PENDING');
}

function fmtDay_(d, tz) { return d ? Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd') : ''; }

function parseDay_(s) {
  var m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : '';
}

/** view: 'PENDING' (default) or 'DONE' (last 40 reviewed). */
function adminGetQueue(token, pin, view) {
  requireAdmin_(token, pin);
  var tz = tz_();
  var quests = {};
  readTable_(TAB.QUESTS).forEach(function (q) { quests[q.QuestID] = q; });
  var subs = readTable_(TAB.SUBS);
  var pendingCount = subs.filter(function (s) { return s.Status === 'PENDING'; }).length;
  var list = view === 'DONE'
    ? subs.filter(function (s) { return s.Status === 'APPROVED' || s.Status === 'REJECTED'; }).slice(-40).reverse()
    : subs.filter(function (s) { return s.Status === 'PENDING'; });
  return {
    pendingCount: pendingCount,
    items: list.map(function (s) {
      var q = quests[s.QuestID] || {};
      return {
        row: s._row, when: s.Timestamp ? Utilities.formatDate(new Date(s.Timestamp), tz, 'MMM d, h:mm a') : '',
        name: s.Name || s.Email, questId: s.QuestID, questTitle: s.QuestTitle || q.Title || '',
        stat: q.Stat || '', proofNeeded: q.Proof || '', what: s.WhatIDid,
        evidence: String(s.Evidence || '').split('\n').map(function (u) { return u.trim(); }).filter(String).map(function (u) {
          return { url: safeUrl_(u), driveId: driveId_(u) };
        }).filter(function (e) { return e.url; }),
        flag: s.Flag, status: s.Status, defaultCoins: num_(q.Coins), defaultXp: num_(q.XP),
        coinsAwarded: s.CoinsAwarded, xpAwarded: s.XPAwarded, note: s.ReviewNote, reviewedBy: s.ReviewedBy
      };
    })
  };
}

/** Send one proof photo to an admin's screen. Only files that are attached to a submission can be read. */
function adminGetProof(token, pin, fileId) {
  requireAdmin_(token, pin);
  fileId = String(fileId || '');
  if (!/^[\w-]{10,}$/.test(fileId)) throw new Error('Proof not found.');
  var known = readTable_(TAB.SUBS).some(function (s) { return String(s.Evidence || '').indexOf(fileId) >= 0; });
  if (!known) throw new Error('Proof not found.');
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var type = String(blob.getContentType() || '');
  if (!/^image\/(jpeg|png|gif|webp)$/.test(type) || file.getSize() > 4000000) return { id: fileId, image: '', name: file.getName() };
  return { id: fileId, image: 'data:' + type + ';base64,' + Utilities.base64Encode(blob.getBytes()), name: file.getName() };
}

function adminReview(token, pin, row, approve, note, coins, xp) {
  var me = requireAdmin_(token, pin);
  withLock_(function () { reviewSubmission_(Number(row), !!approve, norm_(me.Email), { note: note, coins: coins, xp: xp }); });
  return adminGetQueue(token, pin, 'PENDING');
}

function adminGetQuests(token, pin) {
  requireAdmin_(token, pin);
  var tz = tz_();
  var now = new Date();
  var subs = readTable_(TAB.SUBS);
  return readTable_(TAB.QUESTS).filter(function (q) { return q.QuestID; }).map(function (q) {
    return {
      id: q.QuestID, title: q.Title, stat: q.Stat, roles: q.Roles, coins: num_(q.Coins), xp: num_(q.XP),
      proof: q.Proof, repeat: q.Repeat, active: truthy_(q.Active), live: questIsLive_(q, now),
      description: q.Description, link: q.Link, startDate: fmtDay_(q.StartDate, tz), endDate: fmtDay_(q.EndDate, tz),
      maxClaims: num_(q.MaxClaims) || '', claims: claimsFor_(q.QuestID, subs),
      pending: subs.filter(function (s) { return s.QuestID === q.QuestID && s.Status === 'PENDING'; }).length
    };
  });
}

function adminToggleQuest(token, pin, id, active) {
  requireAdmin_(token, pin);
  withLock_(function () {
    var q = findQuest_(id);
    if (!q) throw new Error('Quest not found.');
    setCell_(TAB.QUESTS, q._row, 'Active', !!active);
  });
  return { quests: adminGetQuests(token, pin), message: active ? 'Quest turned ON.' : 'Quest turned OFF.' };
}

/** Create (no id) or update a quest from the Admin screen. */
function adminSaveQuest(token, pin, q) {
  requireAdmin_(token, pin);
  q = q || {};
  var title = String(q.title || '').trim().slice(0, 80);
  if (!title) throw new Error('Give the quest a title.');
  var stat = String(q.stat || '').toUpperCase();
  if (!STATS.some(function (s) { return s.key === stat; })) throw new Error('Pick a stat.');
  var repeat = String(q.repeat || 'unlimited').toLowerCase();
  if (REPEATS.indexOf(repeat) < 0) throw new Error('Pick how often it can be earned.');
  var link = String(q.link || '').trim();
  if (link && !safeUrl_(link)) throw new Error('The link must start with http:// or https://');
  var start = parseDay_(q.startDate), end = parseDay_(q.endDate);
  if (start && end && end < start) throw new Error('End date is before start date.');
  var roles = String(q.roles || 'ALL').split(',').map(function (r) { return r.trim(); }).filter(String).join(',') || 'ALL';
  var rec = {
    Title: title, Stat: stat, Roles: roles,
    Coins: Math.max(0, Math.round(num_(q.coins))), XP: Math.max(0, Math.round(num_(q.xp))),
    Proof: String(q.proof || '').trim().slice(0, 80), Repeat: repeat, Active: !!q.active,
    Description: String(q.description || '').trim().slice(0, 600),
    Link: link, StartDate: start, EndDate: end, MaxClaims: Math.max(0, Math.round(num_(q.maxClaims))) || ''
  };
  var id = withLock_(function () {
    var existing = q.id ? findQuest_(String(q.id)) : null;
    if (existing) { setCells_(TAB.QUESTS, existing._row, rec); return existing.QuestID; }
    var max = 0;
    readTable_(TAB.QUESTS).forEach(function (x) { var n = Number(String(x.QuestID).replace(/\D/g, '')); if (n > max) max = n; });
    rec.QuestID = 'Q' + (max + 1 < 10 ? '0' : '') + (max + 1);
    appendObj_(TAB.QUESTS, rec);
    return rec.QuestID;
  });
  return { quests: adminGetQuests(token, pin), message: 'Saved ' + id + '.', id: id };
}

// ───────────────────────────── scheduled + menu tools ─────────────────────────────

/** Every morning: tell admins if submissions are waiting. */
function dailyDigest() {
  if (!truthy_(cfg_('DAILY_DIGEST', 'TRUE'))) return;
  var pending = readTable_(TAB.SUBS).filter(function (s) { return s.Status === 'PENDING'; });
  if (!pending.length) return;
  var lines = pending.slice(0, 15).map(function (s) { return '• ' + (s.Name || s.Email) + ': ' + (s.QuestTitle || s.QuestID); });
  readTable_(TAB.STAFF).filter(function (s) { return truthy_(s.Active) && truthy_(s.IsAdmin) && s.Email; }).forEach(function (s) {
    try {
      MailApp.sendEmail(String(s.Email), 'Rad Quest: ' + pending.length + ' submission' + (pending.length === 1 ? '' : 's') + ' waiting',
        lines.join('\n') + (pending.length > 15 ? '\n…and ' + (pending.length - 15) + ' more' : '') +
        '\n\nOpen your Rad Quest link and tap ADMIN to review.' + (s.PersonalLink ? '\n' + s.PersonalLink : ''));
    } catch (err) { }
  });
}

function generatePersonalLinks() {
  clearAllCaches_();
  var url = safeUrl_(cfg_('APP_URL', '')).split('?')[0].split('#')[0];
  if (!url) {
    SpreadsheetApp.getUi().alert('First put your GitHub Pages address in Config ▸ APP_URL, e.g. https://yourname.github.io/rad-quest/');
    return;
  }
  var used = {};
  var staff = readSheet_(TAB.STAFF);
  staff.forEach(function (s) { if (s.Token) used[s.Token] = true; });
  staff.forEach(function (s) {
    if (!s.Email) return;
    var tok = s.Token;
    if (!tok) { do { tok = randomToken_(); } while (used[tok]); used[tok] = true; }
    setCells_(TAB.STAFF, s._row, { Token: tok, PersonalLink: url + '?u=' + tok });
  });
  SpreadsheetApp.getUi().alert('Personal links are in the Staff tab. Treat each link like a password.\n\nThe 12-letter code at the end of a link also works on the app\'s sign-in screen.');
}

function emailPersonalLinks() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Email every active staff member their personal Rad Quest link?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var n = 0;
  readSheet_(TAB.STAFF).forEach(function (s) {
    if (!truthy_(s.Active) || !s.Email || !s.PersonalLink) return;
    MailApp.sendEmail(String(s.Email), 'Your Rad Quest link',
      'Hi ' + s.Name + ',\n\nAn egg is waiting for you in Rad Quest!\n\n' +
      'Your personal link (keep it private):\n' + s.PersonalLink + '\n\n' +
      'Open it on your phone, then choose "Add to Home Screen" so Rad Quest opens like an app.\n\n' +
      'Pick a quest, do it, then tap SUBMIT in the app and add a photo as proof.\n' +
      'Please never include patient names, IDs, faces or PACS screenshots.\n\nHave fun!');
    n++;
  });
  ui.alert('Sent ' + n + ' emails.');
}

function showBalances() {
  var ledger = readSheet_(TAB.LEDGER);
  var lines = readSheet_(TAB.STAFF).filter(function (s) { return s.Email; }).map(function (s) {
    var t = totalsFor_(norm_(s.Email), ledger);
    return s.Name + ': ' + t.coins + ' coins, ' + t.xp + ' XP (Lv ' + levelFromXp_(t.xp) + ')';
  });
  SpreadsheetApp.getUi().alert(lines.join('\n') || 'No staff yet.');
}

function showAdminPin() {
  clearAllCaches_();
  SpreadsheetApp.getUi().alert('Admin PIN: ' + cfg_('ADMIN_PIN', '(none set)') + '\n\nChange it in the Config tab.');
}
