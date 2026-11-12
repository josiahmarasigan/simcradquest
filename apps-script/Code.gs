/**
 * SIMC RAD QUEST v3 — gamified staff development
 * Google Apps Script backend (bound to one Google Sheet).
 *
 * The screens live in App.html, in this Apps Script project. Deploying a new version updates everyone.
 * GitHub Pages (/docs) only hosts a tiny loader, set up once: it asks this script for App.html
 * (and the player's data) on every open, keeps a copy for offline, and makes the app installable.
 * Every action is a POST to this web app's /exec URL.
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

var VERSION = '3.2';

// ───────────────────────────── constants ─────────────────────────────

var TAB = {
  CONFIG: 'Config', STAFF: 'Staff', QUESTS: 'Quests', SUBS: 'Submissions', LEDGER: 'Ledger',
  SHOP: 'Shop', REDEEM: 'Redemptions', PETS: 'Pets', BOSS: 'Boss',
  QUIZZES: 'Quizzes', QUESTIONS: 'Questions', ATTEMPTS: 'QuizAttempts'
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
  Boss: ['BossID', 'Name', 'MaxHP', 'HP', 'Stats', 'RewardCoins', 'StartDate', 'EndDate', 'Status', 'Description'],
  Quizzes: ['QuizID', 'Title', 'Category', 'Stat', 'Roles', 'Coins', 'XP', 'Questions', 'PassPercent', 'Repeat', 'Active', 'Description'],
  Questions: ['QuizID', 'QID', 'Question', 'A', 'B', 'C', 'D', 'Answer', 'Explanation', 'Active', 'Source'],
  QuizAttempts: ['Timestamp', 'Email', 'Name', 'QuizID', 'QuizTitle', 'Score', 'Total', 'Percent', 'Passed', 'CoinsAwarded', 'XPAwarded', 'Seconds', 'Wrong']
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
var CACHED_TABS = { Config: 1, Staff: 1, Quests: 1, Shop: 1, Boss: 1, Quizzes: 1, Questions: 1 };
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
  boxes(TAB.QUIZZES, ['Active']);
  boxes(TAB.QUESTIONS, ['Active']);

  // Dropdowns that keep data tidy
  var statRule = SpreadsheetApp.newDataValidation().requireValueInList(STATS.map(function (s) { return s.key; }), true).build();
  sheet_(TAB.QUESTS).getRange(2, HEADERS.Quests.indexOf('Stat') + 1, 500, 1).setDataValidation(statRule);
  sheet_(TAB.QUIZZES).getRange(2, HEADERS.Quizzes.indexOf('Stat') + 1, 200, 1).setDataValidation(statRule);
  var ansRule = SpreadsheetApp.newDataValidation().requireValueInList(['A', 'B', 'C', 'D'], true).build();
  sheet_(TAB.QUESTIONS).getRange(2, HEADERS.Questions.indexOf('Answer') + 1, 500, 1).setDataValidation(ansRule);
  var repRule = SpreadsheetApp.newDataValidation().requireValueInList(REPEATS, true).build();
  sheet_(TAB.QUESTS).getRange(2, HEADERS.Quests.indexOf('Repeat') + 1, 500, 1).setDataValidation(repRule);
  sheet_(TAB.QUIZZES).getRange(2, HEADERS.Quizzes.indexOf('Repeat') + 1, 200, 1).setDataValidation(repRule);
  var roleRule = SpreadsheetApp.newDataValidation().requireValueInList(['Tech', 'Radiologist', 'Admin'], true).build();
  sheet_(TAB.STAFF).getRange(2, HEADERS.Staff.indexOf('Role') + 1, 500, 1).setDataValidation(roleRule);

  seedIfEmpty_();
  seedQuizzes_();
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
    ['QUIZ_TRIES_PER_DAY', 3, 'How many times a day each person may attempt the same quiz'],
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

/** The app's screens (App.html) and a short fingerprint, so phones only download them when they change. */
function appUi_() {
  var html = HtmlService.createHtmlOutputFromFile('App').getContent();
  var v = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, html)).slice(0, 12);
  return { v: v, html: html };
}

/** First call from the GitHub loader: the screens (only if the phone's copy is out of date) + the player's data. */
function boot(token, knownVersion) {
  var ui = appUi_();
  var out = { v: ui.v, ui: String(knownVersion || '') === ui.v ? null : ui, data: null, error: '' };
  if (token) {
    try { out.data = openApp_(token); } catch (err) { out.error = String(err && err.message || err); }
  }
  return out;
}

/** The only actions the app shell may call. Everything checks the caller's personal token itself. */
function api_() {
  return {
    boot: boot, open: openApp_, getState: getState, feedPet: feedPet, buyItem: buyItem, petPet: petPet, namePet: namePet,
    submitQuest: submitQuest, adminLogin: adminLogin, adminGetQueue: adminGetQueue, adminGetProof: adminGetProof,
    adminReview: adminReview, adminGetQuests: adminGetQuests, adminToggleQuest: adminToggleQuest, adminSaveQuest: adminSaveQuest,
    startQuiz: startQuiz, submitQuiz: submitQuiz, adminGetQuizzes: adminGetQuizzes, adminToggleQuiz: adminToggleQuiz
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
    if ((l.Type !== 'EARN' && l.Type !== 'QUIZ') || new Date(l.Timestamp) < seasonStart) return;
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
    quests: quests, quizzes: quizzesForState_(me, now, tz), mySubs: mySubs, shop: shop, requests: myRequests, boss: bossOut, teams: teamList,
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

// ───────────────────────────── quizzes ─────────────────────────────
// Answers never leave the server. startQuiz sends questions with shuffled options; submitQuiz grades them.
// Passing a quiz pays coins + XP straight away (once per repeat period). No admin review needed.

var QUIZ_SECONDS = 45 * 60; // an unfinished quiz expires after 45 minutes

function quizIsLive_(z) { return z && z.QuizID && truthy_(z.Active); }

function findQuiz_(id) {
  return readTable_(TAB.QUIZZES).filter(function (z) { return z.QuizID === id; })[0] || null;
}

function quizPool_(quizId) {
  return readTable_(TAB.QUESTIONS).filter(function (q) {
    return q.QuizID === quizId && truthy_(q.Active) && String(q.Question || '').trim() && /^[ABCD]$/i.test(String(q.Answer || '').trim());
  });
}

function quizOptions_(q) {
  return ['A', 'B', 'C', 'D'].filter(function (L) { return String(q[L] === undefined || q[L] === null ? '' : q[L]).trim() !== ''; });
}

function shuffle_(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) { var j = Math.floor(Math.random() * (i + 1)); var t = a[i]; a[i] = a[j]; a[j] = t; }
  return a;
}

/** { status: 'open' | 'done' | 'limit', triesLeft, best } for one person and quiz. */
function quizStatusFor_(email, z, now, attempts, tz) {
  var key = periodKey_(z.Repeat, now, tz);
  var today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  var limit = Math.max(1, num_(cfg_('QUIZ_TRIES_PER_DAY', 3)) || 3);
  var done = false, triesToday = 0, best = null;
  attempts.forEach(function (a) {
    if (a.QuizID !== z.QuizID || norm_(a.Email) !== email) return;
    var when = new Date(a.Timestamp);
    if (Utilities.formatDate(when, tz, 'yyyy-MM-dd') === today) triesToday++;
    if (key !== null && periodKey_(z.Repeat, when, tz) === key) {
      if (truthy_(a.Passed)) done = true;
      best = best === null ? num_(a.Percent) : Math.max(best, num_(a.Percent));
    }
  });
  var left = Math.max(0, limit - triesToday);
  return { status: done ? 'done' : (left === 0 ? 'limit' : 'open'), triesLeft: left, best: best };
}

function quizzesForState_(me, now, tz) {
  var email = norm_(me.Email);
  var attempts = readTable_(TAB.ATTEMPTS);
  return readTable_(TAB.QUIZZES)
    .filter(function (z) { return quizIsLive_(z) && questFitsRole_(z.Roles, me.Role); })
    .map(function (z) {
      var pool = quizPool_(z.QuizID).length;
      if (!pool) return null;
      var st = quizStatusFor_(email, z, now, attempts, tz);
      return {
        id: z.QuizID, title: z.Title, category: z.Category, stat: z.Stat, coins: num_(z.Coins), xp: num_(z.XP),
        count: Math.min(pool, Math.max(1, num_(z.Questions) || 10)), pass: num_(z.PassPercent) || 80,
        repeat: z.Repeat, description: z.Description, status: st.status, triesLeft: st.triesLeft, best: st.best
      };
    })
    .filter(Boolean);
}

function startQuiz(token, quizId) {
  var me = requireStaff_(token);
  var email = norm_(me.Email);
  var now = new Date();
  var z = findQuiz_(String(quizId || ''));
  if (!quizIsLive_(z)) throw new Error('That quiz is not open right now.');
  if (!questFitsRole_(z.Roles, me.Role)) throw new Error('That quiz is not for your role.');
  var st = quizStatusFor_(email, z, now, readTable_(TAB.ATTEMPTS), tz_());
  if (st.status === 'done') throw new Error('You already passed this quiz. It reopens later.');
  if (st.status === 'limit') throw new Error('No tries left today for this quiz. Come back tomorrow!');
  var pool = quizPool_(z.QuizID);
  if (!pool.length) throw new Error('This quiz has no questions yet.');
  var n = Math.min(pool.length, Math.max(1, num_(z.Questions) || 10));
  var picked = shuffle_(pool).slice(0, n).map(function (q) { return { q: q, order: shuffle_(quizOptions_(q)) }; });
  var id = Utilities.getUuid ? Utilities.getUuid() : randomToken_() + randomToken_();
  var record = { e: email, z: z.QuizID, t: now.getTime(), items: picked.map(function (p) { return { id: String(p.q.QID), o: p.order }; }) };
  var c = cache_();
  if (!c) throw new Error('Quizzes are unavailable right now. Try again in a minute.');
  c.put('quiz_' + id, JSON.stringify(record), QUIZ_SECONDS);
  return {
    attemptId: id, quizId: z.QuizID, title: z.Title, pass: num_(z.PassPercent) || 80, coins: num_(z.Coins), xp: num_(z.XP),
    questions: picked.map(function (p) { return { text: String(p.q.Question), options: p.order.map(function (L) { return String(p.q[L]); }) }; })
  };
}

/** answers: for each question, the index of the option the player picked (in the order it was shown), or -1. */
function submitQuiz(token, attemptId, answers) {
  var me = requireStaff_(token);
  var email = norm_(me.Email);
  var c = cache_();
  var key = 'quiz_' + String(attemptId || '');
  var raw = c ? c.get(key) : null;
  if (!raw) throw new Error('This quiz has expired. Start it again from the Quests screen.');
  var rec = JSON.parse(raw);
  if (rec.e !== email) throw new Error('This quiz belongs to someone else.');
  c.remove(key); // one submission per attempt
  answers = Object.prototype.toString.call(answers) === '[object Array]' ? answers : [];

  var byId = {};
  readTable_(TAB.QUESTIONS).forEach(function (q) { if (q.QuizID === rec.z) byId[String(q.QID)] = q; });
  var score = 0, total = 0, wrong = [];
  var review = rec.items.map(function (it, i) {
    var q = byId[it.id];
    if (!q) return null; // question deleted meanwhile: not counted
    total++;
    var pick = Math.round(num_(answers[i]));
    var chosen = (pick >= 0 && pick < it.o.length) ? it.o[pick] : '';
    var right = String(q.Answer).trim().toUpperCase();
    var ok = chosen === right;
    if (ok) score++; else wrong.push(it.id);
    return {
      text: String(q.Question), options: it.o.map(function (L) { return String(q[L]); }),
      chosen: chosen ? pick : -1, correct: it.o.indexOf(right), ok: ok, explanation: String(q.Explanation || '')
    };
  }).filter(Boolean);
  if (!total) throw new Error('This quiz changed while you were taking it. Please start again.');

  return withLock_(function () {
    var now = new Date();
    var z = findQuiz_(rec.z);
    var title = z ? z.Title : rec.z;
    var pass = z ? (num_(z.PassPercent) || 80) : 80;
    var percent = Math.round(score / total * 100);
    var passed = percent >= pass;
    var st = z ? quizStatusFor_(email, z, now, readTable_(TAB.ATTEMPTS), tz_()) : { status: 'open' };
    var paid = null;
    if (passed && z && st.status !== 'done') {
      paid = { coins: num_(z.Coins), xp: num_(z.XP) };
      appendObj_(TAB.LEDGER, {
        Timestamp: now, Email: email, Type: 'QUIZ', Ref: z.QuizID, Coins: paid.coins, XP: paid.xp, Stat: z.Stat, Note: 'Quiz passed: ' + title
      });
      damageBoss_(z.Stat, paid.xp, now);
    }
    appendObj_(TAB.ATTEMPTS, {
      Timestamp: now, Email: email, Name: me.Name, QuizID: rec.z, QuizTitle: title, Score: score, Total: total, Percent: percent,
      Passed: passed, CoinsAwarded: paid ? paid.coins : 0, XPAwarded: paid ? paid.xp : 0,
      Seconds: Math.round((now.getTime() - rec.t) / 1000), Wrong: wrong.join(', ')
    });
    return {
      title: title, score: score, total: total, percent: percent, pass: pass, passed: passed, paid: paid,
      alreadyEarned: passed && !paid, review: review, state: buildState_(me)
    };
  });
}

function adminGetQuizzes(token, pin) {
  requireAdmin_(token, pin);
  var attempts = readTable_(TAB.ATTEMPTS);
  return readTable_(TAB.QUIZZES).filter(function (z) { return z.QuizID; }).map(function (z) {
    var mine = attempts.filter(function (a) { return a.QuizID === z.QuizID; });
    var passes = mine.filter(function (a) { return truthy_(a.Passed); });
    var people = {};
    passes.forEach(function (a) { people[norm_(a.Email)] = 1; });
    return {
      id: z.QuizID, title: z.Title, category: z.Category, stat: z.Stat, roles: z.Roles, active: truthy_(z.Active),
      pool: quizPool_(z.QuizID).length, perAttempt: num_(z.Questions) || 10, pass: num_(z.PassPercent) || 80,
      coins: num_(z.Coins), xp: num_(z.XP), attempts: mine.length,
      passRate: mine.length ? Math.round(passes.length / mine.length * 100) : null, passedBy: Object.keys(people).length
    };
  });
}

function adminToggleQuiz(token, pin, id, active) {
  requireAdmin_(token, pin);
  withLock_(function () {
    var z = findQuiz_(String(id || ''));
    if (!z) throw new Error('Quiz not found.');
    setCell_(TAB.QUIZZES, z._row, 'Active', !!active);
  });
  return { quizzes: adminGetQuizzes(token, pin), message: active ? 'Quiz turned ON.' : 'Quiz turned OFF.' };
}

/** Bulk-add rows to a tab in one write (used to seed the question bank). */
function appendRows_(name, objs) {
  if (!objs.length) return;
  var sh = sheet_(name);
  var head = HEADERS[name];
  var start = lastDataRow_(sh) + 1;
  var need = start + objs.length - 1;
  if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 50);
  sh.getRange(start, 1, objs.length, head.length).setValues(objs.map(function (o) {
    return head.map(function (h) { return o[h] === undefined ? '' : o[h]; });
  }));
  touch_(name);
}

function seedQuizzes_() {
  var bank = quizBank_();
  if (readSheet_(TAB.QUIZZES).length === 0) {
    appendRows_(TAB.QUIZZES, bank.quizzes.map(function (z) {
      return { QuizID: z[0], Title: z[1], Category: z[2], Stat: z[3], Roles: z[4], Coins: z[5], XP: z[6], Questions: z[7], PassPercent: z[8], Repeat: z[9], Active: true, Description: z[10] };
    }));
  }
  if (readSheet_(TAB.QUESTIONS).length === 0) {
    var counter = {};
    appendRows_(TAB.QUESTIONS, bank.questions.map(function (q) {
      counter[q[0]] = (counter[q[0]] || 0) + 1;
      var n = counter[q[0]];
      var row = { QuizID: q[0], QID: q[0] + '-' + (n < 10 ? '0' : '') + n, Question: q[1], Answer: 'ABCD'.charAt(q[3]), Explanation: q[4], Active: true, Source: q[5] || '' };
      q[2].forEach(function (opt, i) { row['ABCD'.charAt(i)] = opt; });
      return row;
    }));
  }
}

// ───────────────────────────── starter quiz bank ─────────────────────────────
// Used once by setup to fill the Quizzes and Questions tabs. After that, edit questions in the Sheet.
// Each question: [QuizID, question, [options], index of the correct option (0 = first), explanation]

function quizBank_() {
  var P = 'SIMC-RAD-P&P-001 Draft 0';
  var quizzes = [
    // ID, Title, Category, Stat, Roles, Coins, XP, Questions per attempt, Pass %, Repeat, Description
    ['QZ-XRT', 'X-ray: technique & positioning', 'X-ray technical', 'SKILL', 'Tech', 30, 50, 8, 75, 'weekly', 'Exposure factors, grids, markers, decubitus views and digital exposure indicators.'],
    ['QZ-XRK', 'X-ray: image knowledge', 'X-ray knowledge', 'WISDOM', 'Tech,Radiologist', 30, 50, 8, 75, 'weekly', 'Anatomy, views and common findings on plain films.'],
    ['QZ-CTT', 'CT: protocols & technique', 'CT technical', 'SKILL', 'Tech', 30, 50, 8, 75, 'weekly', 'Pitch, windows, contrast timing, artifacts and patient centering.'],
    ['QZ-CTK', 'CT: image knowledge', 'CT knowledge', 'WISDOM', 'Tech,Radiologist', 30, 50, 8, 75, 'weekly', 'Hounsfield units, head CT, stroke, PE, appendicitis and contrast reactions.'],
    ['QZ-UST', 'Ultrasound: scanning technique', 'Ultrasound technical', 'SKILL', 'Tech', 30, 50, 8, 75, 'weekly', 'Transducers, gain, focus, Doppler, preparation and artifacts.'],
    ['QZ-USK', 'Ultrasound: image knowledge', 'Ultrasound knowledge', 'WISDOM', 'Tech,Radiologist', 30, 50, 8, 75, 'weekly', 'Cysts, stones, the biliary tree, FAST, aorta, DVT and early pregnancy.'],
    ['QZ-RP', 'Radiation protection', 'Radiation protection', 'SHIELD', 'ALL', 35, 60, 8, 75, 'weekly', 'Principles, dose limits, dosimeters, PPE and pregnancy.'],
    ['QZ-PHY', 'Radiologic physics', 'Physics', 'WISDOM', 'Tech,Radiologist', 30, 50, 8, 75, 'weekly', 'How X-rays, CT numbers and ultrasound work.'],
    ['QZ-POL', 'SIMC policies & procedures', 'Policies', 'HEART', 'ALL', 40, 60, 10, 80, 'weekly', 'Our department P&P (Draft 0): requests, ID, contrast, results, reprints, incidents.']
  ];

  var q = [
    // ── X-ray technical ──
    ['QZ-XRT', 'Raising kVp while keeping everything else the same will mainly:', ['Increase beam penetration and lower image contrast', 'Decrease penetration and raise contrast', 'Change only the number of photons, not their energy', 'Make the focal spot smaller'], 0, 'kVp controls beam energy. Higher kVp penetrates more and gives a longer grey scale, so contrast falls.'],
    ['QZ-XRT', 'Increasing kVp by about 15% has roughly the same effect on receptor exposure as:', ['Doubling the mAs', 'Halving the mAs', 'Doubling the SID', 'Adding a grid'], 0, 'This is the 15% rule: +15% kVp ≈ 2× mAs for receptor exposure (but with less contrast).'],
    ['QZ-XRT', 'You change the SID from 100 cm to 200 cm. To keep the same receptor exposure, the mAs should be:', ['Multiplied by 4', 'Doubled', 'Halved', 'Left the same'], 0, 'Inverse square law: doubling distance quarters intensity, so mAs must be multiplied by (200/100)² = 4.'],
    ['QZ-XRT', 'The standard SID for an erect PA chest is about:', ['180 cm', '100 cm', '120 cm', '60 cm'], 0, 'About 180 cm (72 inches) reduces magnification of the heart and gives sharper images.'],
    ['QZ-XRT', 'Why is an erect chest X-ray taken PA rather than AP when the patient can stand?', ['The heart is closer to the detector, so it is magnified less', 'It gives a higher dose to the lungs', 'It removes the need for a marker', 'It shows the spine better'], 0, 'In PA the heart sits next to the detector, so the cardiac shadow is not falsely enlarged.'],
    ['QZ-XRT', 'Good inspiration on a PA chest is usually shown by seeing about how many posterior ribs above the diaphragm?', ['10', '6', '8', '14'], 0, 'At least 10 posterior ribs above the hemidiaphragm suggests adequate inspiration.'],
    ['QZ-XRT', 'A grid is generally used when the body part is thicker than about:', ['10 cm', '2 cm', '5 cm', '25 cm'], 0, 'Above roughly 10 cm, scatter increases enough that a grid improves contrast.'],
    ['QZ-XRT', 'The best way to reduce motion blur is to:', ['Use the shortest practical exposure time', 'Use a longer exposure time with lower mA', 'Increase the SID', 'Remove the grid'], 0, 'Short exposure times (higher mA, shorter seconds) freeze motion; clear breathing instructions help too.'],
    ['QZ-XRT', 'Tight collimation to the area of interest:', ['Lowers patient dose and scatter and improves contrast', 'Increases patient dose', 'Has no effect on image quality', 'Makes the exposure indicator meaningless'], 0, 'A smaller field irradiates less tissue, creating less scatter and better contrast.'],
    ['QZ-XRT', 'The correct way to mark the side on a radiograph is:', ['Place a lead marker inside the collimated field before exposure', 'Type the side on the image afterwards only', 'Write it on the request form', 'Rely on the patient to say which side'], 0, 'A lead marker exposed with the image is the reliable record of side. Post-exposure annotation alone can be wrong.'],
    ['QZ-XRT', 'A small LEFT pleural effusion is suspected. Which decubitus view shows it best?', ['Left lateral decubitus (left side down)', 'Right lateral decubitus (right side down)', 'Supine AP', 'Lordotic view'], 0, 'Fluid layers on the dependent side, so the affected side goes down. For a pneumothorax, the affected side goes up.'],
    ['QZ-XRT', 'In digital radiography, a deviation index (DI) of 0 means:', ['Receptor exposure matched the target for that exam', 'No X-rays reached the detector', 'The image was overexposed by double', 'The grid was missing'], 0, 'DI compares actual to target exposure. 0 is on target; about +3 means roughly double the target exposure.'],
    ['QZ-XRT', 'Because of the anode heel effect, the thicker part of the anatomy should be placed under the:', ['Cathode side', 'Anode side', 'Centre of the grid only', 'Collimator light'], 0, 'Beam intensity is higher on the cathode side, so place the thicker part there.'],

    // ── X-ray knowledge ──
    ['QZ-XRK', 'Free air under the diaphragm is best shown on:', ['An erect chest X-ray, after the patient has been upright for several minutes', 'A supine abdomen only', 'A lateral skull view', 'A PA hand view'], 0, 'Air rises. With the patient upright for about 5–10 minutes, free gas collects under the hemidiaphragms.'],
    ['QZ-XRK', 'A suspected scaphoid fracture needs which extra view?', ['PA wrist in ulnar deviation', 'AP forearm only', 'Lateral elbow', 'Skyline knee'], 0, 'Ulnar deviation elongates the scaphoid so a waist fracture is easier to see.'],
    ['QZ-XRK', 'A Colles fracture is a fracture of the:', ['Distal radius with dorsal angulation', 'Proximal ulna', 'Fifth metacarpal neck', 'Scaphoid waist'], 0, 'Classic fall-on-outstretched-hand injury: distal radius, dorsally angled fragment.'],
    ['QZ-XRK', 'Loss of the right heart border on a PA chest suggests disease in the:', ['Right middle lobe', 'Right lower lobe', 'Left upper lobe', 'Right apex'], 0, 'Silhouette sign: the right middle lobe touches the right heart border.'],
    ['QZ-XRK', 'On a PA chest, cardiomegaly is suggested when the cardiothoracic ratio is more than:', ['0.5', '0.3', '0.7', '1.0'], 0, 'A ratio above 0.5 on a good PA film suggests an enlarged heart. AP films overestimate it.'],
    ['QZ-XRK', 'In an adult, the tip of an endotracheal tube should ideally be about:', ['5 cm above the carina', 'At the carina', 'In the right main bronchus', '15 cm above the carina'], 0, 'About 5 ± 2 cm above the carina with the head neutral.'],
    ['QZ-XRK', 'Kerley B lines on a chest X-ray usually indicate:', ['Interstitial pulmonary oedema', 'Pneumothorax', 'Rib fracture', 'Normal lung'], 0, 'Short horizontal lines at the lung bases from fluid-filled interlobular septa.'],
    ['QZ-XRK', 'The occipitomental (Waters) view is mainly used for the:', ['Maxillary sinuses', 'Odontoid peg', 'Lumbar spine', 'Knee joint'], 0, 'It projects the petrous ridges below the maxillary sinuses.'],
    ['QZ-XRK', 'The odontoid peg (C2) is best shown on which view?', ['Open-mouth AP', 'Lateral lumbar', 'PA chest', 'Townes'], 0, 'The open mouth clears the teeth from the view of C1–C2.'],
    ['QZ-XRK', 'Which system classifies fractures through the growth plate in children?', ['Salter–Harris', 'Glasgow', 'Fleischner', 'BI-RADS'], 0, 'Salter–Harris types I–V describe physeal fractures.'],
    ['QZ-XRK', 'By the 3-6-9 rule, the colon (excluding caecum) is abnormally dilated above:', ['6 cm', '3 cm', '9 cm', '12 cm'], 0, 'Small bowel > 3 cm, colon > 6 cm, caecum > 9 cm.'],
    ['QZ-XRK', 'Multiple dilated small-bowel loops with air-fluid levels on an erect abdomen suggest:', ['Small-bowel obstruction', 'Normal bowel gas', 'Pneumothorax', 'Kidney stone'], 0, 'Dilated loops with stepped air-fluid levels are typical of obstruction.'],

    // ── CT technical ──
    ['QZ-CTT', 'In helical CT, pitch is:', ['Table travel per rotation divided by total beam collimation', 'Tube current times rotation time', 'Slice thickness divided by FOV', 'Window width divided by level'], 0, 'Pitch = table feed per rotation ÷ beam collimation. Higher pitch (other factors equal) lowers dose.'],
    ['QZ-CTT', 'The CT number (HU) of water is:', ['0', '−1000', '+1000', '+100'], 0, 'HU are defined relative to water (0). Air is about −1000.'],
    ['QZ-CTT', 'Which window setting best shows lung parenchyma?', ['W 1500 / L −600', 'W 80 / L 40', 'W 400 / L 40', 'W 2000 / L 500'], 0, 'A wide window centred at low HU shows lung detail. W80/L40 is a brain window.'],
    ['QZ-CTT', 'A typical portal venous phase starts about how long after the start of IV contrast injection?', ['60–70 seconds', '10 seconds', '25 seconds', '10 minutes'], 0, 'About 60–70 s for the liver portal venous phase. Arterial phases are earlier, delayed phases later.'],
    ['QZ-CTT', 'Bolus tracking works by:', ['Watching enhancement in a region of interest and starting the scan when a threshold is reached', 'Scanning at a fixed delay for every patient', 'Injecting contrast after the scan', 'Measuring the patient\'s weight'], 0, 'A monitoring scan over a vessel triggers the diagnostic scan at the right moment.'],
    ['QZ-CTT', 'Dark streaks between two dense structures (e.g. the petrous bones) are mainly due to:', ['Beam hardening', 'Patient motion', 'Aliasing', 'Ring artifact'], 0, 'Low-energy photons are filtered out by dense bone, causing dark bands between them.'],
    ['QZ-CTT', 'Partial volume artifact can be reduced by:', ['Using thinner slices', 'Using a wider window', 'Lowering the kVp only', 'Increasing the pitch'], 0, 'Thinner slices average less tissue in each voxel.'],
    ['QZ-CTT', 'Why should the patient be centred in the gantry isocentre?', ['Bowtie filters and tube current modulation work as designed; off-centring raises dose and noise', 'It makes the scan faster', 'It is only for comfort', 'It removes the need for a scout'], 0, 'Off-centre patients get mis-shaped filtration and modulation, giving more dose and noise.'],
    ['QZ-CTT', 'The scout (topogram) is used to:', ['Plan the scan range and set dose modulation', 'Replace the diagnostic scan', 'Measure renal function', 'Check the injector'], 0, 'The scout defines coverage and feeds automatic exposure control.'],
    ['QZ-CTT', 'A contrast extravasation is noticed during injection. The first action is to:', ['Stop the injection and notify the radiologist, then follow the extravasation protocol', 'Increase the injection rate', 'Continue the scan and tell no one', 'Massage the arm vigorously'], 0, 'Stop, assess, elevate the limb, notify and document as per protocol.'],
    ['QZ-CTT', 'Before IV iodinated contrast, which lab result is most relevant?', ['Serum creatinine / eGFR', 'Blood sugar only', 'Haemoglobin only', 'Platelet count only'], 0, 'Renal function helps identify patients at higher risk from contrast.'],
    ['QZ-CTT', 'For children, CT should use:', ['Size- or weight-based paediatric protocols', 'The adult protocol to save time', 'The highest available mAs', 'Double-pass scanning routinely'], 0, 'Children are more radiosensitive; Image Gently principles call for child-sized settings.'],
    ['QZ-CTT', 'Which helps most to reduce breathing motion in a chest CT?', ['Clear breath-hold instructions and a short scan time', 'A longer rotation time', 'A thicker slice only', 'Lower kVp only'], 0, 'Explain, practise the breath-hold and keep the acquisition short.'],

    // ── CT knowledge ──
    ['QZ-CTK', 'Acute blood on a non-contrast head CT appears:', ['Hyperdense (bright)', 'Hypodense (dark)', 'The same as CSF', 'Black like air'], 0, 'Fresh clotted blood is denser than brain, roughly 50–80 HU.'],
    ['QZ-CTK', 'A lens-shaped (biconvex) collection that does not cross suture lines is typical of:', ['Extradural (epidural) haematoma', 'Subdural haematoma', 'Subarachnoid haemorrhage', 'Hydrocephalus'], 0, 'Epidural blood is limited by sutures. Subdural blood is crescent-shaped and crosses sutures.'],
    ['QZ-CTK', 'The first imaging test for suspected acute stroke is usually:', ['Non-contrast CT head', 'Abdominal ultrasound', 'Chest X-ray', 'CT abdomen with contrast'], 0, 'It quickly excludes haemorrhage before treatment decisions.'],
    ['QZ-CTK', 'A bright middle cerebral artery on non-contrast CT (hyperdense MCA sign) suggests:', ['Thrombus in the MCA', 'Normal calcification only', 'Brain abscess', 'Sinusitis'], 0, 'Clot in the vessel looks denser than flowing blood.'],
    ['QZ-CTK', 'The usual CT test for suspected pulmonary embolism is:', ['CT pulmonary angiography', 'Non-contrast CT head', 'CT KUB', 'Low-dose CT sinus'], 0, 'Contrast timed to the pulmonary arteries shows emboli as filling defects.'],
    ['QZ-CTK', 'A structure measuring about −80 HU most likely contains:', ['Fat', 'Calcium', 'Fresh blood', 'Metal'], 0, 'Fat is roughly −50 to −150 HU.'],
    ['QZ-CTK', 'On CT, findings that support acute appendicitis include:', ['Appendix over 6 mm with wall thickening and surrounding fat stranding', 'Appendix under 3 mm filled with air', 'Normal fat around a thin appendix', 'Calcified spleen'], 0, 'Enlarged, inflamed appendix with periappendiceal stranding.'],
    ['QZ-CTK', 'The most sensitive test for urinary tract stones is:', ['Non-contrast CT KUB', 'Contrast CT in portal phase only', 'Plain chest X-ray', 'CT head'], 0, 'Low-dose non-contrast CT shows nearly all stones.'],
    ['QZ-CTK', 'Small amounts of free intraperitoneal air on CT are best seen with a:', ['Wide (lung) window', 'Brain window', 'Bone window only', 'Narrow liver window only'], 0, 'A wide window separates air clearly from fat.'],
    ['QZ-CTK', 'The first-line drug for anaphylaxis to contrast is:', ['Intramuscular epinephrine (adrenaline)', 'Oral antihistamine only', 'IV contrast re-challenge', 'Paracetamol'], 0, 'IM epinephrine is first line; follow the hospital emergency protocol.'],
    ['QZ-CTK', 'The strongest risk factor for kidney injury after contrast is:', ['Pre-existing reduced kidney function', 'Young age', 'Normal eGFR', 'Low body temperature'], 0, 'Patients with low eGFR are at highest risk.'],
    ['QZ-CTK', 'A crescent-shaped collection that crosses suture lines is typical of:', ['Subdural haematoma', 'Extradural haematoma', 'Normal ventricle', 'Pituitary tumour'], 0, 'Subdural blood spreads along the convexity and crosses sutures but not the midline falx.'],

    // ── Ultrasound technical ──
    ['QZ-UST', 'Compared with a low-frequency transducer, a high-frequency transducer gives:', ['Better resolution but less penetration', 'More penetration but worse resolution', 'More penetration and better resolution', 'No difference'], 0, 'Higher frequency means finer detail, but sound is attenuated faster with depth.'],
    ['QZ-UST', 'Why is gel applied between the transducer and skin?', ['To remove air so sound passes into the body', 'To cool the transducer', 'To disinfect the skin', 'To increase the frame rate'], 0, 'Air reflects almost all sound. Gel couples the transducer to the skin.'],
    ['QZ-UST', 'The focal zone should be placed:', ['At or just below the area of interest', 'At the skin surface always', 'At the bottom of the screen always', 'Anywhere; it has no effect'], 0, 'Lateral resolution is best at the focus.'],
    ['QZ-UST', 'Time gain compensation (TGC) is used to:', ['Adjust brightness at different depths to make up for attenuation', 'Change the transducer frequency', 'Measure blood velocity', 'Freeze the image'], 0, 'Deeper echoes are weaker; TGC evens out brightness with depth.'],
    ['QZ-UST', 'Typical preparation for a gallbladder scan is:', ['Fasting for about 6–8 hours', 'A full bladder', 'No preparation at all', 'A fatty meal just before'], 0, 'Fasting distends the gallbladder and reduces bowel gas.'],
    ['QZ-UST', 'For a transabdominal pelvic scan, the bladder should be:', ['Full, to act as an acoustic window', 'Empty', 'Filled with air', 'It does not matter'], 0, 'A full bladder pushes bowel away and transmits sound well.'],
    ['QZ-UST', 'For Doppler velocity measurements, the angle of insonation should be:', ['60° or less', 'Exactly 90°', 'More than 75°', 'Any angle'], 0, 'At 90° no Doppler shift is detected; above 60° the velocity error rises sharply.'],
    ['QZ-UST', 'Which transducer is usually best for the thyroid?', ['High-frequency linear', 'Low-frequency curvilinear', 'Phased array cardiac', 'Endocavity'], 0, 'Superficial structures need high frequency and a linear footprint.'],
    ['QZ-UST', 'Colour Doppler shows aliasing. A good first step is to:', ['Increase the scale (PRF)', 'Decrease the scale', 'Increase the overall gain only', 'Switch to M-mode'], 0, 'Raising the PRF/scale (or moving the baseline) reduces aliasing.'],
    ['QZ-UST', 'A duplicate copy of the liver appearing above the diaphragm is a:', ['Mirror-image artifact', 'Acoustic shadow', 'Posterior enhancement', 'Real mass'], 0, 'The diaphragm acts as a strong reflector, creating a mirror image.'],
    ['QZ-UST', 'Turning up the overall gain:', ['Brightens the whole image, including noise', 'Improves resolution', 'Increases penetration depth', 'Reduces patient exposure'], 0, 'Gain amplifies all received echoes; too much adds noise.'],
    ['QZ-UST', 'Each ultrasound image saved for the report should be labelled with:', ['Patient details, date, body part and side where relevant', 'Only the machine name', 'Nothing; the radiologist remembers', 'The sonographer\'s phone number'], 0, 'Correct labelling makes images traceable and safe to report (P&P Process 4).'],

    // ── Ultrasound knowledge ──
    ['QZ-USK', 'A simple cyst on ultrasound is typically:', ['Anechoic, thin-walled, with posterior acoustic enhancement', 'Echogenic with a shadow', 'Solid with internal blood flow', 'Calcified'], 0, 'Clear fluid transmits sound well, brightening the tissue behind it.'],
    ['QZ-USK', 'A gallstone typically appears as:', ['An echogenic focus with posterior shadowing that moves with position', 'An anechoic area with enhancement', 'A thick hypoechoic wall only', 'Nothing visible'], 0, 'Stones reflect sound strongly and cast shadows; most move when the patient turns.'],
    ['QZ-USK', 'A positive sonographic Murphy sign suggests:', ['Acute cholecystitis', 'Kidney stone', 'Normal gallbladder', 'Liver cyst'], 0, 'Maximum tenderness when pressing over the gallbladder with the probe.'],
    ['QZ-USK', 'In an adult without previous gallbladder surgery, a normal common bile duct is usually:', ['About 6 mm or less', 'About 15 mm', 'Over 20 mm', 'Not visible ever'], 0, 'Around 6 mm or less is usually normal; it may be slightly larger with age.'],
    ['QZ-USK', 'Hydronephrosis means:', ['Dilatation of the renal pelvis and calyces', 'A small, scarred kidney', 'A kidney stone without obstruction', 'A normal kidney'], 0, 'Urine is backed up by obstruction or reflux.'],
    ['QZ-USK', 'In the FAST exam, the right upper quadrant view looks for free fluid in:', ['Morison\'s pouch (hepatorenal space)', 'The thyroid', 'The gallbladder lumen', 'The spinal canal'], 0, 'Free fluid collects early between the liver and right kidney.'],
    ['QZ-USK', 'A fatty liver usually appears:', ['Brighter (more echogenic) than the right kidney cortex', 'Darker than the kidney', 'Full of anechoic cysts', 'Smaller than normal'], 0, 'Fat increases echogenicity and can blur deeper structures.'],
    ['QZ-USK', 'An abdominal aorta is considered aneurysmal at a diameter of:', ['3 cm or more', '1 cm or more', '1.5 cm or more', '6 cm only'], 0, 'The usual threshold for AAA is ≥ 3 cm.'],
    ['QZ-USK', 'The key ultrasound sign of a deep vein thrombosis is:', ['A vein that does not compress with probe pressure', 'A vein that compresses fully', 'Increased arterial flow', 'A bright calcified wall'], 0, 'Normal veins collapse with gentle pressure; a clot prevents this.'],
    ['QZ-USK', 'In the first trimester, gestational age is best estimated by:', ['Crown–rump length', 'Femur length', 'Abdominal circumference', 'Placental thickness'], 0, 'CRL is the most accurate early dating measurement.'],
    ['QZ-USK', 'A positive pregnancy test with an empty uterus and an adnexal mass is worrying for:', ['Ectopic pregnancy', 'Normal early pregnancy only', 'Fibroid', 'Ovarian cyst only'], 0, 'This combination needs urgent clinical review.'],
    ['QZ-USK', 'On ultrasound, an inflamed appendix is typically:', ['Non-compressible and more than 6 mm in diameter', 'Compressible and 3 mm', 'Always invisible', 'Filled with air only'], 0, 'A blind-ending, non-compressible tube over 6 mm supports appendicitis.'],

    // ── Radiation protection ──
    ['QZ-RP', 'The three principles of radiation protection are:', ['Justification, optimisation and dose limitation', 'Speed, accuracy and cost', 'Time, cost and comfort', 'Shielding, shielding and shielding'], 0, 'Every exposure must be justified, kept as low as reasonably achievable, and within limits.'],
    ['QZ-RP', 'The three practical ways to reduce your own exposure are:', ['Time, distance and shielding', 'kVp, mAs and SID', 'Gloves, mask and gown', 'Speed, collimation and filters'], 0, 'Spend less time near the source, stay as far away as possible, and use barriers.'],
    ['QZ-RP', 'If you double your distance from a radiation source, your dose rate becomes about:', ['One quarter', 'One half', 'Double', 'Unchanged'], 0, 'Inverse square law: 2× distance → ¼ dose rate.'],
    ['QZ-RP', 'The usual occupational effective dose limit for radiation workers is:', ['20 mSv per year averaged over 5 years (max 50 mSv in any one year)', '1 mSv per year', '100 mSv per year', 'There is no limit'], 0, 'Based on ICRP / IAEA Basic Safety Standards. Check our licence conditions with the RPO.'],
    ['QZ-RP', 'The effective dose limit for members of the public is:', ['1 mSv per year', '20 mSv per year', '50 mSv per year', '5 mSv per month'], 0, 'Public limits are much lower than occupational limits.'],
    ['QZ-RP', 'The current occupational dose limit for the lens of the eye is:', ['20 mSv per year (averaged over 5 years)', '150 mSv per year', '500 mSv per year', 'No limit'], 0, 'The lens limit was lowered to 20 mSv/yr in the IAEA Basic Safety Standards.'],
    ['QZ-RP', 'Once a radiation worker declares a pregnancy, the dose to the embryo/fetus should be kept to about:', ['1 mSv for the rest of the pregnancy', '20 mSv per month', '50 mSv total', 'No special limit'], 0, 'Declare pregnancy to the RPO so work can be adjusted.'],
    ['QZ-RP', 'Your personal dosimeter should be:', ['Worn as the RPO directs whenever you do radiation work, and never shared or left in the room', 'Kept in the X-ray room so it records the room', 'Shared with whoever is on duty', 'Worn only during CT'], 0, 'It measures your dose only if you wear it correctly. The RPO reviews readings (monthly per our manual).'],
    ['QZ-RP', 'You find a crack in a lead apron. You should:', ['Remove it from use and report it to the RPO', 'Keep using it carefully', 'Fold it to cover the crack', 'Throw it away without telling anyone'], 0, 'Damaged PPE is removed from use pending assessment (P&P Process 14).'],
    ['QZ-RP', 'Cancer induction from radiation is an example of a:', ['Stochastic effect (probability rises with dose, no safe threshold assumed)', 'Deterministic effect with a fixed threshold', 'Effect that only happens above 10 Sv', 'Effect unrelated to dose'], 0, 'Stochastic risk is assumed to rise with any dose, which is why we optimise every exposure.'],
    ['QZ-RP', 'A child needs to be held still for an X-ray. The best choice is:', ['Use immobilisation aids, or a non-pregnant relative wearing a lead apron', 'A radiographer holds the child every time', 'A pregnant relative holds the child', 'Hold the child without protection to save time'], 0, 'Staff should not routinely hold patients. Use aids or a protected, non-pregnant carer.'],
    ['QZ-RP', 'Before an abdominal or pelvic X-ray on a woman of reproductive age, you should:', ['Ask about the possibility of pregnancy and follow protocol', 'Skip it if she looks young', 'Always refuse the exam', 'Ask only after the exposure'], 0, 'Pregnancy screening comes before exposure (P&P Processes 2 and 3).'],
    ['QZ-RP', 'Which exam usually gives the highest effective dose?', ['CT abdomen and pelvis', 'PA chest X-ray', 'Hand X-ray', 'Abdominal ultrasound'], 0, 'A CT abdomen-pelvis is roughly 10 mSv; a PA chest is about 0.02 mSv; ultrasound uses no ionising radiation.'],

    // ── Physics ──
    ['QZ-PHY', 'In an X-ray tube, most of the energy of the electrons hitting the anode becomes:', ['Heat', 'X-rays', 'Light', 'Sound'], 0, 'About 99% becomes heat, which is why tube heat limits matter.'],
    ['QZ-PHY', 'Most X-rays in the diagnostic beam are produced by:', ['Bremsstrahlung (braking) radiation', 'Characteristic radiation only', 'Radioactive decay', 'Piezoelectric crystals'], 0, 'Electrons slowing near nuclei produce a continuous spectrum.'],
    ['QZ-PHY', 'The photoelectric effect is most likely with:', ['Low photon energy and high atomic number', 'High photon energy and low atomic number', 'Ultrasound waves', 'Magnetic fields'], 0, 'It is why bone and iodine show strong contrast at diagnostic energies.'],
    ['QZ-PHY', 'The main source of scatter and of staff exposure in radiography is:', ['Compton scattering in the patient', 'The control console', 'The collimator light', 'Photoelectric absorption'], 0, 'Compton scatter leaves the patient in many directions.'],
    ['QZ-PHY', 'mAs mainly controls:', ['The quantity (number) of X-ray photons', 'The energy of the photons', 'The focal spot size', 'The SID'], 0, 'Doubling mAs doubles the number of photons; kVp controls their energy.'],
    ['QZ-PHY', 'Added filtration in the X-ray beam:', ['Removes low-energy photons and lowers skin dose', 'Increases skin dose', 'Removes high-energy photons only', 'Has no effect'], 0, 'Low-energy photons would be absorbed by skin without helping the image.'],
    ['QZ-PHY', 'The half-value layer (HVL) is:', ['The thickness of material that halves the beam intensity', 'Half the tube voltage', 'Half the grid ratio', 'The time for half the dose to decay'], 0, 'HVL describes beam quality (penetrating ability).'],
    ['QZ-PHY', 'The line focus principle allows:', ['A small effective focal spot with a larger area for heat on the anode', 'Higher kVp without a generator', 'Ultrasound focusing', 'CT without a detector'], 0, 'An angled anode makes the projected focal spot smaller.'],
    ['QZ-PHY', 'A smaller focal spot gives:', ['Better spatial resolution but lower heat capacity', 'Worse resolution', 'More scatter', 'Higher patient dose always'], 0, 'Small focus = sharper images, but less tolerance for high loads.'],
    ['QZ-PHY', 'The average speed of sound in soft tissue assumed by ultrasound machines is:', ['1540 m/s', '330 m/s', '3 × 10⁸ m/s', '15 m/s'], 0, 'Machines use 1540 m/s to convert echo time into depth.'],
    ['QZ-PHY', 'Ultrasound transducers create and receive sound using:', ['The piezoelectric effect', 'X-ray tubes', 'Radioactive sources', 'Magnetic resonance'], 0, 'Crystals vibrate when voltage is applied and produce voltage when echoes return.'],
    ['QZ-PHY', 'CT numbers (Hounsfield units) express a tissue\'s attenuation relative to:', ['Water', 'Lead', 'Bone', 'Air only'], 0, 'HU = 1000 × (μ tissue − μ water) / μ water.'],
    ['QZ-PHY', 'The gray (Gy) measures absorbed dose. The sievert (Sv) measures:', ['Equivalent or effective dose (biological risk)', 'Radioactivity', 'Beam energy', 'Exposure time'], 0, 'Sv weights absorbed dose by radiation type and tissue sensitivity.'],

    // ── SIMC policies (Draft 0) ──
    ['QZ-POL', 'A request is unclear and the side seems wrong. What should you do?', ['Stop before the exam, contact the requester, and record the clarification', 'Do both sides to be safe', 'Do the side you think is right', 'Ask the patient\'s companion to decide'], 0, 'Never examine from an unclear or apparently incorrect request (Process 1).'],
    ['QZ-POL', 'An unconscious patient arrives with missing details on the request. You should:', ['Get the information from the clinical team and document what is unavailable, without guessing', 'Fill in likely details yourself', 'Refuse the exam permanently', 'Wait until the patient wakes up'], 0, 'Process 1, step 4: do not guess.'],
    ['QZ-POL', 'Which identifiers does our department use to confirm a patient?', ['Name, date of birth and hospital number where available', 'Room number only', 'First name only', 'The request colour'], 0, 'Process 2 governing policy.'],
    ['QZ-POL', 'When can IV contrast be given for CT?', ['Only after screening, clinician review, consent, IV access and readiness checks are complete', 'As soon as the patient is on the table', 'After the scan if needed', 'When a verbal OK is given in the hallway'], 0, 'Process 5: never substitute a verbal workaround for formal approval.'],
    ['QZ-POL', 'An X-ray had to be repeated because of patient motion. Where is the reason recorded?', ['In the Division Census error/repeat field', 'Nowhere, if the second image is good', 'On the patient\'s hand', 'Only in a group chat'], 0, 'Every repeat reason is recorded (Process 3, step 6).'],
    ['QZ-POL', 'No qualified reader is available for an ultrasound today. What does the current rule say?', ['Do not perform the study; refer or reschedule and inform the requesting clinician', 'Scan it and save the images for next week', 'Let any staff scan and report it', 'Ask the patient to describe the findings'], 0, 'Process 4 authority and step 2.'],
    ['QZ-POL', 'Under the draft, an after-hours urgent CT (when no CT is scheduled) needs:', ['Appropriate indication, screening by the on-duty radiologist, and a referral form signed by the ER doctor', 'Only a text from the patient', 'Nothing; any request is scanned', 'Approval from billing'], 0, 'Process 5, step 9. Note: Process 16 says this workaround still needs MCC approval.'],
    ['QZ-POL', 'May full reports or identifiable images be sent through a group chat?', ['No, unless specifically authorised under hospital privacy/security policy', 'Yes, it is faster', 'Yes, if the chat has fewer than 10 members', 'Only on weekends'], 0, 'Process 8 governing policy.'],
    ['QZ-POL', 'Who may issue the diagnostic interpretation of an exam?', ['An authorised radiologist', 'Any technologist', 'The requesting nurse', 'The admin staff'], 0, 'Process 8 authority.'],
    ['QZ-POL', 'A critical finding is seen on a scan. What is required?', ['Communicate it promptly through the approved direct channel, document recipient and time, and escalate if they can\'t be reached', 'Wait for routine report delivery', 'Mention it at the next team meeting', 'Post it in a group chat'], 0, 'Process 9: critical findings must not wait for routine delivery.'],
    ['QZ-POL', 'Results are delivered physically to a ward. What must be obtained?', ['The receiving nurse\'s signature and the time, for every result', 'Nothing', 'The patient\'s fingerprint', 'A verbal thank you'], 0, 'Process 10, step 2.'],
    ['QZ-POL', 'A relative comes to collect a patient\'s result. You should:', ['Check an acceptable ID and confirm their authority to receive it, following privacy rules', 'Hand it over to anyone who knows the patient\'s name', 'Refuse all relatives always', 'Read the result aloud in the waiting area'], 0, 'Process 10 governing policy and step 3.'],
    ['QZ-POL', 'A patient lost their result. What is the correct action?', ['Reprint the original final report after verifying it; do not repeat the exam or write a new interpretation', 'Repeat the exam', 'Ask the radiologist to re-read and write a new report', 'Tell the patient it cannot be replaced'], 0, 'Process 11 and Process 17.'],
    ['QZ-POL', 'Images turn out to be attached to the wrong patient. First you should:', ['Stop release, isolate the case, notify the radiologist/shift lead and follow Process 17', 'Fix the name quietly and send', 'Delete the images', 'Send them anyway and correct later'], 0, 'Process 7, step 6.'],
    ['QZ-POL', 'An X-ray unit fails its user check. What do you do?', ['Stop using it, label it unavailable and raise a job order', 'Use it for urgent cases only', 'Keep using it and mention it later', 'Restart it until it works'], 0, 'Process 15, steps 3–4.'],
    ['QZ-POL', 'A patient falls in the department. The correct sequence is:', ['Give immediate assistance, call a doctor, then submit an incident report to the Patient Safety Committee', 'Finish the exam first', 'Only tell a colleague', 'Report it only if there is an injury'], 0, 'Process 17 governing policy.'],
    ['QZ-POL', 'Where should ALL radiology incidents be logged?', ['In the radiology portal, with an incident report', 'In a personal notebook only', 'Nowhere if no one was hurt', 'In a group chat'], 0, 'Process 17, step 1.'],
    ['QZ-POL', 'According to the department manual cited in the draft, staff dosimeter readings are reviewed by the RPO:', ['Monthly', 'Once every five years', 'Only after an incident', 'Never'], 0, 'Process 14: monthly review and an annual refresher (to be confirmed against licence conditions).'],
    ['QZ-POL', 'When may a new staff member work independently?', ['Only after the designated supervisor confirms readiness', 'From their first day', 'After one week automatically', 'Whenever they feel ready'], 0, 'Process 18 governing policy.'],
    ['QZ-POL', 'Second readings are classified as:', ['Agreement, minor discrepancy or major discrepancy', 'Good or bad', 'Fast or slow', 'Pass or fail only'], 0, 'Process 12, step 3.'],
    ['QZ-POL', 'A wrong or old result was released. What must happen?', ['Contain/recall it, release the correct report, notify the requesting doctor, and log an IR and Division Census error', 'Ignore it if the patient has left', 'Only delete the file', 'Wait for a complaint'], 0, 'Process 17 governing policy.']
  ];
  q.forEach(function (x) { x.push(x[0] === 'QZ-POL' ? P : ''); });
  return { quizzes: quizzes, questions: q };
}
