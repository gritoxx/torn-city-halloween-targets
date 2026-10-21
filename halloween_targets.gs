/**
 * Halloween Target List for Google Sheets (Apps Script)
 * Author:  Oxx [4114083]
 * Version: 1.1 (October 2026) - 1.1: profile links work in non-English sheets
 *
 * Each person runs their own copy: your key and your data stay in your own
 * Google account. The author never receives your key or your data.
 *
 * Uses only official APIs (no page scraping):
 *   - FFScouter /api/v1/get-targets: name, ID, level, fair fight, est. stats, last action
 *   - Torn API v1 user/{id}?selections=profile (Public access): faction, status, fresh last action
 *
 * Setup: Extensions > Apps Script > paste this > Save > reload the sheet.
 * Then use the "Halloween Targets" menu: Set API key, then Refresh list.
 *
 * Terms of use (Torn API ToS format) - also shown in the "About" tab:
 *   Data Storage:          Persistent - until you delete the rows or the sheet
 *   Data Sharing:          Nobody (only people you give access to this sheet)
 *   Purpose of Use:        Competitive advantage [personal list of attack targets for the Halloween event]
 *   Key Storage & Sharing: Stored in your private Apps Script user properties / Not shared;
 *                          sent only to ffscouter.com and api.torn.com
 *   Key Access Level:      Limited (required by FFScouter); Torn calls use only 'user' -> 'profile' (Public)
 *   Integrated service:    FFScouter - Data Policy and Terms at https://ffscouter.com
 */

// ---------- Settings ----------
// Edit these in the "Settings" tab of the sheet (created automatically).
// A blank value means "use FFScouter's default". With every filter blank,
// FFScouter returns a random mix, which is useful for "Add more targets".
const SETTINGS_SHEET = 'Settings';
const SETTING_ROWS = [
  // [name, default, note]
  ['limit',         50,   'Targets per pull, 1-50'],
  ['minff',         '',   'Minimum fair fight, 1.0 or higher (e.g. 1.5). Blank = default'],
  ['maxff',         '',   'Maximum fair fight, 1.0 or higher (e.g. 3.0). Blank = default'],
  ['minlevel',      '',   'Minimum level, 1-100. Blank = default'],
  ['maxlevel',      '',   'Maximum level, 1-100. Blank = default'],
  ['inactiveonly',  '',   '1 = only players inactive 14+ days (FFScouter default), 0 = include active players'],
  ['factionless',   '',   '1 = factionless players only, 0 or blank = everyone'],
  ['lookupFaction', true, 'Tick to look up faction + status (one Torn API call per target)']
];
const TORN_DELAY_MS = 750; // about 80 calls/min, under Torn's 100/min limit

const SHEET_NAME = 'Targets';
const HEADERS = ['Name', 'User ID', 'Profile URL', 'Fair Fight', 'Est. Total Stats', 'Level',
                 'Faction', 'Status', 'Last Action', 'Days Inactive', 'Updated'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Halloween Targets')
    .addItem('About & terms of use', 'openAbout')
    .addItem('Set API key', 'setApiKey')
    .addItem('Open settings', 'openSettings')
    .addItem('Refresh list (replace)', 'refreshTargets')
    .addItem('Add more targets (keep existing)', 'addTargets')
    .addToUi();
  try {
    if (!SpreadsheetApp.getActive().getSheetByName(ABOUT_SHEET)) writeAboutSheet_();
  } catch (e) { /* not critical */ }
}

// ---------- About & terms of use ----------
const ABOUT_SHEET = 'About';
const ABOUT_ROWS = [
  ['Halloween Target List', 'A personal Google Sheet that builds a list of possible attack targets for the Halloween event using official APIs only.'],
  ['Author', 'Oxx [4114083] - version 1.1 (October 2026)'],
  ['Your copy, your data', 'Everyone runs their own copy. Your key and data stay in your own Google account; the author never receives them.'],
  ['', ''],
  ['WHAT IT DOES', ''],
  ['Target list', 'Asks FFScouter for players in your fair fight / level range (name, ID, level, FF score, estimated total stats, last action).'],
  ['Faction & status', 'Optionally looks up each player\'s faction and current status with the Torn API (\'user\' -> \'profile\', a Public selection).'],
  ['Attacking', 'Done by you, by hand. The sheet only gives you a profile link; it never performs any in-game action.'],
  ['', ''],
  ['TERMS OF USE (Torn API ToS format)', ''],
  ['Data Storage', 'Persistent - until you delete the rows or the sheet.'],
  ['Data Sharing', 'Nobody (only people you choose to give access to this sheet).'],
  ['Purpose of Use', 'Competitive advantage [personal list of attack targets for the Halloween event].'],
  ['Key Storage & Sharing', 'Stored in your private Apps Script user properties, never in a cell / Not shared. Sent only to ffscouter.com and api.torn.com.'],
  ['Key Access Level', 'Limited Access (required by FFScouter registration). Direct Torn API calls use only \'user\' -> \'profile\' (Public).'],
  ['Integrated service', 'FFScouter. Its Data Policy and Terms and Conditions: https://ffscouter.com'],
  ['', ''],
  ['HOW WE FOLLOW THE API GUIDELINES', ''],
  ['API only', 'Data comes only from the Torn API and FFScouter\'s API. No torn.com pages are loaded or scraped.'],
  ['Read-only', 'The APIs are read-only. One click on a target link = one attack you make yourself.'],
  ['Minimal data', 'One Torn selection (\'profile\') per player, and only when the faction lookup is ticked in Settings.'],
  ['Rate limit', 'About 80 Torn requests per minute at most (Torn allows 100). If Torn reports the limit, lookups stop for that run.'],
  ['Invalid keys', 'The script stops at the first key error instead of retrying, so a bad key cannot trigger an IP block.'],
  ['Transparency', 'Every Torn request is tagged "HalloweenSheet", so you can see it in your API key log (\'key\' -> \'log\').'],
  ['Key safety', 'Your key is never written to the sheet. Share the sheet view-only; editors can change the script.'],
  ['Sharing with others', 'If other players will enter their own keys, keep this terms table visible where they do so.']
];

function openAbout() {
  writeAboutSheet_();
  SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName(ABOUT_SHEET));
}

function writeAboutSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(ABOUT_SHEET) || ss.insertSheet(ABOUT_SHEET, 0);
  sh.clear();
  sh.getRange(1, 1, ABOUT_ROWS.length, 2).setValues(ABOUT_ROWS)
    .setWrap(true).setVerticalAlignment('top');
  sh.getRange(1, 1, ABOUT_ROWS.length, 1).setFontWeight('bold');
  sh.getRange(1, 1).setFontSize(14);
  ABOUT_ROWS.forEach((r, i) => {
    if (r[0] && !r[1]) sh.getRange(i + 1, 1, 1, 2).setBackground('#efefef');
  });
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 620);
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'API key',
    'Paste your 16-character Torn API key (the one registered on ffscouter.com).\n\n' +
    'Terms of use: target data is stored in this sheet until you delete it and is shared with nobody. ' +
    'Your key is stored in your private Apps Script settings (never in the sheet), is not shared, and is sent ' +
    'only to ffscouter.com and api.torn.com. Key level: Limited (needed by FFScouter); Torn calls use only ' +
    '\'user\' -> \'profile\'. FFScouter Data Policy & Terms: https://ffscouter.com. Full details: About tab.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!/^[A-Za-z0-9]{16}$/.test(key)) {
    ui.alert('That does not look like a 16-character key.');
    return;
  }
  PropertiesService.getUserProperties().setProperty('TORN_KEY', key);
  ui.alert('Key saved.');
}

function refreshTargets() { run_(false); }
function addTargets() { run_(true); }

function run_(append) {
  const ss = SpreadsheetApp.getActive();
  const key = getKey_();
  const SETTINGS = getSettings_();
  const sheet = getSheet_();

  const existing = new Set();
  if (append && sheet.getLastRow() > 1) {
    sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues()
      .forEach(r => existing.add(Number(r[0])));
  } else if (!append && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).clearContent();
  }

  ss.toast('Fetching targets from FFScouter…');
  const targets = fetchFfTargets_(key, SETTINGS).filter(t => !existing.has(Number(t.player_id)));
  if (!targets.length) {
    ss.toast('No new targets matched your settings.');
    return;
  }

  const now = new Date();
  let tornOk = SETTINGS.lookupFaction;
  const rows = [];

  targets.forEach((t, i) => {
    let faction = '';
    let status = '';
    let lastAction = t.last_action || null;

    if (tornOk) {
      if (i % 10 === 0) ss.toast('Looking up factions… ' + i + '/' + targets.length);
      const p = fetchTornProfile_(key, t.player_id);
      if (p === 'RATE_LIMIT') {
        tornOk = false; // stop further Torn calls, keep the FFScouter data
      } else if (p) {
        faction = (p.faction && p.faction.faction_id)
          ? p.faction.faction_name + ' [' + p.faction.faction_tag + ']'
          : 'None';
        status = p.status ? p.status.description : '';
        if (p.last_action && p.last_action.timestamp) lastAction = p.last_action.timestamp;
      }
      Utilities.sleep(TORN_DELAY_MS);
    }

    const url = 'https://www.torn.com/profiles.php?XID=' + t.player_id;
    const lastDate = lastAction ? new Date(lastAction * 1000) : '';
    const days = lastAction ? Math.floor((now.getTime() / 1000 - lastAction) / 86400) : '';

    rows.push([
      t.name,
      t.player_id,
      url, // turned into a clickable link below (no formula, so it works in every language/locale)
      t.fair_fight,
      t.bs_estimate,
      t.level,
      faction,
      status,
      lastDate,
      days,
      now
    ]);
  });

  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, HEADERS.length).setValues(rows);
  // Clickable profile links without =HYPERLINK, which can break in non-English sheets
  // (some locales use ";" instead of "," between formula arguments).
  const links = rows.map(r => [
    SpreadsheetApp.newRichTextValue().setText(r[2]).setLinkUrl(r[2]).build()
  ]);
  sheet.getRange(start, 3, rows.length, 1).setRichTextValues(links);
  sheet.getRange(start, 4, rows.length, 1).setNumberFormat('0.00');
  sheet.getRange(start, 5, rows.length, 1)
    .setNumberFormat('[>=1000000000]0.00,,,"b";[>=1000000]0.00,,"m";#,##0');
  sheet.getRange(start, 9, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(start, 11, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.autoResizeColumns(1, HEADERS.length);

  ss.toast('Added ' + rows.length + ' targets' + (tornOk || !SETTINGS.lookupFaction ? '.' : ' (Torn rate limit hit, so some factions are blank).'));
}

function openSettings() {
  const sh = getSettingsSheet_();
  SpreadsheetApp.getActive().setActiveSheet(sh);
}

function getSettingsSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Notes']]).setFontWeight('bold');
    sh.getRange(2, 1, SETTING_ROWS.length, 3).setValues(SETTING_ROWS);
    sh.setFrozenRows(1);
    const lookupRow = 2 + SETTING_ROWS.findIndex(r => r[0] === 'lookupFaction');
    sh.getRange(lookupRow, 2).insertCheckboxes().setValue(true);
    sh.autoResizeColumns(1, 3);
  }
  return sh;
}

function getSettings_() {
  const sh = getSettingsSheet_();
  const values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
  const map = {};
  values.forEach(r => { if (r[0] !== '') map[String(r[0]).trim()] = r[1]; });

  const s = {};
  SETTING_ROWS.forEach(([name, def]) => {
    const v = (name in map) ? map[name] : def;
    s[name] = (v === '' || v === null) ? null : v;
  });
  s.lookupFaction = s.lookupFaction === true || String(s.lookupFaction).toUpperCase() === 'TRUE';
  if (s.limit === null) s.limit = 50;

  // Basic checks so mistakes show a clear message instead of an FFScouter error code
  const num = n => s[n] === null || !isNaN(Number(s[n]));
  ['limit', 'minff', 'maxff', 'minlevel', 'maxlevel', 'inactiveonly', 'factionless'].forEach(n => {
    if (!num(n)) throw new Error('Settings: "' + n + '" must be a number or blank.');
  });
  if (s.limit < 1 || s.limit > 50) throw new Error('Settings: limit must be 1-50.');
  if (s.minff !== null && s.minff < 1) throw new Error('Settings: minff must be 1.0 or higher.');
  if (s.maxff !== null && s.maxff < 1) throw new Error('Settings: maxff must be 1.0 or higher.');
  if (s.minff !== null && s.maxff !== null && Number(s.minff) > Number(s.maxff)) {
    throw new Error('Settings: minff cannot be greater than maxff.');
  }
  if (s.minlevel !== null && Number(s.minlevel) > 100) throw new Error('Settings: minlevel must be 1-100.');
  if (s.maxlevel !== null && Number(s.maxlevel) > 100) throw new Error('Settings: maxlevel must be 1-100.');
  if (s.minlevel !== null && s.maxlevel !== null && Number(s.minlevel) > Number(s.maxlevel)) {
    throw new Error('Settings: minlevel cannot be greater than maxlevel.');
  }
  return s;
}

function fetchFfTargets_(key, SETTINGS) {
  const params = ['key=' + encodeURIComponent(key)];
  ['limit', 'minff', 'maxff', 'minlevel', 'maxlevel', 'inactiveonly', 'factionless'].forEach(k => {
    if (SETTINGS[k] !== null && SETTINGS[k] !== undefined) {
      params.push(k + '=' + encodeURIComponent(SETTINGS[k]));
    }
  });
  const res = UrlFetchApp.fetch('https://ffscouter.com/api/v1/get-targets?' + params.join('&'),
                                { muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());
  if (data.code === 17) return []; // no targets matched
  if (data.error) throw new Error('FFScouter error ' + data.code + ': ' + data.error);
  return data.targets || [];
}

function fetchTornProfile_(key, id) {
  const url = 'https://api.torn.com/user/' + id +
              '?selections=profile&comment=HalloweenSheet&key=' + encodeURIComponent(key);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  let data;
  try { data = JSON.parse(res.getContentText()); } catch (e) { return null; }
  if (data.error) {
    const code = data.error.code;
    if (code === 5) return 'RATE_LIMIT';
    // Key problems: stop immediately so a bad key isn't retried (Torn may IP-block repeated invalid keys)
    if ([1, 2, 10, 13, 16, 18].includes(code)) {
      throw new Error('Torn API error ' + code + ': ' + data.error.error);
    }
    return null; // other errors: skip this player
  }
  return data;
}

function getKey_() {
  const key = PropertiesService.getUserProperties().getProperty('TORN_KEY');
  if (!key) throw new Error('No API key set. Use Halloween Targets > Set API key first.');
  return key;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}
