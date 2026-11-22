/**
 * SIEGECRAFT BETA — GOOGLE WORKSPACE + GOOGLE SHEETS
 * Deploy as web app: Execute as ME; Access: anyone in your Workspace DOMAIN.
 * Never share the spreadsheet with students. Their identities come from
 * Session.getActiveUser() on the SERVER, not from any input in the browser.
 *
 * Edit SCHOOL_DOMAIN to match the actual student's account domain.
 */
const SCHOOL_DOMAIN = 'monet.k12.ca.us';
const SHEET_NAME = 'Student Progress';
const COLUMNS = ['Student Email','Display Name','Coins','Best Streak','Mined Coins','Correct Mine Answers','Greek Forge Unlocks','Greek Territories','Updated At','Revision','Save JSON'];

function setupSiegecraft() {
  // Works from BOTH standalone Apps Script projects and Sheets-bound projects.
  // Reuses the same spreadsheet on later runs, including already-saved students.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const existingId = props.getProperty('SIEGECRAFT_SHEET_ID');
    let book;
    let created = false;

    if (existingId) {
      // Protect existing beta progress: NEVER create a second database simply
      // because a previously configured spreadsheet cannot be opened.
      try {
        book = SpreadsheetApp.openById(existingId);
      } catch (err) {
        throw new Error('Siegecraft is already linked to a spreadsheet, but it could not be opened. ' +
          'Check your access to this spreadsheet: https://docs.google.com/spreadsheets/d/' + existingId + '/edit');
      }
    } else {
      book = SpreadsheetApp.getActiveSpreadsheet();
      if (!book) {
        book = SpreadsheetApp.create('Siegecraft – Student Progress');
        created = true;
      }
    }

    let sheet = book.getSheetByName(SHEET_NAME);
    if (!sheet) {
      // New spreadsheets have a blank 'Sheet1' tab; rename rather than leave
      // an unused extra tab. Existing sheets are never renamed or removed.
      const tabs = book.getSheets();
      if (created && tabs.length === 1 && tabs[0].getLastRow() === 0) {
        sheet = tabs[0];
        sheet.setName(SHEET_NAME);
      } else {
        sheet = book.insertSheet(SHEET_NAME);
      }
    }

    if (sheet.getLastRow() === 0) sheet.appendRow(COLUMNS);
    else sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@');
    sheet.getRange('I:I').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.autoResizeColumns(1, 10);

    // Store only after successful setup, then return the URL in the log.
    props.setProperty('SIEGECRAFT_SHEET_ID', book.getId());
    const message = 'Siegecraft spreadsheet ready: ' + book.getUrl() +
      ' | Student data tab: ' + SHEET_NAME;
    Logger.log(message);
    return message;
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Siegecraft · Student Beta');
}

function studentEmail_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) throw new Error('Your Google account could not be identified. Sign in with your district account; ask the teacher to verify the web app is restricted to the school domain.');
  if (!email.endsWith('@' + SCHOOL_DOMAIN)) throw new Error('Please open Siegecraft using your school Google account (' + SCHOOL_DOMAIN + ').');
  return email;
}

function dataSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SIEGECRAFT_SHEET_ID');
  if (!id) throw new Error('Siegecraft has not been configured yet. The teacher must run setupSiegecraft() in the Script editor.');
  const sheet = SpreadsheetApp.openById(id).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Missing Student Progress tab. Run setupSiegecraft() again.');
  return sheet;
}

function findStudentRow_(sheet,email) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const matches = sheet.getRange(2,1,last-1,1).createTextFinder(email).matchEntireCell(true).findNext();
  return matches ? matches.getRow() : 0;
}

function displayName_(email) {
  // Do not assume name is verified by Google; this is only a friendly label.
  const text = email.split('@')[0].replace(/[._-]+/g,' ').replace(/\s+/g,' ').trim();
  return (text || 'Student').slice(0,26);
}
function freshProfile_(name) {
  return {
    name:name, coins:40, totalMined:0, totalCorrect:0, bestStreak:0,
    forge:{archer:false,cavalry:false,catapult:false},
    forgeStep:{archer:0,cavalry:0,catapult:0}, campaign:null
  };
}
function number_(v,min,max) {
  const n=Number(v);
  return Number.isFinite(n)?Math.min(max,Math.max(min,Math.floor(n))):min;
}
function cleanProfile_(raw,name) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw new Error('Invalid save payload.');
  const profile = freshProfile_(name);
  profile.coins=number_(raw.coins,0,1000000);
  profile.totalMined=number_(raw.totalMined,0,10000000);
  profile.totalCorrect=number_(raw.totalCorrect,0,1000000);
  profile.bestStreak=number_(raw.bestStreak,0,100000);
  for (const u of ['archer','cavalry','catapult']) {
    profile.forge[u]=raw.forge?.[u]===true;
    profile.forgeStep[u]=number_(raw.forgeStep?.[u],0,3);
  }
  if (raw.campaign !== null && raw.campaign !== undefined) {
    const c=raw.campaign;
    if (typeof c!=='object'||!Array.isArray(c.terr)||c.terr.length!==8) throw new Error('Invalid territory save.');
    profile.campaign={
      phase:['muster','player'].includes(c.phase)?c.phase:'player',
      round:number_(c.round,1,8),redCoins:number_(c.redCoins,0,1000000),
      terr:c.terr.map(t=>{
        if(!t||typeof t!=='object')throw new Error('Invalid territory entry.');
        const units={};for(const k of ['hoplite','archer','cavalry','catapult']) units[k]=number_(t.units?.[k],0,1000);
        return {owner:['blue','red','none'].includes(t.owner)?t.owner:'none',units,acted:t.acted===true};
      })
    };
  }
  if(JSON.stringify(profile).length>40000) throw new Error('Save is too large.');
  return profile;
}
function summaryRow_(email,name,profile,revision) {
  const unlocked=['archer','cavalry','catapult'].filter(k=>profile.forge[k]).join(', ')||'Hoplite only';
  const blue=profile.campaign?.terr?.filter(t=>t.owner==='blue').length || 0;
  return [email,name,profile.coins,profile.bestStreak,profile.totalMined,profile.totalCorrect,unlocked,blue,new Date(),revision,JSON.stringify(profile)];
}

function getStudentProfile() {
  const email=studentEmail_(),sheet=dataSheet_();
  const lock=LockService.getScriptLock();lock.waitLock(12000);
  try {
    let row=findStudentRow_(sheet,email);
    if (!row) {
      const name=displayName_(email),profile=freshProfile_(name);
      sheet.appendRow(summaryRow_(email,name,profile,1));
      row=sheet.getLastRow();
    }
    const cells=sheet.getRange(row,1,1,COLUMNS.length).getValues()[0];
    const name=String(cells[1]||displayName_(email));
    let parsed;
    try{parsed=JSON.parse(cells[10]);}catch(_){throw new Error('Your saved profile could not be read. Ask the teacher to inspect your spreadsheet row.');}
    return {email,name,revision:number_(cells[9],1,2147483647),profile:cleanProfile_(parsed,name)};
  } finally {lock.releaseLock();}
}

function saveStudentProfile(request) {
  const email=studentEmail_(),sheet=dataSheet_();
  if (!request || typeof request!=='object') throw new Error('Missing save request.');
  const expected=Number(request.revision);
  if (!Number.isSafeInteger(expected)||expected<1) throw new Error('Invalid save revision.');
  const lock=LockService.getScriptLock();lock.waitLock(12000);
  try {
    const row=findStudentRow_(sheet,email);
    if (!row) throw new Error('No student profile exists for this Google account. Refresh the game.');
    const cells=sheet.getRange(row,1,1,COLUMNS.length).getValues()[0];
    const actual=number_(cells[9],1,2147483647);
    if (expected!==actual) return {ok:false,conflict:true,revision:actual,message:'This account has a newer save, perhaps from another tab or computer. Reload to avoid overwriting it.'};
    const name=String(cells[1]||displayName_(email));
    const profile=cleanProfile_(request.profile,name);
    const revision=actual+1;
    sheet.getRange(row,1,1,COLUMNS.length).setValues([summaryRow_(email,name,profile,revision)]);
    SpreadsheetApp.flush();
    return {ok:true,revision,savedAt:new Date().toISOString()};
  } finally {lock.releaseLock();}
}
