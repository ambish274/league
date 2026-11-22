function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 League Manager')
    .addItem('Sync League Standings Now', 'syncLeagueStandings')
    .addItem('Setup Scorecard Viewer', 'setupScorecardViewer')
    .addToUi();
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    
    if (action === 'createTournament') return createTournament(payload);
    if (action === 'getTournament') return getTournament(payload);
    if (action === 'submitEnd') return submitEnd(payload);
    if (action === 'finalizeScore') return finalizeScore(payload);
    if (action === 'submitPaperScore') return submitPaperScore(payload);
    if (action === 'getShootersForPin') return getShootersForPin(payload);
    if (action === 'getLeagueRoster') return getLeagueRoster(payload);
    if (action === 'getLiveLeaderboardData') return getLiveLeaderboardData(payload);
    if (action === 'getLeagueWeeks') return getLeagueWeeks();
    if (action === 'getPaperEntryShooters') return getPaperEntryShooters(payload);
    
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Unknown action'})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------------------------------------------------------
// HELPER FUNCTIONS
// ---------------------------------------------------------

// Enforces proper Title Case for names (e.g., "john doe" -> "John Doe")
function toTitleCase(str) {
  if (!str) return "";
  return str.toString().toLowerCase().replace(/(?:^|[\s-])\w/g, function(match) {
    return match.toUpperCase();
  });
}


// ---------------------------------------------------------
// ROUTER FUNCTIONS
// ---------------------------------------------------------

function getSetupSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Setup");
  if (!sheet) {
    sheet = ss.insertSheet("Setup");
    sheet.appendRow(["Timestamp", "PIN", "ConfigJSON"]);
  }
  return sheet;
}

function createTournament(payload) {
  const sheet = getSetupSheet();
  sheet.appendRow([new Date(), payload.pin.toString().trim(), JSON.stringify(payload.config)]);
  return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
}

function getTournament(payload) {
  const sheet = getSetupSheet();
  const data = sheet.getDataRange().getValues();
  
  const searchPin = payload.pin ? payload.pin.toString().trim() : "";
  
  for (let i = data.length - 1; i >= 0; i--) {
    const rowPin = data[i][1] ? data[i][1].toString().trim() : "";
    
    if (rowPin === searchPin) {
      try {
        return ContentService.createTextOutput(JSON.stringify({
          status: 'success', 
          config: JSON.parse(data[i][2])
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (e) {
        continue; 
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Tournament not found'})).setMimeType(ContentService.MimeType.JSON);
}

function getScoresSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Raw Scores");
  if (!sheet) {
    sheet = ss.insertSheet("Raw Scores");
    sheet.appendRow(["Timestamp", "PIN", "Shooter Name", "Class", "End Number", "Arrow Values", "End Total", "Running Total", "Scorekeeper ID", "Locked Status"]);
  }
  return sheet;
}

function submitEnd(payload) {
  const sheet = getScoresSheet();
  sheet.appendRow([
    new Date(), payload.tournamentId, toTitleCase(payload.shooterName), payload.shooterClass, 
    payload.endNumber, payload.arrowValues, payload.endTotal, payload.runningTotal, 
    payload.scorekeeperId, false
  ]);
  syncLeagueStandings(true); 
  return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
}

function finalizeScore(payload) {
  const sheet = getScoresSheet();
  sheet.appendRow([
    new Date(), payload.tournamentId, toTitleCase(payload.shooterName), payload.shooterClass, 
    "FINAL", "", payload.finalTotal, payload.finalTotal, 
    payload.scorekeeperId, true
  ]);
  syncLeagueStandings(true); 
  return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
}

function submitPaperScore(payload) {
  const sheet = getScoresSheet();
  const data = sheet.getDataRange().getValues();
  const targetPin = payload.pin ? payload.pin.toString().trim() : "";
  const targetName = payload.shooterName ? payload.shooterName.toString().trim().toLowerCase() : "";
  let replaced = false;

  // Remove any existing rows for this shooter and PIN (bottom-to-top) to prevent double-counting
  for (let i = data.length - 1; i >= 0; i--) {
    const rowPin = data[i][1] ? data[i][1].toString().trim() : "";
    const rowName = data[i][2] ? data[i][2].toString().trim().toLowerCase() : "";
    if (rowPin === targetPin && rowName === targetName) {
      sheet.deleteRow(i + 1);
      replaced = true;
    }
  }

  sheet.appendRow([
    new Date(), targetPin, toTitleCase(payload.shooterName), payload.shooterClass, 
    "PAPER", "Stat1:"+payload.stat1+"|Stat2:"+payload.stat2, payload.totalScore, payload.totalScore, 
    "PAPER_ENTRY", true
  ]);
  syncLeagueStandings(true); 
  return ContentService.createTextOutput(JSON.stringify({status: 'success', replaced: replaced})).setMimeType(ContentService.MimeType.JSON);
}

function getShootersForPin(payload) {
  const sheet = getScoresSheet();
  const data = sheet.getDataRange().getValues();
  const shooters = {};
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString() === payload.pin.toString()) {
      const name = toTitleCase(data[i][2]);
      const sClass = data[i][3];
      shooters[name] = { name: name, sClass: sClass, shooterId: name };
    }
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    shooters: Object.values(shooters)
  })).setMimeType(ContentService.MimeType.JSON);
}

function getLeagueRoster(payload) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("League Roster");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'League Roster tab not found. Please create one to use this feature.'})).setMimeType(ContentService.MimeType.JSON);
  }
  
  const data = sheet.getDataRange().getValues();
  let scoredShooters = new Set();
  
  if (payload.pin) {
    let rawSheet = ss.getSheetByName("Raw Scores");
    if (rawSheet) {
      const rawData = rawSheet.getDataRange().getValues();
      for (let i = 0; i < rawData.length; i++) {
        if (rawData[i][1] && rawData[i][1].toString() === payload.pin.toString()) {
          scoredShooters.add(rawData[i][2].toString().trim().toLowerCase());
        }
      }
    }
  }

  const roster = [];
  data.forEach((row, idx) => {
    const rawName = row[0] ? row[0].toString().trim() : "";
    const name = toTitleCase(rawName);
    const sClass = row[1] ? row[1].toString().trim() : "";
    
    if (name && name.toLowerCase() !== "shooter name" && name.toLowerCase() !== "name") {
      const hasScored = scoredShooters.has(name.toLowerCase());
      roster.push({ name: name, sClass: sClass, hasScored: hasScored });
    }
  });
  
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    roster: roster
  })).setMimeType(ContentService.MimeType.JSON);
}

function getLeagueWeeks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Live Config");
  if (!configSheet) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Live Config tab not found.'})).setMimeType(ContentService.MimeType.JSON);
  }

  const data = configSheet.getDataRange().getValues();
  const weeks = [];

  data.forEach(row => {
    const weekName = row[0] ? row[0].toString().trim() : "";
    const pin = row[1] ? row[1].toString().trim() : "";
    const roundType = row[2] ? row[2].toString().trim() : "";
    const status = row[3] ? row[3].toString().trim() : "";

    // Skip header rows and upcoming weeks without a PIN assigned
    if (weekName && pin && pin.toUpperCase() !== "PIN" && !weekName.toLowerCase().includes("week number")) {
      weeks.push({
        weekName: weekName,
        pin: pin,
        roundType: roundType || "Standard",
        status: status
      });
    }
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    weeks: weeks
  })).setMimeType(ContentService.MimeType.JSON);
}

function getPaperEntryShooters(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetPin = payload.pin ? payload.pin.toString().trim() : "";
  if (!targetPin) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'No PIN provided.'})).setMimeType(ContentService.MimeType.JSON);
  }

  // Look up round type from Live Config so tie-breakers calculate accurately
  let roundType = "";
  const configSheet = ss.getSheetByName("Live Config");
  if (configSheet) {
    const configData = configSheet.getDataRange().getValues();
    for (let i = 0; i < configData.length; i++) {
      if (configData[i][1] && configData[i][1].toString().trim() === targetPin) {
        roundType = configData[i][2] ? configData[i][2].toString().trim().toLowerCase() : "";
        break;
      }
    }
  }

  // Tally any existing scores in Raw Scores for this PIN
  const scoredMap = {};
  const rawSheet = ss.getSheetByName("Raw Scores");
  if (rawSheet) {
    const rawData = rawSheet.getDataRange().getValues();
    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];
      const pin = row[1] ? row[1].toString().trim() : "";
      if (pin !== targetPin) continue;

      const name = toTitleCase(row[2] ? row[2].toString().trim() : "");
      if (!name) continue;
      const key = name.toLowerCase();
      const sClass = row[3] ? row[3].toString().trim() : "";
      const endNumber = row[4] ? row[4].toString().trim() : "";
      const arrowsStr = row[5] ? row[5].toString().trim() : "";
      const endTotal = parseInt(row[6]) || 0;

      if (!scoredMap[key]) {
        scoredMap[key] = { name: name, sClass: sClass, score: 0, tb1: 0, tb2: 0 };
      }

      if (endNumber === "FINAL") {
        // Skip duplicate total row
      } else if (endNumber === "PAPER") {
        scoredMap[key].score += endTotal;
        if (arrowsStr.includes('|')) {
          const parts = arrowsStr.split('|');
          scoredMap[key].tb1 += parseInt(parts[0].split(':')[1]) || 0;
          scoredMap[key].tb2 += parseInt(parts[1].split(':')[1]) || 0;
        }
      } else {
        scoredMap[key].score += endTotal;
        if (arrowsStr) {
          const allArrows = arrowsStr.split(',');
          allArrows.forEach(a => {
            const valStr = a.trim().toUpperCase();
            const valNum = Number(valStr) || 0;
            if (roundType.includes('asa_champ') || roundType.includes('champ')) { if (valStr === '14') scoredMap[key].tb1++; else if (valStr === '12') scoredMap[key].tb2++; }
            else if (roundType.includes('asa')) { if (valStr === '12') scoredMap[key].tb1++; else if (valStr === '10') scoredMap[key].tb2++; }
            else if (roundType.includes('ibo')) { if (valStr === '11') scoredMap[key].tb1++; else if (valStr === '10') scoredMap[key].tb2++; }
            else if (roundType.includes('vegas')) { if (valStr === 'X' || valStr === '11') scoredMap[key].tb1++; else if (valStr === '10') scoredMap[key].tb2++; }
            else if (roundType.includes('nfaa')) { if (valStr === 'X' || valStr === '6') scoredMap[key].tb1++; else if (valStr === '5') scoredMap[key].tb2++; }
            else if (roundType.includes('fred_bear')) {
              if (valNum === 5 || valNum === 11 || valNum === 10 || valStr === 'X') scoredMap[key].tb1++;
              else if (valNum === -5) scoredMap[key].tb2++;
            }
            else if (roundType.includes('custom')) {
              if (valStr === 'X' || valStr === '14' || valStr === '12' || valStr === '11') scoredMap[key].tb1++;
              else if (valStr === '10' || valStr === '5') scoredMap[key].tb2++;
            }
          });
        }
      }
    }
  }

  // Merge League Roster with any walk-in shooters already in Raw Scores
  const shooters = [];
  const seenKeys = new Set();
  const rosterSheet = ss.getSheetByName("League Roster");

  if (rosterSheet) {
    const rosterData = rosterSheet.getDataRange().getValues();
    rosterData.forEach(row => {
      const rawName = row[0] ? row[0].toString().trim() : "";
      const name = toTitleCase(rawName);
      const sClass = row[1] ? row[1].toString().trim() : "";

      if (name && name.toLowerCase() !== "shooter name" && name.toLowerCase() !== "name") {
        const key = name.toLowerCase();
        seenKeys.add(key);
        const existing = scoredMap[key];
        shooters.push({
          name: name,
          sClass: sClass || (existing ? existing.sClass : ""),
          hasScored: !!existing,
          score: existing ? existing.score : null,
          tb1: existing ? existing.tb1 : null,
          tb2: existing ? existing.tb2 : null
        });
      }
    });
  }

  // Include anyone who scored for this PIN but isn't on the League Roster tab
  Object.keys(scoredMap).forEach(key => {
    if (!seenKeys.has(key)) {
      const s = scoredMap[key];
      shooters.push({
        name: s.name,
        sClass: s.sClass,
        hasScored: true,
        score: s.score,
        tb1: s.tb1,
        tb2: s.tb2
      });
    }
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    shooters: shooters
  })).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------
// LIVE SCOREBOARD DATA API
// ---------------------------------------------------------
function getLiveLeaderboardData(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName("Live Config");
  if (!configSheet) return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'No config sheet'})).setMimeType(ContentService.MimeType.JSON);
  
  const configData = configSheet.getDataRange().getValues();
  let activePin = null;
  let activeWeekRaw = "";
  let roundType = "";
  
  // 1. Find the currently Active PIN and Week label
  for (let i = 0; i < configData.length; i++) {
    const status = configData[i][3] ? configData[i][3].toString().trim().toUpperCase() : "";
    if (status === "ACTIVE") {
      activeWeekRaw = configData[i][0] ? configData[i][0].toString().trim() : "";
      activePin = configData[i][1] ? configData[i][1].toString().trim() : "";
      roundType = configData[i][2] ? configData[i][2].toString().trim().toLowerCase() : "";
      break;
    }
  }
  
  if (!activePin) {
    return ContentService.createTextOutput(JSON.stringify({status: 'success', weekLabel: '', data: []})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // Build a compact week prefix (e.g. "Week 1" or "Week #1" -> "Wk 1")
  let wkLabel = "";
  if (activeWeekRaw) {
    const wkMatch = activeWeekRaw.match(/week\s*#?\s*(\d+)/i);
    if (wkMatch) wkLabel = "Wk " + wkMatch[1];
    else if (/^\d+$/.test(activeWeekRaw)) wkLabel = "Wk " + activeWeekRaw;
    else wkLabel = activeWeekRaw;
  }

  let rawSheet = ss.getSheetByName("Raw Scores");
  if (!rawSheet) return ContentService.createTextOutput(JSON.stringify({status: 'success', weekLabel: activeWeekRaw, data: []})).setMimeType(ContentService.MimeType.JSON);
  
  const rawData = rawSheet.getDataRange().getValues();
  const shooters = {};
  
  // 2. Tally up the scores for the active PIN
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const pin = row[1] ? row[1].toString().trim() : "";
    
    if (pin === activePin) {
      const name = toTitleCase(row[2].toString().trim());
      const className = row[3].toString().trim();
      const endNumber = row[4].toString().trim();
      const endTotal = parseInt(row[6]) || 0;
      
      if (!shooters[className]) shooters[className] = {};
      if (!shooters[className][name]) shooters[className][name] = { score: 0, thruNum: 0, isFinal: false };
      
      const s = shooters[className][name];
      
      if (endNumber === "FINAL" || endNumber === "PAPER") {
         s.isFinal = true;
         if (endNumber === "PAPER") s.score += endTotal;
      } else {
         s.score += endTotal;
         const pt = parseInt(endNumber);
         if (!isNaN(pt)) {
             s.thruNum = pt; // track latest physical target/end
         }
      }
    }
  }
  
  const is3D = roundType.includes('ibo') || roundType.includes('asa') || roundType.includes('fred_bear');
  const labelPrefix = is3D ? "Tgt " : "End ";
  
  const boardData = [];
  
  // 3. Format and Sort the Data
  Object.keys(shooters).forEach(className => {
     const classGroup = { className: className, shooters: [] };
     
     Object.keys(shooters[className]).forEach(name => {
         const s = shooters[className][name];
         let thruText = "";
         if (s.isFinal) {
             thruText = wkLabel ? (wkLabel + " - Finished") : "Finished";
         } else if (s.thruNum > 0) {
             thruText = wkLabel ? (wkLabel + " - " + labelPrefix + s.thruNum) : (labelPrefix + s.thruNum);
         } else {
             thruText = wkLabel ? (wkLabel + " - Starting...") : "Starting...";
         }
         
         classGroup.shooters.push({
             name: name,
             score: s.score,
             thru: thruText,
             isFinal: s.isFinal
         });
     });
     
     // Sort shooters highest to lowest
     classGroup.shooters.sort((a, b) => b.score - a.score);
     
     // Assign numeric ranks
     classGroup.shooters.forEach((s, idx) => {
         s.rank = idx + 1; 
     });
     
     boardData.push(classGroup);
  });
  
  // Sort classes alphabetically
  boardData.sort((a, b) => a.className.localeCompare(b.className));
  
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    weekLabel: activeWeekRaw,
    data: boardData
  })).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------
// THE LEADERBOARD ENGINE
// ---------------------------------------------------------

function syncLeagueStandings(silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName("Live Config");
  if (!configSheet) {
    if (!silent) SpreadsheetApp.getUi().alert("Live Config tab not found. Please create it to map out your league weeks.");
    return;
  }
  
  const dropLowestToggle = configSheet.getRange("G1").getValue().toString().trim().toUpperCase();
  let numDrops = parseInt(configSheet.getRange("G2").getValue());
  if (isNaN(numDrops) || dropLowestToggle !== "YES") numDrops = 0;
  
  const configData = configSheet.getDataRange().getValues();
  const activeWeeks = []; 
  const weekNames = [];
  
  configData.forEach(row => {
    const weekName = row[0] ? row[0].toString().trim() : "";
    const pin = row[1] ? row[1].toString().trim() : "";
    const type = row[2] ? row[2].toString().trim().toLowerCase() : "";
    const status = row[3] ? row[3].toString().trim() : "";
    
    if (weekName && pin && (status.toUpperCase() === "ACTIVE" || status.toUpperCase() === "COMPLETED")) {
      activeWeeks.push({ weekName: weekName, pin: pin, type: type });
      weekNames.push(weekName);
    }
  });
  
  if (activeWeeks.length === 0) {
    if (!silent) SpreadsheetApp.getUi().alert("No Active or Completed weeks found in Live Config.");
    return;
  }
  
  let rawSheet = getScoresSheet();
  const rawData = rawSheet.getDataRange().getValues();
  
  const shooterStats = {};
  const pinMap = {};
  activeWeeks.forEach(w => pinMap[w.pin.toString()] = { weekName: w.weekName, type: w.type });
  
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const pin = row[1] ? row[1].toString().trim() : "";
    
    if (!pinMap[pin]) continue; 
    
    const weekName = pinMap[pin].weekName;
    const type = pinMap[pin].type;
    const name = toTitleCase(row[2].toString().trim());
    const className = row[3].toString().trim();
    const endNumber = row[4].toString().trim();
    const arrowsStr = row[5].toString().trim();
    const endTotal = parseInt(row[6]) || 0;
    
    if (!shooterStats[className]) shooterStats[className] = {};
    if (!shooterStats[className][name]) shooterStats[className][name] = { weeks: {}, grandTotal: 0, tb1: 0, tb2: 0 };
    if (!shooterStats[className][name].weeks[weekName]) shooterStats[className][name].weeks[weekName] = { score: 0, tb1: 0, tb2: 0 };
    
    const currentWeekly = shooterStats[className][name].weeks[weekName];
    
    if (endNumber === "FINAL") {
        // Skip
    } else if (endNumber === "PAPER") {
        currentWeekly.score += endTotal;
        let s1 = 0, s2 = 0;
        if (arrowsStr.includes('|')) {
           const parts = arrowsStr.split('|');
           s1 = parseInt(parts[0].split(':')[1]) || 0;
           s2 = parseInt(parts[1].split(':')[1]) || 0;
        }
        currentWeekly.tb1 += s1;
        currentWeekly.tb2 += s2;
    } else {
        currentWeekly.score += endTotal;
        
        if (arrowsStr) {
            const allArrows = arrowsStr.split(',');
            allArrows.forEach(a => {
                const valStr = a.trim().toUpperCase();
                const valNum = Number(valStr) || 0;
                
                if (type.includes('asa_champ') || type.includes('champ')) { if (valStr === '14') currentWeekly.tb1++; else if (valStr === '12') currentWeekly.tb2++; }
                else if (type.includes('asa')) { if (valStr === '12') currentWeekly.tb1++; else if (valStr === '10') currentWeekly.tb2++; }
                else if (type.includes('ibo')) { if (valStr === '11') currentWeekly.tb1++; else if (valStr === '10') currentWeekly.tb2++; }
                else if (type.includes('vegas')) { if (valStr === 'X' || valStr === '11') currentWeekly.tb1++; else if (valStr === '10') currentWeekly.tb2++; }
                else if (type.includes('nfaa')) { if (valStr === 'X' || valStr === '6') currentWeekly.tb1++; else if (valStr === '5') currentWeekly.tb2++; }
                else if (type.includes('fred_bear')) { 
                    if (valNum === 5 || valNum === 11 || valNum === 10 || valStr === 'X') currentWeekly.tb1++;
                    else if (valNum === -5) currentWeekly.tb2++;
                }
                else if (type.includes('custom')) {
                    if (valStr === 'X' || valStr === '14' || valStr === '12' || valStr === '11') currentWeekly.tb1++;
                    else if (valStr === '10' || valStr === '5') currentWeekly.tb2++;
                }
            });
        }
    }
  }
  
  const overallArray = [];
  const weeklyArray = [];
  
  Object.keys(shooterStats).forEach(className => {
    const shootersInClass = shooterStats[className];
    
    Object.keys(shootersInClass).forEach(name => {
       const sData = shootersInClass[name];
       let scoresList = [];
       
       activeWeeks.forEach(w => {
           if (sData.weeks[w.weekName]) {
               scoresList.push(sData.weeks[w.weekName].score);
               sData.tb1 += sData.weeks[w.weekName].tb1;
               sData.tb2 += sData.weeks[w.weekName].tb2;
               
               weeklyArray.push({
                   week: w.weekName,
                   type: w.type,
                   className: className,
                   name: name,
                   score: sData.weeks[w.weekName].score,
                   tb1: sData.weeks[w.weekName].tb1,
                   tb2: sData.weeks[w.weekName].tb2
               });
           }
       });
       
       scoresList.sort((a,b) => a - b);
       let droppedScoreDisplay = "";
       
       if (numDrops > 0 && scoresList.length > numDrops) {
           for (let d = 0; d < numDrops; d++) {
               const dropped = scoresList.shift(); 
               droppedScoreDisplay += dropped + " ";
           }
       }
       
       sData.grandTotal = scoresList.reduce((a,b) => a+b, 0);
       
       const overallRow = {
           className: className,
           name: name,
           grandTotal: sData.grandTotal,
           tb1: sData.tb1,
           tb2: sData.tb2,
           dropped: droppedScoreDisplay.trim()
       };
       
       activeWeeks.forEach(w => {
           overallRow[w.weekName] = sData.weeks[w.weekName] ? sData.weeks[w.weekName].score : 0;
       });
       
       overallArray.push(overallRow);
    });
  });
  
  // ----------------------------------------------------
  // Output to Weekly Standings Tab
  // ----------------------------------------------------
  let weeklySheet = ss.getSheetByName("Weekly Standings");
  if (!weeklySheet) weeklySheet = ss.insertSheet("Weekly Standings");
  weeklySheet.clear();
  
  const weeklyHeaders = ["Week Number", "Class", "Weekly Rank", "Shooter Name", "Weekly Score", "TB-1", "TB-2"];
  const weeklyOutput = [weeklyHeaders];
  
  weeklyArray.sort((a, b) => {
      if (a.week !== b.week) return a.week.localeCompare(b.week);
      if (a.className !== b.className) return a.className.localeCompare(b.className);
      if (b.score !== a.score) return b.score - a.score;
      if (b.tb1 !== a.tb1) return b.tb1 - a.tb1;
      return b.tb2 - a.tb2;
  });
  
  let currWeek = "";
  let currClassW = "";
  let currRankW = 1;
  let weekHeaderRowIndices = [];
  
  weeklyArray.forEach(w => {
      if (w.week !== currWeek) {
          if (currWeek !== "") {
              weeklyOutput.push(Array(weeklyHeaders.length).fill(""));
          }
          currWeek = w.week;
          currClassW = ""; 
          
          let prettyType = w.type.replace(/_/g, ' ').toUpperCase();
          weeklyOutput.push([`${w.week.toUpperCase()} - ${prettyType} ROUND`, "", "", "", "", "", ""]);
          weekHeaderRowIndices.push(weeklyOutput.length); 
      }

      if (w.className !== currClassW) {
          if (currClassW !== "") {
              weeklyOutput.push(Array(weeklyHeaders.length).fill(""));
          }
          currClassW = w.className;
          currRankW = 1;
      } else {
          currRankW++;
      }
      weeklyOutput.push([w.week, w.className, currRankW, w.name, w.score, w.tb1, w.tb2]);
  });
  
  if (weeklyOutput.length > 1) {
      weeklySheet.getRange(1, 1, weeklyOutput.length, weeklyOutput[0].length).setValues(weeklyOutput);
      
      // Turn on text wrapping for the header so long titles stack, saving space
      weeklySheet.getRange(1, 1, 1, weeklyOutput[0].length).setFontWeight("bold").setBackground("#d9ead3").setWrap(true);
      
      // Control spacing between Rank and Name
      weeklySheet.setColumnWidth(3, 75); // Narrows the Rank column
      weeklySheet.getRange("C:C").setHorizontalAlignment("center");
      weeklySheet.getRange("D:D").setHorizontalAlignment("left");
      
      weekHeaderRowIndices.forEach(rowIndex => {
          weeklySheet.getRange(rowIndex, 1, 1, weeklyOutput[0].length).setFontWeight("bold").setBackground("#f3f4f6");
      });
  }
  
  // ----------------------------------------------------
  // Output to Overall Leaderboard Tab
  // ----------------------------------------------------
  let overallSheet = ss.getSheetByName("Overall Leaderboard");
  if (!overallSheet) overallSheet = ss.insertSheet("Overall Leaderboard");
  overallSheet.clear();
  
  // Headers modified to match the requested wording
  const overallHeaders = ["Class", "Overall Rank", "Shooter Name", "Grand Total Score", "Total TB-1", "Total TB-2", "Dropped Score"];
  activeWeeks.forEach(w => overallHeaders.push(w.weekName + " Score"));
  
  const overallOutput = [overallHeaders];
  
  overallArray.sort((a, b) => {
      if (a.className !== b.className) return a.className.localeCompare(b.className);
      if (b.grandTotal !== a.grandTotal) return b.grandTotal - a.grandTotal;
      if (b.tb1 !== a.tb1) return b.tb1 - a.tb1;
      return b.tb2 - a.tb2;
  });
  
  let currClassO = "";
  let currRankO = 1;
  
  overallArray.forEach(o => {
      if (o.className !== currClassO) {
          if (currClassO !== "") {
              overallOutput.push(Array(overallHeaders.length).fill(""));
          }
          currClassO = o.className;
          currRankO = 1;
      } else {
          currRankO++;
      }
      const row = [o.className, currRankO, o.name, o.grandTotal, o.tb1, o.tb2, o.dropped];
      activeWeeks.forEach(w => row.push(o[w.weekName]));
      overallOutput.push(row);
  });
  
  if (overallOutput.length > 1) {
      overallSheet.getRange(1, 1, overallOutput.length, overallOutput[0].length).setValues(overallOutput);
      
      // Turn on text wrapping for the header so long titles stack, saving space
      overallSheet.getRange(1, 1, 1, overallOutput[0].length).setFontWeight("bold").setBackground("#cfe2f3").setWrap(true);
      
      // Control spacing between Rank and Name
      overallSheet.setColumnWidth(2, 75); // Narrows the Rank column
      overallSheet.getRange("B:B").setHorizontalAlignment("center");
      overallSheet.getRange("C:C").setHorizontalAlignment("left");
  }
  
  if (!silent) SpreadsheetApp.getUi().alert("League Standings Synced Successfully!");
}

// ---------------------------------------------------------
// SCORECARD VIEWER (SPREADSHEET BASED)
// ---------------------------------------------------------

function setupScorecardViewer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Scorecard Viewer");
  if (!sheet) {
    sheet = ss.insertSheet("Scorecard Viewer", 1); 
  }
  sheet.clear();
  
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  
  sheet.getRange("A1:D1").merge().setValue("🏆 LEAGUE SCORECARD VIEWER")
       .setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center")
       .setBackground("#1e293b").setFontColor("white");
  
  sheet.getRange("A2").setValue("Select Week (PIN):").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("A3").setValue("Select Shooter:").setFontWeight("bold").setHorizontalAlignment("right");
  
  sheet.getRange("B2:B3").setBackground("#f8fafc").setBorder(true, true, true, true, false, false);
  
  const configSheet = ss.getSheetByName("Live Config");
  let pins = [];
  if (configSheet) {
    const configData = configSheet.getDataRange().getValues();
    configData.forEach(row => {
      if (row[0] && row[1] && row[1].toString().toUpperCase() !== "PIN") {
          pins.push(row[1].toString().trim());
      }
    });
  }
  
  if (pins.length > 0) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(pins).build();
    sheet.getRange("B2").setDataValidation(rule);
  } else {
    sheet.getRange("B2").setValue("No PINs in Live Config");
  }
  
  sheet.getRange("A5:D5").setValues([["Target / End", "Arrow Values", "End Total", "Running Total"]])
       .setFontWeight("bold").setBackground("#cbd5e1").setHorizontalAlignment("center");
  
  SpreadsheetApp.getUi().alert("Scorecard Viewer tab created!\n\n1. Select a PIN in cell B2.\n2. The script will automatically load the shooters who shot that week into cell B3.\n3. Select a shooter to view their card!");
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Scorecard Viewer") return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  if (row === 2 && col === 2) {
    const pin = e.value;
    sheet.getRange("B3").clearContent().clearDataValidations();
    sheet.getRange("A6:D100").clearContent().setBackground(null);
    
    if (!pin) return;
    
    const rawSheet = e.source.getSheetByName("Raw Scores");
    if (!rawSheet) return;
    
    const rawData = rawSheet.getDataRange().getValues();
    let shooters = new Set();
    
    for (let i = 0; i < rawData.length; i++) {
      if (rawData[i][1] && rawData[i][1].toString() === pin.toString()) {
        shooters.add(toTitleCase(rawData[i][2].toString().trim()));
      }
    }
    
    const shooterArray = Array.from(shooters).sort();
    if (shooterArray.length > 0) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(shooterArray).build();
      sheet.getRange("B3").setDataValidation(rule);
      sheet.getRange("B3").setValue(shooterArray[0]); 
      drawScorecard(sheet, pin, shooterArray[0]);
    } else {
      sheet.getRange("B3").setValue("No scores for this PIN");
    }
  }
  
  if (row === 3 && col === 2) {
    const pin = sheet.getRange("B2").getValue().toString().trim();
    const shooter = e.value;
    if (pin && shooter) {
      drawScorecard(sheet, pin, shooter);
    }
  }
}

function drawScorecard(viewerSheet, pin, shooterName) {
  viewerSheet.getRange("A6:D100").clearContent().setBackground(null);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Raw Scores");
  if (!rawSheet) return;
  
  const rawData = rawSheet.getDataRange().getValues();
  let output = [];
  
  for (let i = 0; i < rawData.length; i++) {
    if (rawData[i][1] && rawData[i][1].toString() === pin.toString() && toTitleCase(rawData[i][2].toString().trim()) === shooterName.toString().trim()) {
      const endNum = rawData[i][4].toString();
      if (endNum === "FINAL") continue; 
      
      const arrows = rawData[i][5].toString();
      const endTotal = rawData[i][6];
      const runTotal = rawData[i][7];
      
      output.push([endNum, arrows, endTotal, runTotal]);
    }
  }
  
  if (output.length > 0) {
    viewerSheet.getRange(6, 1, output.length, 4).setValues(output);
    viewerSheet.getRange(6, 1, output.length, 4).setHorizontalAlignment("center");
    
    for (let r = 0; r < output.length; r++) {
       if (r % 2 === 0) viewerSheet.getRange(6 + r, 1, 1, 4).setBackground("#f1f5f9");
    }
  } else {
    viewerSheet.getRange("A6").setValue("No scorecard found.");
  }
}