// HYROX 13-Week V6.1 Training Logger — Google Apps Script Backend
// Race: July 26, 2026 — Target: Sub 2:00:00
// Paste into Extensions → Apps Script in your Google Sheet

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Gautam's HYROX Logger [v6.1]")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ─── Sheet Setup ───────────────────────────────────────────────────────────────
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Dashboard
  var dash = getOrCreateSheet(ss, 'Dashboard');
  if (dash.getLastRow() < 2) {
    dash.getRange(1, 1, 1, 12).setValues([['Week','Phase','Block','Mon KB','Tue Run','Wed Station','Thu Run','Fri KB','Sat Long','Sessions Done','Overall Feel','Notes']]);
    var dashData = [
      [1,1,'Foundation','KB Foundation','Easy 5.5km','Technique Walk-Through','Easy 5km','TGU + Accessory','Long 10km Easy',0,'',''],
      [2,1,'Foundation','KB Foundation','Easy 6.5km','Light Station Practice','Easy 5km','TGU + Accessory','Long 11km Block',0,'',''],
      [3,2,'Pace Calibration','KB Pood Jump 24kg','200m Repeats CALIBRATION','Sled Pull Focus + Stations','Easy 5km','TGU + Accessory','Long 6km Easy (deload)',0,'',''],
      [4,2,'Swing Ownership','KB 24kg Consolidation','Fast 8-4-2s','Even-Effort Brick A','Easy 6.5km','TGU + Accessory','Long 12km Progressive',0,'',''],
      [5,2,'Swing Ownership','KB 24kg Ownership Check','Rolling 400s','Even-Effort Brick B','Easy 6km','TGU + Windmill','Long 13km Easy',0,'',''],
      [6,3,'1H Swing + Specificity','1H Swing + TM Finish','Easy 6.5km','⭐ Benchmark TT + Sled Pull@Race','Easy 6km','TGU + Accessory (no incline)','Long 14km Progressive Repeat',0,'',''],
      [7,3,'1H Swing + Specificity','1H Swing + TM Finish','400s into 200s','Quarter Rotation + Sled Pull','Easy 6.5km','TGU + Accessory (no incline)','Long 15km Easy (Peak 1)',0,'',''],
      [8,4,'Race-Pattern Strength','Phase 4: Front Squat (-30% vol)','5km TIME TRIAL','⭐ Hard Brick (sled@race)','Easy 5km','TGU + Race-Pattern + Wall Ball Cap','Long 9km Easy (deload)',0,'',''],
      [9,4,'Race-Pattern Strength','Phase 4: Heavy DL + Rack Lunge','Easy 7.5km','Even-Effort Brick C + Sled Pull (Wall Ball Test)','Easy 7.5km','TGU + Race-Pattern (no incline)','Long 13km Easy',0,'',''],
      [10,4,'Race-Pattern Strength','Phase 4: Reduced + TM','Easy 7.5km','⭐ FULL HYROX SIMULATION','Easy 7km','TGU + Light Race-Pattern (no incline)','Long 14km Progressive',0,'',''],
      [11,4,'Last Hard Week','KB Taper + TM','Drop Set 7km','Light Touch-Up + Light Sled','Easy 6km','TGU Light (no incline)','Long 15km Easy (Peak 2)',0,'',''],
      [12,4,'Pre-Taper','KB Pre-Taper','Easy 5km','Very Light Touch','Easy 5km','TGU Minimal','Long 8km Easy',0,'',''],
      [13,4,'Race Week','Race Wake-Up','Easy + 3x400m strides','Race Touch','REST','Final Touch','🏁 RACE Jul 26',0,'','']
    ];
    dash.getRange(2, 1, dashData.length, dashData[0].length).setValues(dashData);
    dash.getRange(1, 1, 1, 12).setFontWeight('bold');
  }

  // RunLog
  var run = getOrCreateSheet(ss, 'RunLog');
  if (run.getLastRow() < 2) {
    run.getRange(1, 1, 1, 16).setValues([['Week','Day','Date','SessionType','Role','PlannedDist','TargetPace','HRCap','ActualDist','ActualTime','ActualPace','AvgHR','MaxHR','HRCapHeld','EffortFeel','Notes']]);
    run.getRange(1, 1, 1, 16).setFontWeight('bold');
    var runData = buildRunLogData();
    if (runData.length > 0) {
      run.getRange(2, 1, runData.length, runData[0].length).setValues(runData);
      run.getRange(2, 7, runData.length, 1).setNumberFormat('@'); // TargetPace as text
      run.getRange(2, 11, runData.length, 1).setNumberFormat('@'); // ActualPace as text
    }
  }

  // KBLog
  var kb = getOrCreateSheet(ss, 'KBLog');
  if (kb.getLastRow() < 2) {
    kb.getRange(1, 1, 1, 16).setValues([['Week','Day','Date','Role','SessionLabel','Movement','PlannedSetsRepsKg','ActualSets','ActualReps','ActualKg','ActualSetData','HRCapHeld','FloatCheck','FormBreak','EffortFeel','Notes']]);
    kb.getRange(1, 1, 1, 16).setFontWeight('bold');
    var kbData = buildKBLogData();
    if (kbData.length > 0) kb.getRange(2, 1, kbData.length, kbData[0].length).setValues(kbData);
  }

  // StationLog
  var station = getOrCreateSheet(ss, 'StationLog');
  if (station.getLastRow() < 2) {
    station.getRange(1, 1, 1, 18).setValues([['Week','Date','SessionType','Role','SkiErg','SledPush','SledPull','BurpeeBJ','Row','Farmers','Sandbag','WallBalls','SledPullRaceLoad','SledPullTimes','TemplateHR','EvenEffortHeld','SessionFeel','Notes']]);
    station.getRange(1, 1, 1, 18).setFontWeight('bold');
    var stationData = buildStationLogData();
    if (stationData.length > 0) station.getRange(2, 1, stationData.length, stationData[0].length).setValues(stationData);
  }

  // Benchmarks
  var bench = getOrCreateSheet(ss, 'Benchmarks');
  if (bench.getLastRow() < 2) {
    bench.getRange(1, 1, 1, 9).setValues([['Benchmark','Date','SkiErg500m','Row500m','5kmTT','SledPullRace4x25m','WallBallTest','OverallFeel','Notes']]);
    bench.getRange(1, 1, 1, 9).setFontWeight('bold');
    bench.getRange(2, 1, 4, 1).setValues([['W6 Benchmark TT'],['W8 5km TT'],['W9 Wall Ball Capacity Test'],['W10 Full Simulation']]);
  }

  // HRVLog (NEW)
  var hrv = getOrCreateSheet(ss, 'HRVLog');
  if (hrv.getLastRow() < 1) {
    hrv.getRange(1, 1, 1, 6).setValues([['Date','HRV','OuraReadiness','RestingHR','SleepScore','Notes']]);
    hrv.getRange(1, 1, 1, 6).setFontWeight('bold');
  }

  return 'Setup complete';
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// ─── Data Builders ─────────────────────────────────────────────────────────────
function buildRunLogData() {
  // Tuesday paces are EXPECTED scenario (V6.1 fallback table default)
  // [week, sessionType, role, plannedDist, targetPace, hrCap]
  var tueData = {
    1:  ['Easy Run',         'RECOVERY',  '5.5km easy',                     '',                          148],
    2:  ['Easy Run',         'RECOVERY',  '6.5km easy',                     '',                          148],
    3:  ['200m Repeats',     'CALIBRATION','6×200m + 200m jog (4.5km)',     '5:00-5:30/km (calibration)', ''],
    4:  ['Fast 8-4-2s',      'HIGH COST', '8min/4min/2min descending (5km)','8min:6:30-7:00 / 4min:6:00-6:30 / 2min:5:30-6:00', ''],
    5:  ['Rolling 400s',     'HIGH COST', '6×400m + 200m jog (6km)',        '5:30-6:00/km',              ''],
    6:  ['Easy Run',         'RECOVERY',  '6.5km easy',                     '',                          148],
    7:  ['400s into 200s',   'HIGH COST', '4×(400m+200m) (6km)',            '400s:5:30-6:00 / 200s:5:00-5:30', ''],
    8:  ['5km TIME TRIAL',   'HIGH COST', '5km max sustainable',            'sub 5:30/km avg',           ''],
    9:  ['Easy Run',         'RECOVERY',  '7.5km easy',                     '',                          150],
    10: ['Easy Run',         'RECOVERY',  '7.5km easy',                     '',                          148],
    11: ['Drop Set',         'HIGH COST', '1.5/1.5/1.5km (7km total)',      '6:15 / 6:30 / 7:00 EXPECTED', ''],
    12: ['Easy Run',         'RECOVERY',  '5km easy',                       '',                          145],
    13: ['Race Pace Feel',   'RECOVERY',  '3km easy + 3×400m + 1km cool',   '5:30-6:00/km on 400s',      '']
  };
  // Thursday & Saturday distances
  var thuDist = {1:'5km',2:'5km',3:'5km',4:'6.5km',5:'6km',6:'6km',7:'6.5km',8:'5km',9:'7.5km',10:'7km',11:'6km',12:'5km',13:'REST'};
  var satDist = {
    1:  ['Easy',          '10km easy',                                      152],
    2:  ['Block',         '11km (3 easy + 5 steady HR<155 + 3 easy)',       155],
    3:  ['Easy DELOAD',   '6km easy',                                       152],
    4:  ['Progressive',   '12km (4 easy + 4 @6:30-7:00 + 4 @6:00-6:30)',    ''],
    5:  ['Easy',          '13km easy',                                      152],
    6:  ['Progressive Repeat','14km (3e + 2s + 1h, repeat 2x + 2e)',        160],
    7:  ['Easy Peak 1',   '15km easy',                                      155],
    8:  ['Easy DELOAD',   '9km easy',                                       152],
    9:  ['Easy',          '13km easy',                                      152],
    10: ['Progressive',   '14km (5 easy + 5 @6:30-7:00 + 4 @6:00-6:30)',    ''],
    11: ['Easy Peak 2',   '15km easy',                                      155],
    12: ['Easy',          '8km easy',                                       148],
    13: ['🏁 RACE DAY',   'July 26',                                        '']
  };

  var rows = [];
  for (var w = 1; w <= 13; w++) {
    var t = tueData[w];
    rows.push([w,'Tuesday','',t[0],t[1],t[2],t[3],t[4],'','','','','','','','']);
    rows.push([w,'Thursday','','Easy Run','RECOVERY',thuDist[w],'',150,'','','','','','','','']);
    var s = satDist[w];
    rows.push([w,'Saturday','',s[0],'RECOVERY',s[1],'',s[2],'','','','','','','','']);
  }
  return rows;
}

function buildKBLogData() {
  // Each entry: [movement, plannedSetsRepsKg]
  // Phase 1 (W1-2): Foundation — Pavel hinge, 16kg swings
  // Phase 2 (W3-5): Pood jump to 24kg, ownership
  // Phase 3 (W6-7): 1H swing + TM finish
  // Phase 4 (W8-13): Race-pattern strength (no Friday Press in V6.1)
  var weeks = {
    1:  { mon:[['KB Deadlift','3×5 @ 32kg'],['KB Goblet Squat','3×5 @ 24-28kg'],['2H Swing','10×10 @ 16kg'],['KB Press','3×6e @ 16-20kg'],['KB Row','3×6e @ 20-24kg']],
          fri:[['TGU','3×1e @ 8kg'],['KB Row','3×10e @ 20kg'],['KB Carry','2×30m @ 16-20kg']] },
    2:  { mon:[['KB Deadlift','3×5 @ 32kg'],['KB Goblet Squat','3×5 @ 28kg'],['2H Swing','10×10 @ 16kg'],['KB Press','3×6e @ 20kg'],['KB Row','3×6e @ 24kg']],
          fri:[['TGU','3×1e @ 10-12kg'],['KB Row','3×10e @ 20-24kg'],['KB Carry','2×30m @ 16-20kg']] },
    3:  { mon:[['KB Deadlift','3×5 @ 32-36kg'],['KB Goblet Squat','3×5 @ 28-32kg'],['2H Swing','10×10 @ 24kg POOD JUMP'],['KB Press','3×6e @ 20-24kg'],['KB Row','3×6e @ 24-28kg']],
          fri:[['TGU','3×2e @ 12-16kg'],['KB Row','3×12e @ 20-24kg'],['KB Carry','3×30m @ 16-20kg']] },
    4:  { mon:[['KB Deadlift','3×5 @ 32-36kg'],['KB Goblet Squat','3×5 @ 32kg'],['2H Swing','10×10 @ 24kg'],['KB Press','3×6e @ 20-24kg'],['KB Row','3×6e @ 24-28kg']],
          fri:[['TGU','3×2e @ 14-16kg'],['KB Row','3×12e @ 24kg'],['KB Carry','3×30m @ 20-24kg']] },
    5:  { mon:[['KB Deadlift','3×5 @ 36kg'],['KB Goblet Squat','3×5 @ 32kg'],['2H Swing','10×10 @ 24kg ownership'],['KB Press','3×6e @ 24kg'],['KB Row','3×6e @ 28-32kg']],
          fri:[['TGU','3×3e @ 16-20kg'],['KB Row','3×12e @ 24kg'],['KB Windmill','2×5e @ 12kg']] },
    6:  { mon:[['KB Goblet Squat','3×5 @ 32kg'],['2H Swing (warm-up)','2×8 @ 24kg'],['1H Swing','5×5e @ 24kg'],['KB Press','3×8e @ 24kg'],['KB Row','3×8e @ 28-32kg'],['TM Finish','5min @ 8.0km/h']],
          fri:[['TGU','3×3e @ 16-20kg'],['KB Row','3×10e @ 24-28kg'],['KB Carry','2×30m @ 20-24kg']] },
    7:  { mon:[['KB Goblet Squat','3×5 @ 32-36kg'],['2H Swing (warm-up)','2×8 @ 24kg'],['1H Swing','5×5e @ 24kg'],['KB Press','3×8e @ 24kg'],['KB Row','3×8e @ 28-32kg'],['TM Finish','5min @ 8.0km/h']],
          fri:[['TGU','3×3e @ 20-24kg'],['KB Row','3×10e @ 28-32kg'],['KB Carry','3×30m @ 24kg']] },
    8:  { mon:[['Front Squat (double KB)','4×5 @ 24-28kg/hand'],['2H Swing','3×8 @ 24-32kg (32 optional)'],['1-arm Carry','3×40m @ 24-28kg/side'],['KB Press','3×6e @ 20-24kg'],['KB Row','3×6e @ 28-32kg'],['TM Finish','5min @ 8.0km/h']],
          fri:[['TGU','3×3e @ 20-24kg'],['Bulgarian Split Squat','3×6/leg @ 16kg'],['Heavy 1-arm Row','3×8/side @ 22-24kg'],['1-arm Carry','2×30m @ 20-24kg'],['Wall Ball Capacity','4×20 unbroken @ 90s rest']] },
    9:  { mon:[['Heavy Deadlift (double KB)','4×5 @ 28-32kg/hand'],['Rack Lunge','3×10/leg @ 20-24kg/hand'],['2H Swing','3×8 @ 24-28kg'],['KB Press','3×8e @ 24kg'],['KB Row','3×8e @ 28-32kg']],
          fri:[['TGU','3×3e @ 20-24kg PEAK'],['Bulgarian Split Squat','3×8/leg @ 16kg'],['Heavy 1-arm Row','3×10/side @ 24kg'],['1-arm Carry','2×30m @ 24kg']] },
    10: { mon:[['Front Squat (double KB)','3×5 @ 24kg/hand'],['2H Swing','2×8 @ 24kg'],['1-arm Carry','2×40m @ 24kg/side'],['KB Press','2×6e @ 24kg'],['KB Row','2×6e @ 28kg'],['TM Finish','3min @ 7.8km/h']],
          fri:[['TGU','3×3e @ 24kg PEAK'],['Bulgarian Split Squat','2×6/leg @ 16kg'],['Heavy 1-arm Row','2×8/side @ 22kg']] },
    11: { mon:[['Front Squat (double KB)','2×5 @ 20-24kg/hand'],['2H Swing','2×8 @ 20-24kg'],['1-arm Carry','2×30m @ 20kg/side'],['KB Press','2×6e @ 20kg'],['KB Row','2×6e @ 24-28kg'],['TM Finish','3min @ 7.8km/h']],
          fri:[['TGU','2×2e @ 20kg'],['Bulgarian Split Squat','2×6/leg @ 16kg'],['Heavy 1-arm Row','2×8/side @ 20-22kg']] },
    12: { mon:[['2H Swing','2×8 @ 20kg'],['KB Press','2×6e @ 20kg'],['KB Row','2×6e @ 24kg'],['TM Finish','3min @ 7.5km/h']],
          fri:[['TGU','1×2e @ 16kg']] },
    13: { mon:[['2H Swing','2×5 @ 16-20kg'],['TGU check','1×1e @ 12kg'],['KB Press','2×6e @ 16kg'],['TM Finish','3min @ 7.5km/h']],
          fri:[['TGU','1×1e @ 12kg']] }
  };

  var labelMap = { 1:'Foundation',2:'Foundation',3:'Pace Calibration',4:'Swing Ownership',5:'Swing Ownership',6:'1H Swing + Specificity',7:'1H Swing + Specificity',8:'Phase 4 - Race Pattern',9:'Phase 4 - Race Pattern',10:'Phase 4 - Reduced',11:'Last Hard Week',12:'Pre-Taper',13:'Race Week' };

  var rows = [];
  for (var w = 1; w <= 13; w++) {
    var wk = weeks[w];
    var label = labelMap[w];
    wk.mon.forEach(function(m) {
      rows.push([w, 'Monday', '', 'SUPPORT', label + ' - Mon KB', m[0], m[1], '', '', '', '', '', '', '', '', '']);
    });
    wk.fri.forEach(function(m) {
      rows.push([w, 'Friday', '', 'ACCESSORY', label + ' - Fri Yin', m[0], m[1], '', '', '', '', '', '', '', '', '']);
    });
  }
  return rows;
}

function buildStationLogData() {
  // [week, sessionType, role]
  var types = [
    [1, 'Technique Walk-Through',                'RECOVERY'],
    [2, 'Light Station Practice',                'SUPPORT'],
    [3, 'Sled Pull Focus + Stations',            'SUPPORT'],
    [4, 'Even-Effort Brick A',                   'SUPPORT'],
    [5, 'Even-Effort Brick B',                   'SUPPORT'],
    [6, '⭐ Benchmark TT + Sled Pull Race Load', 'HIGH COST'],
    [7, 'Quarter Rotation + Sled Pull',          'SUPPORT'],
    [8, '⭐ Hard Brick - Race Sequence',         'HIGH COST'],
    [9, 'Even-Effort Brick C + Sled Pull (Wall Ball Capacity Test)', 'SUPPORT'],
    [10,'⭐ FULL HYROX SIMULATION with Kabeer',  'SIM'],
    [11,'Light Touch-Up + Light Sled Pull',      'SUPPORT'],
    [12,'Very Light Touch',                      'RECOVERY'],
    [13,'Race Touch',                            'RECOVERY']
  ];
  var rows = [];
  types.forEach(function(t) {
    rows.push([t[0],'',t[1],t[2],'','','','','','','','','','','','','','']);
  });
  return rows;
}

// ─── Helper: sanitize Date objects and time strings ────────────────────────────
function sanitizeRow(row) {
  for (var key in row) {
    var v = row[key];
    if (v instanceof Date) {
      if (v.getFullYear() <= 1900) {
        var h = v.getHours();
        var m = v.getMinutes();
        var s = v.getSeconds();
        if (h > 0) {
          row[key] = h + ':' + (m < 10 ? '0' : '') + m;
        } else {
          row[key] = m + ':' + (s < 10 ? '0' : '') + s;
        }
      } else {
        row[key] = v.toISOString().split('T')[0];
      }
    } else if (v === null || v === undefined) {
      row[key] = '';
    }
  }
  return row;
}

// ─── API: Read week data ───────────────────────────────────────────────────────
function getWeekData(week) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = { week: week, phase: '', block: '', sessions: {} };

    var dash = ss.getSheetByName('Dashboard');
    if (dash) {
      var dashData = dash.getDataRange().getValues();
      for (var i = 1; i < dashData.length; i++) {
        if (dashData[i][0] == week) {
          result.phase = dashData[i][1];
          result.block = dashData[i][2];
          result.sessions.monLabel = String(dashData[i][3] || '');
          result.sessions.tueLabel = String(dashData[i][4] || '');
          result.sessions.wedLabel = String(dashData[i][5] || '');
          result.sessions.thuLabel = String(dashData[i][6] || '');
          result.sessions.friLabel = String(dashData[i][7] || '');
          result.sessions.satLabel = String(dashData[i][8] || '');
          result.sessionsDone = dashData[i][9];
          break;
        }
      }
    }

    var run = ss.getSheetByName('RunLog');
    if (run) {
      var runData = run.getDataRange().getValues();
      var headers = runData[0];
      result.runs = [];
      for (var i = 1; i < runData.length; i++) {
        if (runData[i][0] == week) {
          var row = {};
          for (var j = 0; j < headers.length; j++) row[headers[j]] = runData[i][j];
          row._row = i + 1;
          sanitizeRow(row);
          result.runs.push(row);
        }
      }
    }

    var kb = ss.getSheetByName('KBLog');
    if (kb) {
      var kbData = kb.getDataRange().getValues();
      var kbHeaders = kbData[0];
      result.kb = [];
      for (var i = 1; i < kbData.length; i++) {
        if (kbData[i][0] == week) {
          var row = {};
          for (var j = 0; j < kbHeaders.length; j++) row[kbHeaders[j]] = kbData[i][j];
          row._row = i + 1;
          sanitizeRow(row);
          if (!row.ActualSetData || row.ActualSetData === '') row.ActualSetData = '[]';
          row.ActualSetData = String(row.ActualSetData);
          result.kb.push(row);
        }
      }
    }

    var station = ss.getSheetByName('StationLog');
    if (station) {
      var stData = station.getDataRange().getValues();
      var stHeaders = stData[0];
      result.station = null;
      for (var i = 1; i < stData.length; i++) {
        if (stData[i][0] == week) {
          var row = {};
          for (var j = 0; j < stHeaders.length; j++) row[stHeaders[j]] = stData[i][j];
          row._row = i + 1;
          sanitizeRow(row);
          if (!row.SledPullTimes || row.SledPullTimes === '') row.SledPullTimes = '[]';
          row.SledPullTimes = String(row.SledPullTimes);
          result.station = row;
          break;
        }
      }
    }

    return result;
  } catch(e) {
    return { error: e.message, stack: e.stack, week: week };
  }
}

// ─── API: Save entries ─────────────────────────────────────────────────────────
function saveRunEntry(rowNum, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RunLog');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var key in data) {
    var col = headers.indexOf(key);
    if (col >= 0) sheet.getRange(rowNum, col + 1).setValue(data[key]);
  }
  updateSessionCount(data.Week || sheet.getRange(rowNum, 1).getValue());
  return 'saved';
}

function saveKBEntry(rowNum, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('KBLog');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var key in data) {
    var col = headers.indexOf(key);
    if (col >= 0) sheet.getRange(rowNum, col + 1).setValue(data[key]);
  }
  return 'saved';
}

function saveKBSession(rows) {
  rows.forEach(function(r) { saveKBEntry(r.rowNum, r.data); });
  if (rows.length > 0 && rows[0].data.Week) updateSessionCount(rows[0].data.Week);
  return 'saved';
}

function saveStationEntry(rowNum, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('StationLog');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var key in data) {
    var col = headers.indexOf(key);
    if (col >= 0) sheet.getRange(rowNum, col + 1).setValue(data[key]);
  }
  updateSessionCount(data.Week || sheet.getRange(rowNum, 1).getValue());
  return 'saved';
}

function updateSessionCount(week) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var count = 0;

  var run = ss.getSheetByName('RunLog');
  if (run) {
    var runData = run.getDataRange().getValues();
    for (var i = 1; i < runData.length; i++) {
      if (runData[i][0] == week && (runData[i][8] || runData[i][9])) count++;
    }
  }

  var kb = ss.getSheetByName('KBLog');
  if (kb) {
    var kbData = kb.getDataRange().getValues();
    var kbDays = {};
    for (var i = 1; i < kbData.length; i++) {
      if (kbData[i][0] == week && (kbData[i][7] || kbData[i][8] || kbData[i][9])) {
        kbDays[kbData[i][1]] = true;
      }
    }
    count += Object.keys(kbDays).length;
  }

  var station = ss.getSheetByName('StationLog');
  if (station) {
    var stData = station.getDataRange().getValues();
    for (var i = 1; i < stData.length; i++) {
      if (stData[i][0] == week && (stData[i][4] || stData[i][6] || stData[i][12])) count++;
    }
  }

  var dash = ss.getSheetByName('Dashboard');
  if (dash) {
    var dashData = dash.getDataRange().getValues();
    for (var i = 1; i < dashData.length; i++) {
      if (dashData[i][0] == week) {
        dash.getRange(i + 1, 10).setValue(count);
        break;
      }
    }
  }
}

// ─── API: HRV Daily Log ────────────────────────────────────────────────────────
function saveHRVEntry(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('HRVLog');
  if (!sheet) sheet = setupSheets() && ss.getSheetByName('HRVLog');
  var dataRange = sheet.getDataRange().getValues();
  // Check if entry for this date exists; update or append
  var foundRow = -1;
  for (var i = 1; i < dataRange.length; i++) {
    if (String(dataRange[i][0]).indexOf(data.Date) === 0) { foundRow = i + 1; break; }
  }
  var newRow = [data.Date, data.HRV || '', data.OuraReadiness || '', data.RestingHR || '', data.SleepScore || '', data.Notes || ''];
  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, 6).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
  }
  return 'saved';
}

function getHRVData(limit) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('HRVLog');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var d = data[i][0];
      if (d instanceof Date) d = d.toISOString().split('T')[0];
      result.push({
        Date: String(d),
        HRV: data[i][1] || '',
        OuraReadiness: data[i][2] || '',
        RestingHR: data[i][3] || '',
        SleepScore: data[i][4] || '',
        Notes: String(data[i][5] || '')
      });
    }
    result.sort(function(a, b) { return a.Date < b.Date ? 1 : -1; });
    if (limit) result = result.slice(0, limit);
    return result;
  } catch(e) {
    return [];
  }
}

function getTodayHRV() {
  try {
    var today = new Date().toISOString().split('T')[0];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('HRVLog');
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var d = data[i][0];
      if (d instanceof Date) d = d.toISOString().split('T')[0];
      if (String(d).indexOf(today) === 0) {
        return { Date: today, HRV: data[i][1] || '', OuraReadiness: data[i][2] || '', RestingHR: data[i][3] || '', SleepScore: data[i][4] || '', Notes: String(data[i][5] || '') };
      }
    }
    return null;
  } catch(e) {
    return null;
  }
}

// ─── API: Progress data ────────────────────────────────────────────────────────
function getProgressData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = { tuePace: [], kbWeights: [], sessionsPerWeek: [], satDist: [], tguWeight: [], sledPullProgression: [] };

    var run = ss.getSheetByName('RunLog');
    if (run) {
      var runData = run.getDataRange().getValues();
      for (var i = 1; i < runData.length; i++) {
        if (runData[i][1] === 'Tuesday' && runData[i][10]) {
          result.tuePace.push({ week: runData[i][0], pace: String(runData[i][10]), hr: runData[i][11] });
        }
        if (runData[i][1] === 'Saturday' && runData[i][8]) {
          result.satDist.push({ week: runData[i][0], dist: parseFloat(runData[i][8]) });
        }
      }
    }

    var kb = ss.getSheetByName('KBLog');
    if (kb) {
      var kbData = kb.getDataRange().getValues();
      var swingByWeek = {};
      var tguByWeek = {};
      for (var i = 1; i < kbData.length; i++) {
        var w = kbData[i][0];
        var movement = String(kbData[i][5] || '');
        var actualKg = parseFloat(kbData[i][9]);
        if (actualKg && !isNaN(actualKg)) {
          if (movement.indexOf('Swing') >= 0) {
            if (!swingByWeek[w] || actualKg > swingByWeek[w]) swingByWeek[w] = actualKg;
          }
          if (movement.indexOf('TGU') >= 0) {
            if (!tguByWeek[w] || actualKg > tguByWeek[w]) tguByWeek[w] = actualKg;
          }
        }
      }
      for (var w in swingByWeek) result.kbWeights.push({ week: parseInt(w), kg: swingByWeek[w] });
      for (var w in tguByWeek) result.tguWeight.push({ week: parseInt(w), kg: tguByWeek[w] });
      result.kbWeights.sort(function(a,b) { return a.week - b.week; });
      result.tguWeight.sort(function(a,b) { return a.week - b.week; });
    }

    var dash = ss.getSheetByName('Dashboard');
    if (dash) {
      var dashData = dash.getDataRange().getValues();
      for (var i = 1; i < dashData.length; i++) {
        if (dashData[i][0]) {
          result.sessionsPerWeek.push({ week: dashData[i][0], count: dashData[i][9] || 0 });
        }
      }
    }

    var station = ss.getSheetByName('StationLog');
    if (station) {
      var stData = station.getDataRange().getValues();
      for (var i = 1; i < stData.length; i++) {
        var times = stData[i][13]; // SledPullTimes JSON
        if (times) {
          try {
            var arr = (typeof times === 'string') ? JSON.parse(times) : times;
            if (Array.isArray(arr) && arr.length > 0) {
              var validTimes = arr.map(function(t) { return parseFloat(t); }).filter(function(t) { return !isNaN(t) && t > 0; });
              if (validTimes.length > 0) {
                var avg = validTimes.reduce(function(a,b){return a+b;},0) / validTimes.length;
                result.sledPullProgression.push({ week: stData[i][0], avgTime: avg, count: validTimes.length });
              }
            }
          } catch(e) {}
        }
      }
      result.sledPullProgression.sort(function(a,b) { return a.week - b.week; });
    }

    return result;
  } catch(e) {
    return { error: e.message };
  }
}

function getSpreadsheetId() {
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

function saveBenchmark(rowNum, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Benchmarks');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var key in data) {
    var col = headers.indexOf(key);
    if (col >= 0) sheet.getRange(rowNum, col + 1).setValue(data[key]);
  }
  return 'saved';
}
