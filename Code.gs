// HYROX 12-Week Training Logger — Google Apps Script Backend
// Paste this into Extensions → Apps Script in your Google Sheet

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Gautam's HYROX Logger [v2.25.04]")
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
      [1,1,'Recovery','Neural KB','Easy','Technique Walk-Through','Easy Run','TGU + Accessory','6km easy',0,'',''],
      [2,1,'Recovery','Neural KB','Easy','Light Station Work','Easy Run','TGU + Accessory','6km easy',0,'',''],
      [3,2,'Base','Neural KB','Tempo A1 @ 9.0 km/h','Station Work - Sled Pull Focus','Easy Run','TGU + Accessory','8km easy',0,'',''],
      [4,2,'Base','Neural KB','Tempo A2 @ 9.0 km/h','Even-Effort Brick A','Easy Run','TGU + Accessory','9km easy',0,'',''],
      [5,2,'Base','Neural KB','Tempo B1 @ 9.5 km/h','Benchmark Day (SkiErg + Row TT)','Easy Run','TGU + Accessory','10km easy',0,'',''],
      [6,3,'Threshold','Neural KB + TM Finish','Tempo B2 @ 9.5 km/h','Even-Effort Brick B','Easy Run','TGU + Accessory + Incline','10km easy',0,'',''],
      [7,3,'Threshold','Neural KB + TM Finish','Threshold C1 @ 9.5 km/h','Full Station Rotation','Easy Run','TGU + Accessory + Incline','11km easy',0,'',''],
      [8,3,'Threshold','Neural KB + TM Finish','Intervals @ 10.5 km/h','Hard Brick - Race Sequence','Easy Run','TGU + Accessory + Incline','10km easy',0,'',''],
      [9,3,'Threshold','Neural KB + TM Finish','Threshold Peak @ 10.0 km/h','Benchmark Retest','Easy Run','TGU + Accessory + Incline','12km easy',0,'',''],
      [10,4,'Peak/Sim','Neural KB + TM Finish','Race Sharp @ 10.5 km/h','FULL SIMULATION','Easy Run','TGU + Accessory + Incline','7km easy',0,'',''],
      [11,4,'Taper','Neural KB + TM Finish','Strides @ 11.0 km/h','Light Station Touch-Up','Easy Run','TGU + Accessory + Incline','7km easy',0,'',''],
      [12,4,'Race Week','Neural KB + TM Finish','Race Feel @ 11.0 km/h','Final Light Touch','Easy Run','TGU + Accessory + Incline','RACE DAY',0,'','']
    ];
    dash.getRange(2, 1, dashData.length, dashData[0].length).setValues(dashData);
    dash.getRange(1, 1, 1, 12).setFontWeight('bold');
  }

  // RunLog
  var run = getOrCreateSheet(ss, 'RunLog');
  if (run.getLastRow() < 2) {
    run.getRange(1, 1, 1, 16).setValues([['Week','Day','Date','SessionType','Role','PlannedDist','BeltSpeed','TargetPace','ActualDist','ActualTime','ActualPace','AvgHR','MaxHR','HRCapHeld','EffortFeel','Notes']]);
    run.getRange(1, 1, 1, 16).setFontWeight('bold');
    var runData = buildRunLogData();
    if (runData.length > 0) {
      run.getRange(2, 1, runData.length, runData[0].length).setValues(runData);
      // Format TargetPace column (H) and ActualPace column (K) as plain text
      run.getRange(2, 8, runData.length, 1).setNumberFormat('@');
      run.getRange(2, 11, runData.length, 1).setNumberFormat('@');
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
    station.getRange(1, 1, 1, 14).setValues([['Week','Date','SessionType','Role','SkiErg','SledPush','SledPull','BurpeeBJ','Row','Farmers','Sandbag','WallBalls','SessionFeel','Notes']]);
    station.getRange(1, 1, 1, 14).setFontWeight('bold');
    var stationData = buildStationLogData();
    if (stationData.length > 0) station.getRange(2, 1, stationData.length, stationData[0].length).setValues(stationData);
  }

  // Benchmarks
  var bench = getOrCreateSheet(ss, 'Benchmarks');
  if (bench.getLastRow() < 2) {
    bench.getRange(1, 1, 1, 8).setValues([['Benchmark','Date','SkiErg500m','Row500m','SledPullLoad','WallBallMaxUnbroken','OverallFeel','Notes']]);
    bench.getRange(1, 1, 1, 8).setFontWeight('bold');
    bench.getRange(2, 1, 2, 1).setValues([['Week 5 - Baseline'],['Week 9 - Retest']]);
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
  var rows = [];
  var weeks = [
    // [week, tuType, tuRole, tuDist, tuBelt, tuPace, thuDist, satDist]
    [1,'Easy','HIGH COST','4-5km easy','','','4-5km easy','6km easy'],
    [2,'Easy','HIGH COST','5km easy','','','5km easy','6km easy'],
    [3,'Tempo A1','HIGH COST','2x10min',9.0,'6:40','6km easy','8km easy'],
    [4,'Tempo A2','HIGH COST','3x10min',9.0,'6:40','6km easy','9km easy'],
    [5,'Tempo B1','HIGH COST','3x10min',9.5,'6:19','6km easy','10km easy'],
    [6,'Tempo B2','HIGH COST','4x8min',9.5,'6:19','6km easy','10km easy'],
    [7,'Threshold C1','HIGH COST','20min continuous',9.5,'6:19','6km easy','11km easy'],
    [8,'Intervals','HIGH COST','6x1km',10.5,'5:43','6km easy','10km easy'],
    [9,'Threshold Peak','HIGH COST','30min continuous',10.0,'6:00','6km easy','12km easy'],
    [10,'Race Sharp','HIGH COST','4x1km',10.5,'5:43','6km easy','7km easy'],
    [11,'Strides','HIGH COST','Easy + 3x200m strides',11.0,'5:27','6km easy','7km easy'],
    [12,'Race Feel','HIGH COST','Easy + 3x400m',11.0,'5:27','6km easy','3km easy']
  ];
  weeks.forEach(function(w) {
    rows.push([w[0],'Tuesday','',w[1],w[2],w[3],w[4],w[5],'','','','','','','','']);
    rows.push([w[0],'Thursday','','Easy Run','RECOVERY',w[6],'','','','','','','','','','']);
    rows.push([w[0],'Saturday','','Long Easy Run','RECOVERY',w[7],'','','','','','','','','','']);
  });
  return rows;
}

function buildKBLogData() {
  var rows = [];
  var weeks = [
    // Week 1
    {w:1, mon:[['KB Deadlift','3x5x28-32kg'],['2H Swing','10x10x16kg'],['KB Press','3x6x16-20kg'],['KB Row','3x6x20-24kg']],
           fri:[['TGU','3x1e x 8kg'],['KB Press','3x10x16kg'],['KB Row','3x10x20kg'],['KB Carry','2x30m x 16-20kg']]},
    // Week 2
    {w:2, mon:[['KB Deadlift','3x5x32kg'],['2H Swing','10x10x16kg'],['KB Press','3x6x20kg'],['KB Row','3x6x24kg']],
           fri:[['TGU','3x1e x 10-12kg'],['KB Press','3x10x16-20kg'],['KB Row','3x10x20-24kg'],['KB Carry','2x30m x 16-20kg']]},
    // Week 3
    {w:3, mon:[['KB Deadlift','3x5x32-36kg'],['2H Swing','10x10x24kg'],['KB Press','3x6x20-24kg'],['KB Row','3x6x24-28kg']],
           fri:[['TGU','3x2e x 12-16kg'],['KB Press','3x10x16-20kg'],['KB Row','3x12x20-24kg'],['KB Carry','3x30m x 16-20kg']]},
    // Week 4
    {w:4, mon:[['KB Deadlift','3x5x32-36kg'],['2H Swing','10x10x24kg'],['KB Press','3x6x20-24kg'],['KB Row','3x6x24-28kg']],
           fri:[['TGU','3x2e x 14-16kg'],['KB Press','3x10x20kg'],['KB Row','3x12x24kg'],['KB Carry','3x30m x 20-24kg']]},
    // Week 5
    {w:5, mon:[['KB Deadlift','3x5x36kg'],['2H Swing','10x10x24kg'],['KB Press','3x6x24kg'],['KB Row','3x6x28-32kg']],
           fri:[['TGU','3x3e x 16-20kg'],['KB Press','3x10x20kg'],['KB Row','3x12x24kg'],['KB Windmill','2x5e x 12kg']]},
    // Week 6
    {w:6, mon:[['2H Swing (warm-up)','2x8x24kg'],['1H Swing','5x5e x 24kg'],['KB Press','3x8x24kg'],['KB Row','3x8x28-32kg'],['TM Finish','5min @ 8.0km/h']],
           fri:[['TGU','3x3e x 16-20kg'],['KB Press','3x8x20-24kg'],['KB Row','3x10x24-28kg'],['KB Carry','2x30m x 20-24kg'],['Incline TM','3x3min @ 10-12%']]},
    // Week 7
    {w:7, mon:[['2H Swing (warm-up)','2x8x24kg'],['1H Swing','5x5e x 24kg'],['KB Press','3x8x24kg'],['KB Row','3x8x28-32kg'],['TM Finish','5min @ 8.0km/h']],
           fri:[['TGU','3x3e x 20-24kg'],['KB Press','3x8x24kg'],['KB Row','3x10x28-32kg'],['KB Carry','3x30m x 24kg'],['Incline TM','4x3min @ 10-12%']]},
    // Week 8
    {w:8, mon:[['2H Swing','5x8x32kg'],['Clean','3x3e x 24kg'],['KB Press','3x8x24-28kg'],['KB Row','3x8x32-36kg'],['TM Finish','5min @ 8.0km/h']],
           fri:[['TGU','3x3e x 20-24kg'],['KB Press','3x8x20kg'],['KB Row','3x10x24kg'],['Incline TM','3x3min @ 10-12%']]},
    // Week 9
    {w:9, mon:[['2H Swing','5x8x32kg'],['Clean','3x3e x 24kg'],['KB Press','3x8x24-28kg'],['KB Row','3x8x32-36kg'],['TM Finish','5min @ 8.2km/h']],
           fri:[['TGU','3x3e x 24kg'],['KB Press','3x8x20-24kg'],['KB Row','3x10x24-28kg'],['Incline TM','3x3min @ 10-12%']]},
    // Week 10
    {w:10, mon:[['2H Swing','3x8x28kg'],['Clean','3x3e x 24kg'],['KB Press','3x6x24kg'],['KB Row','3x6x28-32kg'],['TM Finish','5min @ 7.8km/h']],
            fri:[['TGU','3x3e x 24kg'],['KB Press','2x10x16-20kg'],['KB Row','2x10x20-24kg'],['Incline TM','2x3min @ 10%']]},
    // Week 11
    {w:11, mon:[['2H Swing','2x8x24kg'],['Clean','2x3e x 24kg'],['KB Press','2x6x20kg'],['KB Row','2x6x24-28kg'],['TM Finish','3min @ 7.8km/h']],
            fri:[['TGU','2x2e x 20kg'],['KB Press','2x8x16kg'],['KB Row','2x10x20kg'],['Incline TM','1x3min @ 10%']]},
    // Week 12
    {w:12, mon:[['2H Swing','2x5x16-20kg'],['TGU check','1x1e x 12kg'],['KB Press','2x6x16kg']],
            fri:[['TGU (wake-up)','1x1e x 12kg'],['KB Press','2x6x16kg']]}
  ];

  weeks.forEach(function(wk) {
    var label = wk.w <= 2 ? 'Recovery' : wk.w <= 5 ? 'Base' : wk.w <= 9 ? 'Threshold' : wk.w <= 10 ? 'Peak/Sim' : wk.w === 11 ? 'Taper' : 'Race Week';
    wk.mon.forEach(function(m, i) {
      rows.push([wk.w, 'Monday', '', 'SUPPORT', label + ' - Neural KB', m[0], m[1], '', '', '', '', '', '', '', '', '']);
    });
    wk.fri.forEach(function(m, i) {
      rows.push([wk.w, 'Friday', '', 'ACCESSORY', label + ' - Yin/TGU', m[0], m[1], '', '', '', '', '', '', '', '', '']);
    });
  });
  return rows;
}

function buildStationLogData() {
  var rows = [];
  var types = [
    [1,'Technique Walk-Through','SUPPORT'],
    [2,'Light Station Work','SUPPORT'],
    [3,'Station Work - Sled Pull Focus','SUPPORT'],
    [4,'Even-Effort Brick A','SUPPORT'],
    [5,'Benchmark Day (SkiErg + Row TT)','HIGH COST'],
    [6,'Even-Effort Brick B','SUPPORT'],
    [7,'Full Station Rotation - Half Volume','HIGH COST'],
    [8,'Hard Brick - Race Sequence','HIGH COST'],
    [9,'Benchmark Retest','HIGH COST'],
    [10,'FULL SIMULATION with Kabeer','SIM'],
    [11,'Light Station Touch-Up','SUPPORT'],
    [12,'Final Light Touch','SUPPORT']
  ];
  types.forEach(function(t) {
    rows.push([t[0],'',t[1],t[2],'','','','','','','','','','']);
  });
  return rows;
}

// ─── API Functions (called from frontend) ──────────────────────────────────────

function sanitizeRow(row) {
  for (var key in row) {
    var v = row[key];
    if (v instanceof Date) {
      // Check if it's a time value (year 1899/1900 = Google Sheets time-only)
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

function getWeekData(week) {
  try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = { week: week, phase: '', block: '', sessions: {} };

  // Dashboard
  var dash = ss.getSheetByName('Dashboard');
  if (dash) {
    var dashData = dash.getDataRange().getValues();
    for (var i = 1; i < dashData.length; i++) {
      if (dashData[i][0] == week) {
        result.phase = dashData[i][1];
        result.block = dashData[i][2];
        result.sessions.monLabel = dashData[i][3];
        result.sessions.tueLabel = dashData[i][4];
        result.sessions.wedLabel = dashData[i][5];
        result.sessions.thuLabel = dashData[i][6];
        result.sessions.friLabel = dashData[i][7];
        result.sessions.satLabel = dashData[i][8];
        result.sessionsDone = dashData[i][9];
        break;
      }
    }
  }

  // Run Log
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

  // KB Log
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
        // Ensure ActualSetData is a plain string
        if (!row.ActualSetData || row.ActualSetData === '') row.ActualSetData = '[]';
        row.ActualSetData = String(row.ActualSetData);
        result.kb.push(row);
      }
    }
  }

  // Station Log
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

  // Count runs with data
  var run = ss.getSheetByName('RunLog');
  if (run) {
    var runData = run.getDataRange().getValues();
    for (var i = 1; i < runData.length; i++) {
      if (runData[i][0] == week && (runData[i][8] || runData[i][9])) count++;
    }
  }

  // Count KB days with data (unique days)
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

  // Count station with data
  var station = ss.getSheetByName('StationLog');
  if (station) {
    var stData = station.getDataRange().getValues();
    for (var i = 1; i < stData.length; i++) {
      if (stData[i][0] == week && (stData[i][4] || stData[i][5] || stData[i][6])) count++;
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

function getProgressData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = { tuePace: [], kbWeights: [], sessionsPerWeek: [], satDist: [], tguWeight: [], effortFeel: [] };

  // Tuesday pace
  var run = ss.getSheetByName('RunLog');
  if (run) {
    var runData = run.getDataRange().getValues();
    for (var i = 1; i < runData.length; i++) {
      if (runData[i][1] === 'Tuesday' && runData[i][10]) {
        result.tuePace.push({ week: runData[i][0], pace: runData[i][10], hr: runData[i][11] });
      }
      if (runData[i][1] === 'Saturday' && runData[i][8]) {
        result.satDist.push({ week: runData[i][0], dist: runData[i][8] });
      }
    }
  }

  // KB weights (max swing weight per week)
  var kb = ss.getSheetByName('KBLog');
  if (kb) {
    var kbData = kb.getDataRange().getValues();
    var swingByWeek = {};
    var tguByWeek = {};
    for (var i = 1; i < kbData.length; i++) {
      var w = kbData[i][0];
      var movement = kbData[i][5];
      var actualKg = kbData[i][9];
      if (actualKg) {
        if (movement && (movement.indexOf('Swing') >= 0)) {
          if (!swingByWeek[w] || actualKg > swingByWeek[w]) swingByWeek[w] = actualKg;
        }
        if (movement === 'TGU' || (movement && movement.indexOf('TGU') >= 0)) {
          if (!tguByWeek[w] || actualKg > tguByWeek[w]) tguByWeek[w] = actualKg;
        }
      }
    }
    for (var w in swingByWeek) result.kbWeights.push({ week: parseInt(w), kg: swingByWeek[w] });
    for (var w in tguByWeek) result.tguWeight.push({ week: parseInt(w), kg: tguByWeek[w] });
    result.kbWeights.sort(function(a,b) { return a.week - b.week; });
    result.tguWeight.sort(function(a,b) { return a.week - b.week; });
  }

  // Sessions per week
  var dash = ss.getSheetByName('Dashboard');
  if (dash) {
    var dashData = dash.getDataRange().getValues();
    for (var i = 1; i < dashData.length; i++) {
      if (dashData[i][0]) {
        result.sessionsPerWeek.push({ week: dashData[i][0], count: dashData[i][9] || 0 });
      }
    }
  }

  return result;
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
