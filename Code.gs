const TW = {
  sheets: {
    settings: 'Settings',
    questions: 'Questions',
    responses: 'Responses',
    dashboard: 'Dashboard'
  },
  timezone: 'Asia/Tehran',
  scoreMin: 1,
  scoreMax: 5
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Teamwork')
    .addItem('Setup / Refresh System', 'setupTeamworkSystem')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Show Web App Link', 'showWebAppLink')
    .addToUi();
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Teamwork Review')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupTeamworkSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettingsSheet_(ss);
  ensureQuestionsSheet_(ss);
  ensureResponsesSheet_(ss);
  ensureDashboardSheet_(ss);
  refreshDashboard();
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'The Teamwork review system is ready. Deploy the Apps Script project as a Web app.'
  );
}

function showWebAppLink() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(
    url || 'Deploy the project first from Deploy > New deployment > Web app.'
  );
}

function getBootstrapData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employees = getEmployees_(ss);
  const sections = getSections_(ss);
  const monthKey = Utilities.formatDate(new Date(), TW.timezone, 'yyyy-MM');
  return { employees, sections, monthKey };
}

function getEvaluatorStatus(evaluator, monthKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employees = getEmployees_(ss);

  if (!employees.includes(evaluator)) {
    throw new Error('Invalid evaluator.');
  }

  monthKey = normalizeMonthKey_(monthKey);

  const rows = getResponseRows_(ss);
  const completed = rows
    .filter(r => r.monthKey === monthKey && r.evaluator === evaluator)
    .map(r => r.target);

  const uniqueCompleted = [...new Set(completed)];

  return {
    completed: uniqueCompleted,
    remaining: employees.filter(x => x !== evaluator && !uniqueCompleted.includes(x)),
    totalRequired: Math.max(0, employees.length - 1)
  };
}

function submitEvaluation(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employees = getEmployees_(ss);
    const sections = getSections_(ss);

    const evaluator = String(payload.evaluator || '').trim();
    const target = String(payload.target || '').trim();
    const monthKey = normalizeMonthKey_(payload.monthKey);
    const positive = String(payload.positive || '').trim();
    const improvement = String(payload.improvement || '').trim();
    const answers = payload.answers || {};

    if (!employees.includes(evaluator)) throw new Error('Invalid evaluator.');
    if (!employees.includes(target)) throw new Error('Invalid team member.');
    if (evaluator === target) throw new Error('Self-evaluation is not allowed.');
    if (positive.length < 3 || improvement.length < 3) {
      throw new Error('Please complete both written feedback fields.');
    }

    const flatQuestions = sections.flatMap(section =>
      section.questions.map(q => ({ ...q, sectionId: section.id }))
    );

    flatQuestions.forEach(q => {
      const value = Number(answers[q.key]);
      if (!Number.isInteger(value) || value < TW.scoreMin || value > TW.scoreMax) {
        throw new Error('Please answer every question using a score from 1 to 5.');
      }
    });

    const rows = getResponseRows_(ss);

    if (rows.some(r =>
      r.monthKey === monthKey &&
      r.evaluator === evaluator &&
      r.target === target
    )) {
      throw new Error('This review has already been submitted for this month.');
    }

    const sectionScores = {};

    sections.forEach(section => {
      const values = section.questions.map(q => Number(answers[q.key]));
      sectionScores[section.id] = values.reduce((a, b) => a + b, 0) / values.length;
    });

    const totalWeight =
      sections.reduce((sum, section) => sum + Number(section.weight || 0), 0) || 100;

    const weightedScore =
      sections.reduce(
        (sum, section) =>
          sum + sectionScores[section.id] * Number(section.weight || 0),
        0
      ) / totalWeight;

    const sheet = ss.getSheetByName(TW.sheets.responses);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const rowMap = {
      'Timestamp': new Date(),
      'Month': monthKey,
      'Evaluator': evaluator,
      'Target': target,
      'Weighted Score': round_(weightedScore, 3),
      'Strength': positive,
      'Improvement Suggestion': improvement,
      'Response ID': Utilities.getUuid()
    };

    flatQuestions.forEach(q => {
      rowMap[q.key] = Number(answers[q.key]);
    });

    sections.forEach(section => {
      rowMap['SCORE_' + section.id] = round_(sectionScores[section.id], 3);
    });

    sheet.appendRow(
      headers.map(header =>
        rowMap[header] !== undefined ? rowMap[header] : ''
      )
    );

    refreshDashboard();

    return {
      success: true,
      message: 'Review submitted successfully.',
      status: getEvaluatorStatus(evaluator, monthKey)
    };
  } finally {
    lock.releaseLock();
  }
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName(TW.sheets.dashboard) || ss.insertSheet(TW.sheets.dashboard);

  const employees = getEmployees_(ss);
  const sections = getSections_(ss);

  const rawMonth = sheet.getRange('B3').getValue();
  let monthKey;

  if (rawMonth instanceof Date && !isNaN(rawMonth)) {
    monthKey = Utilities.formatDate(rawMonth, TW.timezone, 'yyyy-MM');
  } else {
    monthKey = String(rawMonth || '').trim();
  }

  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(monthKey)) {
    monthKey = Utilities.formatDate(new Date(), TW.timezone, 'yyyy-MM');
  }

  const rows = getResponseRows_(ss).filter(r => r.monthKey === monthKey);

  sheet.clear();
  sheet.setRightToLeft(false);

  sheet.getRange('A1:H1')
    .merge()
    .setValue('Monthly Teamwork Review Dashboard');

  sheet.getRange('A2:H2')
    .merge()
    .setValue('Team scores, review completion, and category-level performance');

  sheet.getRange('A3').setValue('Report Month');
  sheet.getRange('B3').setValue(monthKey);

  const scores = rows
    .map(r => Number(r.weightedScore))
    .filter(Number.isFinite);

  const teamAvg = scores.length
    ? scores.reduce((a, b) => a + b, 0) / scores.length
    : 0;

  const required = employees.length * Math.max(0, employees.length - 1);
  const completion = required ? rows.length / required : 0;

  sheet.getRange('A5:D6').setValues([
    ['Team Average', 'Submitted Reviews', 'Expected Reviews', 'Completion'],
    [round_(teamAvg, 2), rows.length, required, completion]
  ]);

  const head = ['Team Member', 'Reviews Received', 'Overall Score'];
  sections.forEach(section => head.push(section.title));

  const table = [head];

  employees.forEach(name => {
    const personRows = rows.filter(r => r.target === name);
    const avg = personRows.length
      ? avg_(personRows.map(r => Number(r.weightedScore)).filter(Number.isFinite))
      : 0;

    const values = [name, personRows.length, round_(avg, 2)];

    sections.forEach(section => {
      values.push(
        round_(
          avg_(
            personRows
              .map(r => Number(r.sectionScores[section.id]))
              .filter(Number.isFinite)
          ),
          2
        )
      );
    });

    table.push(values);
  });

  sheet.getRange(9, 1, table.length, table[0].length).setValues(table);

  const completionHead = [
    'Evaluator',
    'Completed',
    'Remaining',
    'Completion'
  ];

  const completionRows = [completionHead];

  employees.forEach(evaluator => {
    const done = new Set(
      rows
        .filter(r => r.evaluator === evaluator)
        .map(r => r.target)
    ).size;

    const req = Math.max(0, employees.length - 1);

    completionRows.push([
      evaluator,
      done,
      Math.max(0, req - done),
      req ? done / req : 0
    ]);
  });

  const start = 12 + employees.length;

  sheet
    .getRange(start, 1, completionRows.length, completionRows[0].length)
    .setValues(completionRows);

  // Styling
  sheet
    .getRange('A1:H1')
    .setBackground('#7d1838')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');

  sheet
    .getRange('A2:H2')
    .setBackground('#f8edf1')
    .setFontColor('#303139')
    .setHorizontalAlignment('center');

  sheet
    .getRange('A5:D5')
    .setBackground('#191b20')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet
    .getRange(9, 1, 1, table[0].length)
    .setBackground('#7d1838')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet
    .getRange(start, 1, 1, 4)
    .setBackground('#191b20')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.getRange('D6').setNumberFormat('0%');
  sheet.getRange(start + 1, 4, employees.length, 1).setNumberFormat('0%');

  sheet.setFrozenRows(3);
  sheet.autoResizeColumns(1, Math.max(8, table[0].length));
  sheet.setColumnWidth(1, 180);

  for (let c = 2; c <= table[0].length; c++) {
    sheet.setColumnWidth(c, 125);
  }
}

function ensureSettingsSheet_(ss) {
  const sheet =
    ss.getSheetByName(TW.sheets.settings) || ss.insertSheet(TW.sheets.settings);

  sheet.setRightToLeft(false);

  const employees = [
    'Alex Morgan',
    'Jordan Lee',
    'Taylor Brooks',
    'Casey Parker'
  ];

  if (sheet.getLastRow() < 2) {
    sheet.getRange('A1:B6').setValues([
      ['Setting', 'Value'],
      ['System Title', 'Monthly Teamwork Review'],
      ['Score Scale', '1 to 5'],
      ['Submission Rule', 'Each evaluator reviews each teammate once per month'],
      ['', ''],
      ['Team Members', 'Active']
    ]);

    sheet
      .getRange(7, 1, employees.length, 2)
      .setValues(employees.map(name => [name, true]));
  }

  sheet
    .getRange('A1:B1')
    .setBackground('#7d1838')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.autoResizeColumns(1, 2);
}

function ensureQuestionsSheet_(ss) {
  const sheet =
    ss.getSheetByName(TW.sheets.questions) || ss.insertSheet(TW.sheets.questions);

  sheet.setRightToLeft(false);

  if (sheet.getLastRow() > 1) return;

  const data = defaultQuestionRows_();

  sheet
    .getRange(1, 1, 1, 7)
    .setValues([[
      'Section ID',
      'Category',
      'Weight',
      'Question Order',
      'Question Key',
      'Question',
      'Active'
    ]]);

  sheet.getRange(2, 1, data.length, 7).setValues(data);

  sheet
    .getRange('A1:G1')
    .setBackground('#7d1838')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.getRange('F:F').setWrap(true);
  sheet.setColumnWidth(2, 190);
  sheet.setColumnWidth(6, 620);
  sheet.setFrozenRows(1);
}

function ensureResponsesSheet_(ss) {
  const sheet =
    ss.getSheetByName(TW.sheets.responses) || ss.insertSheet(TW.sheets.responses);

  sheet.setRightToLeft(false);

  if (sheet.getLastColumn() > 1 && sheet.getRange(1, 1).getValue()) return;

  const sections = getSections_(ss);
  const qKeys = sections.flatMap(section =>
    section.questions.map(question => question.key)
  );
  const scoreKeys = sections.map(section => 'SCORE_' + section.id);

  const headers = [
    'Timestamp',
    'Month',
    'Evaluator',
    'Target',
    ...qKeys,
    ...scoreKeys,
    'Weighted Score',
    'Strength',
    'Improvement Suggestion',
    'Response ID'
  ];

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet
    .getRange(1, 1, 1, headers.length)
    .setBackground('#191b20')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 145);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);

  if (qKeys.length) {
    sheet.setColumnWidths(5, qKeys.length, 75);
  }

  sheet.setColumnWidth(headers.indexOf('Strength') + 1, 300);
  sheet.setColumnWidth(headers.indexOf('Improvement Suggestion') + 1, 300);
}

function ensureDashboardSheet_(ss) {
  if (!ss.getSheetByName(TW.sheets.dashboard)) {
    ss.insertSheet(TW.sheets.dashboard);
  }
}

function getEmployees_(ss) {
  const sheet = ss.getSheetByName(TW.sheets.settings);
  if (!sheet) return [];

  const last = sheet.getLastRow();
  if (last < 7) return [];

  return sheet
    .getRange(7, 1, last - 6, 2)
    .getValues()
    .filter(row =>
      String(row[0]).trim() &&
      (
        row[1] === true ||
        String(row[1]).toLowerCase() === 'true' ||
        row[1] === 'Active'
      )
    )
    .map(row => String(row[0]).trim());
}

function getSections_(ss) {
  const sheet = ss.getSheetByName(TW.sheets.questions);

  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 7)
    .getValues()
    .filter(row => row[6] !== false && String(row[0]).trim());

  const map = new Map();

  rows.forEach(row => {
    const id = String(row[0]).trim();

    if (!map.has(id)) {
      map.set(id, {
        id,
        title: String(row[1]),
        weight: Number(row[2]),
        questions: []
      });
    }

    map.get(id).questions.push({
      order: Number(row[3]),
      key: String(row[4]),
      text: String(row[5])
    });
  });

  return [...map.values()].map(section => ({
    ...section,
    questions: section.questions.sort((a, b) => a.order - b.order)
  }));
}

function getResponseRows_(ss) {
  const sheet = ss.getSheetByName(TW.sheets.responses);

  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const idx = Object.fromEntries(headers.map((header, i) => [header, i]));
  const sections = getSections_(ss);

  return values
    .filter(row => row.some(value => value !== ''))
    .map(row => ({
      monthKey: String(row[idx['Month']] || ''),
      evaluator: String(row[idx['Evaluator']] || ''),
      target: String(row[idx['Target']] || ''),
      weightedScore: Number(row[idx['Weighted Score']]),
      sectionScores: Object.fromEntries(
        sections.map(section => [
          section.id,
          Number(row[idx['SCORE_' + section.id]])
        ])
      )
    }));
}

function normalizeMonthKey_(value) {
  const fallback = Utilities.formatDate(
    new Date(),
    TW.timezone,
    'yyyy-MM'
  );

  const s = String(value || fallback).trim();

  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(s)) {
    throw new Error('Month must use the format YYYY-MM.');
  }

  return s;
}

function avg_(arr) {
  return arr.length
    ? arr.reduce((a, b) => a + b, 0) / arr.length
    : 0;
}

function round_(n, d) {
  const p = Math.pow(10, d || 0);
  return Math.round((Number(n) || 0) * p) / p;
}

function defaultQuestionRows_() {
  const sections = [
    ['casework', 'Casework Collaboration', 20, [
      'When a teammate struggles with a case, they actively help solve it.',
      'They collaborate with teammates to find solutions for complex cases.',
      'They share useful experience or solutions that can help a teammate.'
    ]],
    ['responsiveness', 'Team Responsiveness', 15, [
      'They respond appropriately when a teammate asks for help or advice.',
      'They do not unnecessarily ignore teammates’ work-related requests.',
      'If they cannot respond immediately, they follow up at an appropriate time.'
    ]],
    ['handoff', 'Effective Handoffs', 15, [
      'They route work to the correct teammate when ownership belongs elsewhere.',
      'They provide the necessary context when handing work over.',
      'They avoid repeated back-and-forth caused by incomplete handoffs.',
      'Before handing work over, they check whether they can resolve it themselves.'
    ]],
    ['voluntary', 'Voluntary Support', 10, [
      'When a teammate has a heavy workload, they offer help without being asked.',
      'They help move team work forward when shared tasks fall behind.',
      'During busy periods, they contribute beyond only their individual workload.'
    ]],
    ['knowledge', 'Knowledge Sharing', 10, [
      'They share useful new experiences and solutions with the team.',
      'When they know how to solve a problem, they make the solution available to others.',
      'They do not keep important process knowledge to themselves.',
      'When needed, they help teammates learn a process or topic.'
    ]],
    ['coverage', 'Work Coverage', 10, [
      'They help cover essential work when a teammate is unavailable.',
      'When a teammate’s workload increases, they help reduce the pressure where possible.',
      'Teammates can rely on them for temporary coverage when needed.'
    ]],
    ['problemSolving', 'Shared Problem Solving', 10, [
      'When complex issues arise, they focus on solving the problem rather than assigning blame.',
      'They consider teammates’ perspectives when making decisions on shared cases.',
      'They participate in finding practical solutions to team-level problems.',
      'When a team process goes wrong, they cooperate to correct it.'
    ]],
    ['managerReferral', 'Appropriate Manager Escalation', 10, [
      'Before escalating an issue to a manager, they consult teammates when appropriate.',
      'They do not escalate issues unnecessarily when the team can resolve them.',
      'When escalating, they provide the relevant context and checks already completed.',
      'They take ownership instead of waiting for a manager to decide every issue.'
    ]]
  ];

  const rows = [];

  sections.forEach(section => {
    section[3].forEach((question, index) => {
      rows.push([
        section[0],
        section[1],
        section[2],
        index + 1,
        `${section[0]}_${index}`,
        question,
        true
      ]);
    });
  });

  return rows;
}
