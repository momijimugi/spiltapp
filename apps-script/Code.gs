const PROJECT_HEADERS = [
  'id', 'title', 'status', 'client', 'deadline', 'bpm', 'duration',
  'splitA', 'splitB', 'finalA', 'analysisJson', 'tasksJson', 'createdAt', 'updatedAt',
  'archived', 'archivedAt'
];
const LOG_HEADERS = [
  'id', 'projectId', 'person', 'type', 'name', 'count', 'duration',
  'events', 'details', 'createdAt', 'scope', 'effort'
];
const DELETED_PROJECT_HEADERS = PROJECT_HEADERS.concat(['deletedAt']);
const DELETED_LOG_HEADERS = LOG_HEADERS.concat(['projectTitle', 'deletedAt']);

/**
 * Run once from the Apps Script editor after setting SHEET_ID and API_KEY
 * in Project Settings > Script Properties.
 */
function setupDatabase() {
  const spreadsheet = getSpreadsheet_();
  ensureSheet_(spreadsheet, 'Projects', PROJECT_HEADERS);
  ensureSheet_(spreadsheet, 'ProductionLogs', LOG_HEADERS);
  ensureSheet_(spreadsheet, 'DeletedProjects', DELETED_PROJECT_HEADERS);
  ensureSheet_(spreadsheet, 'DeletedProductionLogs', DELETED_LOG_HEADERS);
  return 'SPLITLAB database is ready.';
}

function doGet() {
  return json_({
    ok: true,
    service: 'SPLITLAB Sheets API',
    version: 1
  });
}

function doPost(event) {
  try {
    const parameters = event && event.parameter ? event.parameter : {};
    authorize_(parameters.apiKey || '');
    const payload = JSON.parse(parameters.payload || 'null');
    if (parameters.action === 'sync') {
      if (!Array.isArray(payload)) throw new Error('Payload must be an array.');
      return json_({ ok: true, projects: syncProjects_(payload) });
    }
    if (parameters.action === 'parseLogs') {
      return json_({ ok: true, logs: parseNarrative_(payload) });
    }
    if (parameters.action === 'analyze') {
      return json_({ ok: true, analysis: analyzeProject_(payload) });
    }
    if (parameters.action === 'deleteProject') {
      if (!payload || !payload.projectId) throw new Error('Project ID is required.');
      return json_({ ok: true, ...deleteProject_(payload.projectId) });
    }
    throw new Error('Unsupported action.');
  } catch (error) {
    return json_({ ok: false, error: String(error.message || error) });
  }
}

function authorize_(providedKey) {
  const expectedKey = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!expectedKey) throw new Error('API_KEY is not configured.');
  if (!providedKey || providedKey !== expectedKey) throw new Error('Unauthorized.');
}

function getSpreadsheet_() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) throw new Error('SHEET_ID is not configured.');
  return SpreadsheetApp.openById(sheetId);
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#11141c')
    .setFontColor('#c7ff3d')
    .setFontWeight('bold');
  return sheet;
}

function syncProjects_(incomingProjects) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const spreadsheet = getSpreadsheet_();
    const projectSheet = ensureSheet_(spreadsheet, 'Projects', PROJECT_HEADERS);
    const logSheet = ensureSheet_(spreadsheet, 'ProductionLogs', LOG_HEADERS);
    const deletedProjectSheet = ensureSheet_(spreadsheet, 'DeletedProjects', DELETED_PROJECT_HEADERS);
    const existingRows = readRows_(projectSheet, PROJECT_HEADERS);
    const existingById = Object.fromEntries(existingRows.map(record => [record.id, record]));
    const deletedIds = new Set(readRows_(deletedProjectSheet, DELETED_PROJECT_HEADERS).map(record => String(record.id)));

    incomingProjects.forEach(project => {
      if (!project || !project.id) return;
      if (deletedIds.has(String(project.id))) return;
      const existing = existingById[project.id];
      const incomingTime = Date.parse(project.updatedAt || '') || 0;
      const existingTime = Date.parse(existing && existing.updatedAt || '') || 0;
      if (!existing || incomingTime >= existingTime) {
        upsertProject_(projectSheet, project);
        replaceProjectLogs_(logSheet, project.id, Array.isArray(project.logs) ? project.logs : []);
      }
    });
    SpreadsheetApp.flush();
    return readProjects_(projectSheet, logSheet);
  } finally {
    lock.releaseLock();
  }
}

function deleteProject_(projectId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const spreadsheet = getSpreadsheet_();
    const projectSheet = ensureSheet_(spreadsheet, 'Projects', PROJECT_HEADERS);
    const logSheet = ensureSheet_(spreadsheet, 'ProductionLogs', LOG_HEADERS);
    const deletedProjectSheet = ensureSheet_(spreadsheet, 'DeletedProjects', DELETED_PROJECT_HEADERS);
    const deletedLogSheet = ensureSheet_(spreadsheet, 'DeletedProductionLogs', DELETED_LOG_HEADERS);
    const projectRowNumber = findRowById_(projectSheet, projectId);
    if (!projectRowNumber) throw new Error('案件が見つかりません。先に同期してください。');

    const projectRow = projectSheet.getRange(projectRowNumber, 1, 1, PROJECT_HEADERS.length).getValues()[0];
    const archivedValue = projectRow[PROJECT_HEADERS.indexOf('archived')];
    const isArchived = archivedValue === true || String(archivedValue).toLowerCase() === 'true';
    if (!isArchived) throw new Error('削除できるのはアーカイブ済み案件だけです。');

    const deletedAt = new Date().toISOString();
    const projectTitle = String(projectRow[PROJECT_HEADERS.indexOf('title')] || 'Untitled Track');
    deletedProjectSheet.appendRow(projectRow.concat([deletedAt]));

    const logRowsToDelete = [];
    const deletedLogRows = [];
    if (logSheet.getLastRow() > 1) {
      const rows = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, LOG_HEADERS.length).getValues();
      rows.forEach((row, index) => {
        if (String(row[1]) !== String(projectId)) return;
        deletedLogRows.push(row.concat([projectTitle, deletedAt]));
        logRowsToDelete.push(index + 2);
      });
    }
    if (deletedLogRows.length) {
      deletedLogSheet.getRange(deletedLogSheet.getLastRow() + 1, 1, deletedLogRows.length, DELETED_LOG_HEADERS.length).setValues(deletedLogRows);
    }

    logRowsToDelete.reverse().forEach(rowNumber => logSheet.deleteRow(rowNumber));
    projectSheet.deleteRow(projectRowNumber);
    SpreadsheetApp.flush();
    return {
      deletedProjectId: String(projectId),
      archivedLogCount: deletedLogRows.length,
      projects: readProjects_(projectSheet, logSheet)
    };
  } finally {
    lock.releaseLock();
  }
}

function upsertProject_(sheet, project) {
  const row = [
    project.id,
    project.title || 'Untitled Track',
    project.status || '制作中',
    project.client || '',
    project.deadline || '',
    project.bpm || '',
    Number(project.duration) || 180,
    Number(project.splitA) || 0,
    Number(project.splitB) || 0,
    Number(project.finalA ?? project.splitA) || 0,
    JSON.stringify(project.analysis || null),
    JSON.stringify(project.tasks || {}),
    project.createdAt || new Date().toISOString(),
    project.updatedAt || new Date().toISOString(),
    project.archived === true,
    project.archivedAt || ''
  ];
  const rowNumber = findRowById_(sheet, project.id);
  if (rowNumber) sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
}

function replaceProjectLogs_(sheet, projectId, logs) {
  if (sheet.getLastRow() > 1) {
    const projectIds = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    for (let index = projectIds.length - 1; index >= 0; index -= 1) {
      if (String(projectIds[index][0]) === String(projectId)) sheet.deleteRow(index + 2);
    }
  }
  if (!logs.length) return;
  const rows = logs.map(log => [
    log.id || Utilities.getUuid(),
    projectId,
    log.person === 'B' ? 'riku' : 'tada',
    log.type || 'instrument',
    log.name || '',
    Math.max(1, Number(log.count) || 1),
    Math.max(0, Number(log.duration) || 0),
    Math.max(0, Number(log.events) || 0),
    log.details || '',
    log.createdAt || new Date().toISOString(),
    level_(log.scope, 0),
    level_(log.effort, 0)
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, LOG_HEADERS.length).setValues(rows);
}

function readProjects_(projectSheet, logSheet) {
  const logs = readRows_(logSheet, LOG_HEADERS);
  const logsByProject = {};
  logs.forEach(log => {
    if (!logsByProject[log.projectId]) logsByProject[log.projectId] = [];
    logsByProject[log.projectId].push({
      id: log.id,
      person: String(log.person).toLowerCase() === 'riku' || log.person === 'B' ? 'B' : 'A',
      type: log.type,
      name: log.name,
      count: Number(log.count) || 1,
      duration: Math.max(0, Number(log.duration) || 0),
      events: Math.max(0, Number(log.events) || 0),
      scope: level_(log.scope, 0),
      effort: level_(log.effort, 0),
      details: log.details,
      createdAt: log.createdAt
    });
  });
  return readRows_(projectSheet, PROJECT_HEADERS).map(record => ({
    id: record.id,
    title: record.title,
    status: record.status,
    client: record.client,
    deadline: record.deadline,
    bpm: record.bpm,
    duration: Number(record.duration) || 180,
    splitA: Number(record.splitA) || 0,
    splitB: Number(record.splitB) || 0,
    finalA: Number(record.finalA ?? record.splitA) || 0,
    analysis: parseJson_(record.analysisJson, null),
    tasks: parseJson_(record.tasksJson, {}),
    logs: logsByProject[record.id] || [],
    archived: record.archived === true || String(record.archived).toLowerCase() === 'true',
    archivedAt: record.archivedAt || '',
    createdAt: record.createdAt,
    updatedAt: record.updatedAt
  }));
}

function parseNarrative_(payload) {
  const text = payload && String(payload.text || '').trim();
  if (!text) throw new Error('分解する文章がありません。');
  if (text.length > 12000) throw new Error('文章は12,000文字以内にしてください。');
  const properties = PropertiesService.getScriptProperties();
  const geminiKey = properties.getProperty('GEMINI_API_KEY');
  const model = properties.getProperty('GEMINI_MODEL') || 'gemini-3.1-flash-lite';
  if (!geminiKey) throw new Error('GEMINI_API_KEY is not configured.');
  const allowedTypes = ['melody', 'structure', 'motif', 'harmony', 'beat', 'bass', 'guitar', 'instrument', 'sound', 'sample', 'mix', 'delivery'];
  const prompt = [
    '音楽制作コライトの作業報告を、登録可能な制作ログへ分解してください。',
    '複数の担当者、作業、カテゴリが含まれる場合は必ず別ログに分けてください。文章に書かれていない数値や機材、採用箇所は推測しないでください。',
    '担当者はtadaならA、rikuならB。不明なら空文字にしてuncertainFieldsへ「担当」を追加してください。',
    'typeは次から選択: melody, structure, motif, harmony, beat, bass, guitar, instrument, sound, sample, mix, delivery。',
    'countは本数。不明なら1。scopeは貢献範囲、effortは制作負荷（カロリー）を次の5段階から選んでください。',
    'scope: 1=ワンポイント、2=1セクション、3=複数セクション（1〜2構成）、4=曲全体、5=全体を通した継続作業。',
    'effort: 1=軽作業、2=やや軽い、3=標準、4=高負荷、5=非常に高負荷。難易度・専門性・試行錯誤・所要労力を総合してください。',
    '原文から判断できないscopeまたはeffortは中立値3にし、uncertainFieldsへ「貢献範囲」または「制作負荷」を追加してください。',
    'nameは短い作業名。detailsには曲中での役割、使用機材、加工、採用箇所など、原文にある事実を残してください。',
    '不明または曖昧な項目名をuncertainFields配列に入れてください。最大20件です。',
    'JSONのみを返してください。形式: {"logs":[{"person":"A|B|","type":"...","name":"...","count":1,"scope":3,"effort":3,"details":"...","uncertainFields":["担当","貢献範囲"]}]}',
    `案件情報: ${JSON.stringify(payload.project || {})}`,
    `作業報告: ${text}`
  ].join('\n');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': geminiKey },
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.05, responseMimeType: 'application/json' }
    }),
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  const responseBody = response.getContentText();
  if (status < 200 || status >= 300) throw new Error(`Gemini API error ${status}: ${responseBody.slice(0, 300)}`);
  const geminiResponse = JSON.parse(responseBody);
  const responseText = geminiResponse.candidates && geminiResponse.candidates[0] && geminiResponse.candidates[0].content.parts[0].text;
  if (!responseText) throw new Error('Gemini API returned no log candidates.');
  const parsed = JSON.parse(responseText);
  const logs = Array.isArray(parsed.logs) ? parsed.logs.slice(0, 20) : [];
  return logs.map(log => {
    const uncertain = Array.isArray(log.uncertainFields) ? log.uncertainFields.map(String) : [];
    const person = log.person === 'B' ? 'B' : log.person === 'A' ? 'A' : '';
    const type = allowedTypes.indexOf(log.type) >= 0 ? log.type : 'instrument';
    if (!person && uncertain.indexOf('担当') < 0) uncertain.push('担当');
    if (allowedTypes.indexOf(log.type) < 0 && uncertain.indexOf('カテゴリ') < 0) uncertain.push('カテゴリ');
    return {
      person: person,
      type: type,
      name: String(log.name || '制作作業').slice(0, 200),
      count: Math.max(1, Math.round(Number(log.count) || 1)),
      scope: level_(log.scope, 3),
      effort: level_(log.effort, 3),
      details: String(log.details || '').slice(0, 1000),
      uncertainFields: uncertain.slice(0, 8)
    };
  });
}
function analyzeProject_(payload) {
  if (!payload || !Array.isArray(payload.logs) || !payload.logs.length) {
    throw new Error('分析する制作ログがありません。');
  }
  const properties = PropertiesService.getScriptProperties();
  const geminiKey = properties.getProperty('GEMINI_API_KEY');
  const model = properties.getProperty('GEMINI_MODEL') || 'gemini-3.1-flash-lite';
  if (!geminiKey) throw new Error('GEMINI_API_KEY is not configured.');
  const baseline = payload.baseline || {};
  const prompt = [
    'あなたは音楽制作コライトの貢献分析者です。品質や人物の優劣ではなく、完成曲の成立に対する音楽的中心性だけを評価してください。',
    'メインメロディー、曲構成、固有モチーフは高い比重。コード、ビート、ベース、ミックスは曲への影響範囲で評価。scopeは影響範囲、effortは制作負荷として扱い、単純なトラック追加やサンプル配置は物量が多くても音楽的比重を過大評価しないでください。',
    'tada（内部値A）の音楽的比重を0〜100のmusicalAで返してください。証拠はログに書かれた事実だけを使い、推測しないでください。',
    'JSONのみを返してください。形式: {"musicalA":number,"summary":string,"evidence":[string],"musicalDetail":string}',
    `案件: ${JSON.stringify(payload.project || {})}`,
    `制作ログ: ${JSON.stringify(payload.logs)}`
  ].join('\n');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': geminiKey },
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: {
        temperature: 0.15,
        responseMimeType: 'application/json'
      }
    }),
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  const responseBody = response.getContentText();
  if (status < 200 || status >= 300) throw new Error(`Gemini API error ${status}: ${responseBody.slice(0, 300)}`);
  const geminiResponse = JSON.parse(responseBody);
  const text = geminiResponse.candidates && geminiResponse.candidates[0] && geminiResponse.candidates[0].content.parts[0].text;
  if (!text) throw new Error('Gemini API returned no analysis.');
  const parsed = JSON.parse(text);
  const quantityA = clamp_(Number(baseline.quantityA), 0, 100, 50);
  const musicalA = clamp_(Number(parsed.musicalA), 0, 100, 50);
  return {
    metricVersion: 2,
    quantityA: quantityA,
    musicalA: musicalA,
    recommendedA: quantityA * 0.4 + musicalA * 0.6,
    quantityDetail: baseline.quantityDetail || '本数・貢献範囲・制作負荷から機械算出',
    musicalDetail: parsed.musicalDetail || '制作ログの役割と採用範囲から評価',
    summary: parsed.summary || '',
    evidence: Array.isArray(parsed.evidence) ? parsed.evidence.slice(0, 6) : [],
    model: model
  };
}

function level_(value, fallback) {
  const level = Math.round(Number(value));
  return Number.isFinite(level) && level >= 1 && level <= 5 ? level : fallback;
}

function clamp_(value, minimum, maximum, fallback) {
  return Number.isFinite(value) ? Math.max(minimum, Math.min(maximum, value)) : fallback;
}

function readRows_(sheet, headers) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues()
    .filter(row => row[0] !== '')
    .map(row => Object.fromEntries(headers.map((header, index) => [header, serializeCell_(row[index])])));
}

function findRowById_(sheet, id) {
  if (sheet.getLastRow() < 2) return 0;
  const match = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(id))
    .matchEntireCell(true)
    .findNext();
  return match ? match.getRow() : 0;
}

function serializeCell_(value) {
  return value instanceof Date ? value.toISOString() : value;
}

function parseJson_(value, fallback) {
  try { return JSON.parse(value || ''); }
  catch (error) { return fallback; }
}

function json_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}