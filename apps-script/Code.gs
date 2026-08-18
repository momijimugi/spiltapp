const PROJECT_HEADERS = [
  'id', 'title', 'status', 'client', 'deadline', 'bpm', 'duration',
  'splitA', 'splitB', 'finalA', 'analysisJson', 'tasksJson', 'createdAt', 'updatedAt',
  'archived', 'archivedAt'
];
const LOG_HEADERS = [
  'id', 'projectId', 'person', 'type', 'name', 'count', 'duration',
  'events', 'details', 'createdAt', 'scope', 'effort',
  // 新しい列は必ず末尾へ追加する（既存シートの列位置を壊さないため）。
  'contributionMode'
];
// 任意入力。未選択（空文字）を必ず許容する。
const CONTRIBUTION_MODES = {
  creation: '新規作成',
  development: '発展・再構築',
  modification: '修正・調整',
  integration: '統合・判断',
  execution: '実装・演奏'
};
// 5軸のデフォルトウェイト。調整はここだけで済むよう一箇所に集約する。
const ANALYSIS_WEIGHTS = {
  quantity: 0.25,
  musical: 0.30,
  agency: 0.20,
  resolution: 0.20,
  fiveAxis: 0.05
};
// これ未満のconfidenceのAI評価は最終スプリットへ入れない。
const MIN_AXIS_CONFIDENCE = 0.70;
const DELETED_PROJECT_HEADERS = PROJECT_HEADERS.concat(['deletedAt']);
const DELETED_LOG_HEADERS = LOG_HEADERS.concat(['projectTitle', 'deletedAt']);
const BETA_ANALYSIS_HEADERS = ['projectId', 'projectTitle', 'status', 'analysisJson', 'updatedAt', 'confirmedAt'];
const DELETED_BETA_ANALYSIS_HEADERS = BETA_ANALYSIS_HEADERS.concat(['deletedAt']);

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
  ensureSheet_(spreadsheet, 'BetaAnalyses', BETA_ANALYSIS_HEADERS);
  ensureSheet_(spreadsheet, 'DeletedBetaAnalyses', DELETED_BETA_ANALYSIS_HEADERS);
  return 'SPLITLAB database is ready.';
}

function doGet() {
  return json_({
    ok: true,
    service: 'SPLITLAB Sheets API',
    version: 3
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
    if (parameters.action === 'syncDelta') {
      if (!payload || typeof payload !== 'object') throw new Error('Delta payload is required.');
      return json_({ ok: true, ...syncDelta_(payload) });
    }
    if (parameters.action === 'parseLogs') {
      return json_({ ok: true, logs: parseNarrative_(payload) });
    }
    if (parameters.action === 'analyze') {
      return json_({ ok: true, analysis: analyzeProject_(payload) });
    }
    if (parameters.action === 'analyzeCombined') {
      return json_({ ok: true, analysis: analyzeCombined_(payload) });
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
  return syncProjectChanges_(incomingProjects, '', true).projects;
}

function syncDelta_(payload) {
  const incomingProjects = Array.isArray(payload.changes) ? payload.changes : [];
  return syncProjectChanges_(incomingProjects, payload.since || '', payload.full === true);
}

function syncProjectChanges_(incomingProjects, since, full) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) throw new Error('別端末の同期処理中です。少し待って自動再試行します。');
  try {
    const spreadsheet = getSpreadsheet_();
    const projectSheet = ensureSheet_(spreadsheet, 'Projects', PROJECT_HEADERS);
    const logSheet = ensureSheet_(spreadsheet, 'ProductionLogs', LOG_HEADERS);
    const deletedProjectSheet = ensureSheet_(spreadsheet, 'DeletedProjects', DELETED_PROJECT_HEADERS);
    const existingRows = readRows_(projectSheet, PROJECT_HEADERS);
    const existingById = Object.fromEntries(existingRows.map(record => [record.id, record]));
    const existingLogsByProject = {};
    readRows_(logSheet, LOG_HEADERS).forEach(log => {
      if (!existingLogsByProject[log.projectId]) existingLogsByProject[log.projectId] = [];
      existingLogsByProject[log.projectId].push(log);
    });
    const deletedRows = readRows_(deletedProjectSheet, DELETED_PROJECT_HEADERS);
    const deletedIds = new Set(deletedRows.map(record => String(record.id)));
    const incomingIds = new Set();

    incomingProjects.forEach(project => {
      if (!project || !project.id) return;
      const projectId = String(project.id);
      incomingIds.add(projectId);
      if (deletedIds.has(projectId)) return;
      const existing = existingById[project.id];
      const incomingTime = Date.parse(project.updatedAt || '') || 0;
      const existingTime = Date.parse(existing && existing.updatedAt || '') || 0;
      if (!existing || incomingTime > existingTime) {
        const incomingLogs = Array.isArray(project.logs) ? project.logs : [];
        upsertProject_(projectSheet, project);
        if (!existing || !logsEquivalent_(existingLogsByProject[project.id] || [], incomingLogs)) {
          replaceProjectLogs_(logSheet, project.id, incomingLogs);
        }
      }
    });

    SpreadsheetApp.flush();
    const sinceTime = Date.parse(since || '') || 0;
    const allProjects = readProjects_(projectSheet, logSheet);
    const projects = full
      ? allProjects
      : allProjects.filter(project => incomingIds.has(String(project.id)) || (Date.parse(project.updatedAt || '') || 0) > sinceTime);
    const deletedProjectIds = deletedRows
      .filter(record => full || incomingIds.has(String(record.id)) || (Date.parse(record.deletedAt || '') || 0) > sinceTime)
      .map(record => String(record.id));
    return {
      projects,
      deletedProjectIds,
      full,
      serverTime: new Date().toISOString()
    };
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
    const betaSheet = ensureSheet_(spreadsheet, 'BetaAnalyses', BETA_ANALYSIS_HEADERS);
    const deletedBetaSheet = ensureSheet_(spreadsheet, 'DeletedBetaAnalyses', DELETED_BETA_ANALYSIS_HEADERS);
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

    const betaRowNumber = findRowById_(betaSheet, projectId);
    let archivedBetaAnalysis = false;
    if (betaRowNumber) {
      const betaRow = betaSheet.getRange(betaRowNumber, 1, 1, BETA_ANALYSIS_HEADERS.length).getValues()[0];
      deletedBetaSheet.appendRow(betaRow.concat([deletedAt]));
      deleteDataRows_(betaSheet, [betaRowNumber]);
      archivedBetaAnalysis = true;
    }

    deleteDataRows_(logSheet, logRowsToDelete);
    deleteDataRows_(projectSheet, [projectRowNumber]);
    SpreadsheetApp.flush();
    return {
      deletedProjectId: String(projectId),
      archivedLogCount: deletedLogRows.length,
      archivedBetaAnalysis: archivedBetaAnalysis,
      projects: readProjects_(projectSheet, logSheet),
      serverTime: new Date().toISOString()
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
  const rowsToDelete = [];
  if (sheet.getLastRow() > 1) {
    const projectIds = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    for (let index = projectIds.length - 1; index >= 0; index -= 1) {
      if (String(projectIds[index][0]) === String(projectId)) rowsToDelete.push(index + 2);
    }
  }
  deleteDataRows_(sheet, rowsToDelete);
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
    level_(log.effort, 0),
    contributionMode_(log.contributionMode)
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, LOG_HEADERS.length).setValues(rows);
}

function logsEquivalent_(storedLogs, incomingLogs) {
  const normalizeStored = log => [
    String(log.id || ''),
    String(log.person).toLowerCase() === 'riku' || log.person === 'B' ? 'B' : 'A',
    String(log.type || 'instrument'),
    String(log.name || ''),
    Math.max(1, Number(log.count) || 1),
    Math.max(0, Number(log.duration) || 0),
    Math.max(0, Number(log.events) || 0),
    String(log.details || ''),
    String(log.createdAt || ''),
    level_(log.scope, 0),
    level_(log.effort, 0),
    contributionMode_(log.contributionMode)
  ];
  const normalizeIncoming = log => normalizeStored({
    ...log,
    person: log.person === 'B' ? 'B' : 'A'
  });
  return JSON.stringify(storedLogs.map(normalizeStored)) === JSON.stringify(incomingLogs.map(normalizeIncoming));
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
      contributionMode: contributionMode_(log.contributionMode),
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
    '担当者は必ず名前で扱い、tadaまたはrikuを返してください。不明なら空文字にしてuncertainFieldsへ「担当」を追加してください。記号や代替名で表現しないでください。',
    'typeは次から選択: melody, structure, motif, harmony, beat, bass, guitar, instrument, sound, sample, mix, delivery。',
    'countは本数。不明なら1。scopeは貢献範囲、effortは制作負荷（カロリー）を次の5段階から選んでください。',
    'scope: 1=ワンポイント（単発の差し込み・装飾）、2=1セクション（Aメロ・サビなど一部の構成）、3=複数セクション（2〜3程度の構成にまたがる範囲）、4=曲の大部分（サビ以外にも複数箇所で使われる、曲の骨格に近い要素）、5=曲全体（一曲を通して存在する要素）。',
    'effort: 1=軽作業、2=やや工夫が必要、3=標準的な負荷、4=高負荷、5=非常に高負荷。難易度・専門性・試行錯誤・所要労力を総合してください。',
    '原文から判断できないscopeまたはeffortは中立値3にし、uncertainFieldsへ「貢献範囲」または「制作負荷」を追加してください。',
    'nameは短い作業名。detailsには役割、元になった案、変更内容、採用箇所など、原文にある事実を残してください。',
    'contributionModeは任意項目です。原文からその人の関わり方が明確に読み取れる場合だけ creation(新規作成) / development(発展・再構築) / modification(修正・調整) / integration(統合・判断) / execution(実装・演奏) から選んでください。',
    '例:「rikuがサビメロを0から作った」→"creation"。「tadaがrikuのコード案を元にサビを全面的に作り直した」→"development"。判断できない場合は必ず空文字""にし、推測で埋めないでください。contributionModeが空でも問題ありません。uncertainFieldsへ追加する必要もありません。',
    '不明または曖昧な項目名をuncertainFields配列に入れてください。最大20件です。',
    'JSONのみを返してください。形式: {"logs":[{"person":"tada|riku|","type":"...","name":"...","count":1,"scope":3,"effort":3,"contributionMode":"","details":"...","uncertainFields":["担当","貢献範囲"]}]}',
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
  const parsed = JSON.parse(extractJsonObject_(responseText));
  const logs = Array.isArray(parsed.logs) ? parsed.logs.slice(0, 20) : [];
  return logs.map(log => {
    const uncertain = Array.isArray(log.uncertainFields) ? log.uncertainFields.map(String) : [];
    const personName = String(log.person || '').trim().toLowerCase();
    const person = personName === 'riku' ? 'B' : personName === 'tada' ? 'A' : '';
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
      contributionMode: contributionMode_(log.contributionMode),
      details: String(log.details || '').slice(0, 1000),
      uncertainFields: uncertain.slice(0, 8)
    };
  });
}
function analyzeBeta_(payload) {
  if (!payload || !Array.isArray(payload.logs) || !payload.logs.length) {
    throw new Error('分析する制作ログがありません。');
  }
  if (payload.logs.length > 200) throw new Error('制作ログは200件以内で分析してください。');
  const properties = PropertiesService.getScriptProperties();
  const geminiKey = properties.getProperty('GEMINI_API_KEY');
  const model = properties.getProperty('GEMINI_MODEL') || 'gemini-3.1-flash-lite';
  if (!geminiKey) throw new Error('GEMINI_API_KEY is not configured.');
  const namedLogs = payload.logs.map(log => ({
    id: String(log.id || ''),
    person: String(log.person || '').toLowerCase() === 'riku' || log.person === 'B' ? 'riku' : 'tada',
    category: String(log.type || 'instrument'),
    name: String(log.name || ''),
    scope: level_(log.scope, 3),
    details: String(log.details || '')
  }));
  const project = payload.project || {};
  const projectSummary = {
    title: String(project.title || 'Untitled Track'),
    status: String(project.status || ''),
    client: String(project.client || ''),
    bpm: Number(project.bpm) || '',
    duration: Number(project.duration) || 0
  };
  const prompt = [
    'あなたは音楽制作コライトの構造的な貢献を整理する分析者です。人物の優劣、作業時間、制作負荷、演奏技術の上手さ、好みは評価しません。完成曲へ残った音楽的な影響だけを整理してください。',
    '担当者名はtadaとrikuです。回答文とpersonには必ずこの名前を使い、記号・頭文字・代替名を使用しないでください。',
    '関連する複数ログが同じメロディー、モチーフ、コード、ビート、音色、ミックス処理を指す場合は、必ず1つのmusical elementへ統合してください。ログ件数を貢献点にしないでください。',
    'カテゴリは melody, structure, motif, harmony, beat, bass, guitar, instrument, sound, sample, mix, delivery のいずれかです。完成曲に実際に存在するカテゴリだけに重要度を与え、categoryWeightsのweight合計を100にしてください。',
    '各elementのidentityScoreは曲の同一性・核への近さ、scopeScoreは完成曲での影響範囲を1〜5で評価してください。',
    'contributorsのroleScoreは 1=微調整、2=整理・統合、3=発展、4=大幅な変形・再構築、5=核となる原案。roleには短い日本語の役割名を入れてください。',
    'adoptionScoreは最終版への残存度を0〜5で評価し、完全に不採用なら0にしてください。irreplaceabilityScoreはその貢献を抜いた場合に曲の印象・成立が変わる程度を1〜5で評価してください。',
    '各点数には制作ログに明記された事実だけを根拠としてevidenceへ日本語で記載し、推測しないでください。confidenceは0〜1です。根拠不足は中立値3（adoptionのみ3）としconfidenceを0.69以下にしてください。',
    'AIは最終割合を決めません。割合・人物評価・勝敗に関する文章を出さないでください。',
    'summaryには、抽出したmusical elementの名前とカテゴリ、各担当者がどの要素にどう関わったか(原案・発展・整理など)、なぜそのcategoryWeightsになったのかを、具体例を挙げながら4〜6文程度で説明してください。「メロディーの比重が高い」のような抽象的な一文だけで終えないでください。',
    'JSONのみを返してください。形式: {"summary":"分析の詳しい説明","categoryWeights":[{"category":"melody","weight":25,"reason":"理由","confidence":0.8}],"elements":[{"id":"element-1","name":"サビの主旋律","category":"melody","identityScore":5,"scopeScore":3,"evidence":"根拠","confidence":0.9,"contributors":[{"person":"tada","role":"核となる原案","roleScore":5,"adoptionScore":5,"irreplaceabilityScore":5,"evidence":"根拠","confidence":0.9}]}]}',
    `案件: ${JSON.stringify(projectSummary)}`,
    `制作ログ: ${JSON.stringify(namedLogs)}`
  ].join('\n');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': geminiKey },
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.08, responseMimeType: 'application/json' }
    }),
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  const responseBody = response.getContentText();
  if (status < 200 || status >= 300) throw new Error(`Gemini API error ${status}: ${responseBody.slice(0, 300)}`);
  const geminiResponse = JSON.parse(responseBody);
  const responseText = geminiResponse.candidates && geminiResponse.candidates[0] && geminiResponse.candidates[0].content.parts[0].text;
  if (!responseText) throw new Error('Gemini API returned no beta analysis.');
  const normalized = normalizeBetaAnalysis_(JSON.parse(extractJsonObject_(responseText)), model);
  normalized.calculation = calculateBetaScores_(normalized);
  return normalized;
}

function saveBetaAnalysis_(payload) {
  const projectId = payload && String(payload.projectId || '').trim();
  if (!projectId || !payload.analysis) throw new Error('保存するベータ分析がありません。');
  const status = payload.status === 'confirmed' ? 'confirmed' : 'draft';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) throw new Error('別端末の保存処理中です。少し待って再試行してください。');
  try {
    const spreadsheet = getSpreadsheet_();
    const betaSheet = ensureSheet_(spreadsheet, 'BetaAnalyses', BETA_ANALYSIS_HEADERS);
    const normalized = normalizeBetaAnalysis_(payload.analysis, String(payload.analysis.model || 'manual-review'));
    normalized.status = status;
    normalized.updatedAt = new Date().toISOString();
    normalized.confirmedAt = status === 'confirmed' ? normalized.updatedAt : '';
    normalized.calculation = calculateBetaScores_(normalized);
    const projectTitle = String(payload.projectTitle || 'Untitled Track').slice(0, 300);
    const row = [projectId, projectTitle, status, JSON.stringify(normalized), normalized.updatedAt, normalized.confirmedAt];
    const rowNumber = findRowById_(betaSheet, projectId);
    if (rowNumber) betaSheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    else betaSheet.appendRow(row);
    SpreadsheetApp.flush();
    return normalized;
  } finally {
    lock.releaseLock();
  }
}

function normalizeBetaAnalysis_(input, model) {
  const source = input && typeof input === 'object' ? input : {};
  const allowedCategories = ['melody', 'structure', 'motif', 'harmony', 'beat', 'bass', 'guitar', 'instrument', 'sound', 'sample', 'mix', 'delivery'];
  const elements = (Array.isArray(source.elements) ? source.elements : []).slice(0, 40).map((element, elementIndex) => {
    const category = allowedCategories.indexOf(String(element.category || '')) >= 0 ? String(element.category) : 'instrument';
    const contributorsByName = {};
    (Array.isArray(element.contributors) ? element.contributors : []).forEach(contributor => {
      const person = String(contributor.person || '').trim().toLowerCase();
      if (person !== 'tada' && person !== 'riku') return;
      const normalizedContributor = {
        person: person,
        role: String(contributor.role || '役割不明').slice(0, 120),
        roleScore: level_(contributor.roleScore, 3),
        adoptionScore: clamp_(Math.round(Number(contributor.adoptionScore)), 0, 5, 3),
        irreplaceabilityScore: level_(contributor.irreplaceabilityScore, 3),
        evidence: String(contributor.evidence || '').slice(0, 1000),
        confidence: confidence_(contributor.confidence)
      };
      const existing = contributorsByName[person];
      if (!existing) contributorsByName[person] = normalizedContributor;
      else {
        existing.roleScore = Math.max(existing.roleScore, normalizedContributor.roleScore);
        existing.adoptionScore = Math.max(existing.adoptionScore, normalizedContributor.adoptionScore);
        existing.irreplaceabilityScore = Math.max(existing.irreplaceabilityScore, normalizedContributor.irreplaceabilityScore);
        existing.confidence = Math.max(existing.confidence, normalizedContributor.confidence);
        if (!existing.evidence && normalizedContributor.evidence) existing.evidence = normalizedContributor.evidence;
      }
    });
    return {
      id: String(element.id || `element-${elementIndex + 1}`).slice(0, 120),
      name: String(element.name || `音楽要素 ${elementIndex + 1}`).slice(0, 200),
      category: category,
      identityScore: level_(element.identityScore, 3),
      scopeScore: level_(element.scopeScore, 3),
      evidence: String(element.evidence || '').slice(0, 1000),
      confidence: confidence_(element.confidence),
      contributors: Object.keys(contributorsByName).map(person => contributorsByName[person])
    };
  }).filter(element => element.contributors.length > 0);
  if (!elements.length) throw new Error('音楽要素を抽出できませんでした。制作ログの詳細を追加してください。');

  const categoriesInElements = [...new Set(elements.map(element => element.category))];
  const suppliedWeights = Array.isArray(source.categoryWeights) ? source.categoryWeights : [];
  const weightsByCategory = {};
  suppliedWeights.forEach(item => {
    const category = String(item.category || '');
    if (categoriesInElements.indexOf(category) < 0) return;
    weightsByCategory[category] = {
      category: category,
      weight: Math.max(0, Number(item.weight) || 0),
      reason: String(item.reason || '').slice(0, 500),
      confidence: confidence_(item.confidence)
    };
  });
  categoriesInElements.forEach(category => {
    if (!weightsByCategory[category]) weightsByCategory[category] = { category: category, weight: 1, reason: '制作ログに含まれる要素', confidence: 0.5 };
  });
  const categoryWeights = normalizeCategoryWeights_(Object.keys(weightsByCategory).map(category => weightsByCategory[category]));
  const normalized = {
    metricVersion: 3,
    status: source.status === 'confirmed' ? 'confirmed' : 'draft',
    summary: String(source.summary || '制作ログを5軸へ分解したベータ分析です。').slice(0, 1200),
    categoryWeights: categoryWeights,
    elements: elements,
    model: String(model || source.model || ''),
    updatedAt: String(source.updatedAt || new Date().toISOString()),
    confirmedAt: String(source.confirmedAt || '')
  };
  return normalized;
}

function normalizeCategoryWeights_(weights) {
  const clean = weights.map(item => ({ ...item, weight: Math.max(0, Number(item.weight) || 0) }));
  let total = clean.reduce((sum, item) => sum + item.weight, 0);
  if (!total) {
    clean.forEach(item => { item.weight = 1; });
    total = clean.length;
  }
  let allocated = 0;
  clean.forEach((item, index) => {
    item.weight = index === clean.length - 1 ? Math.max(0, 100 - allocated) : Math.round(item.weight / total * 1000) / 10;
    allocated += item.weight;
  });
  return clean;
}

function calculateBetaScores_(analysis) {
  const categoryWeights = Object.fromEntries((analysis.categoryWeights || []).map(item => [item.category, Number(item.weight) || 0]));
  // confidenceが低い評価は中立値3のままフルスコア加算されないよう、計算対象から除外する。
  const excludedForConfidence = [];
  const eligibleElements = (analysis.elements || []).filter(element => {
    if (confidence_(element.confidence) >= MIN_AXIS_CONFIDENCE) return true;
    excludedForConfidence.push(`要素: ${element.name}`);
    return false;
  }).map(element => ({
    ...element,
    contributors: (element.contributors || []).filter(contributor => {
      if (confidence_(contributor.confidence) < MIN_AXIS_CONFIDENCE) {
        excludedForConfidence.push(`${element.name}: ${contributor.person}`);
        return false;
      }
      return Number(contributor.adoptionScore) > 0;
    })
  })).filter(element => element.contributors.length > 0);
  const categories = [...new Set(eligibleElements.map(element => element.category))];
  let tadaPoints = 0;
  let rikuPoints = 0;
  const categoryBreakdown = [];
  categories.forEach(category => {
    const categoryElements = eligibleElements.filter(element => element.category === category);
    const factors = categoryElements.map(element => 0.6 * level_(element.identityScore, 3) / 5 + 0.4 * level_(element.scopeScore, 3) / 5);
    const factorTotal = factors.reduce((sum, factor) => sum + factor, 0) || 1;
    let categoryTada = 0;
    let categoryRiku = 0;
    categoryElements.forEach((element, index) => {
      const elementPoints = (categoryWeights[category] || 0) * factors[index] / factorTotal;
      const contributorScores = element.contributors.map(contributor => ({
        person: contributor.person,
        score: 0.4 * level_(contributor.roleScore, 3) / 5 + 0.35 * clamp_(Number(contributor.adoptionScore), 0, 5, 3) / 5 + 0.25 * level_(contributor.irreplaceabilityScore, 3) / 5
      }));
      const contributorTotal = contributorScores.reduce((sum, item) => sum + item.score, 0) || 1;
      contributorScores.forEach(item => {
        const points = elementPoints * item.score / contributorTotal;
        if (item.person === 'riku') categoryRiku += points;
        else categoryTada += points;
      });
    });
    tadaPoints += categoryTada;
    rikuPoints += categoryRiku;
    categoryBreakdown.push({ category: category, weight: categoryWeights[category] || 0, tadaPoints: categoryTada, rikuPoints: categoryRiku });
  });
  const total = tadaPoints + rikuPoints;
  // 十分な比較材料が残らない場合は「判定保留」とし、50:50を作らない。
  const contributorsLeft = [...new Set(eligibleElements.reduce((names, element) => names.concat(element.contributors.map(c => c.person)), []))];
  const available = eligibleElements.length > 0 && total > 0 && contributorsLeft.length > 0;
  const tadaPercent = available ? tadaPoints / total * 100 : null;
  const reviewItems = [];
  (analysis.categoryWeights || []).forEach(item => {
    if (Number(item.confidence) < 0.7 || !String(item.reason || '').trim()) reviewItems.push(`カテゴリ: ${item.category}`);
  });
  (analysis.elements || []).forEach(element => {
    if (Number(element.confidence) < 0.7 || !String(element.evidence || '').trim()) reviewItems.push(`要素: ${element.name}`);
    (element.contributors || []).forEach(contributor => {
      if (Number(contributor.confidence) < 0.7 || !String(contributor.evidence || '').trim()) reviewItems.push(`${element.name}: ${contributor.person}`);
    });
  });
  return {
    available: available,
    tadaPercent: tadaPercent,
    rikuPercent: available ? 100 - tadaPercent : null,
    usedElementCount: eligibleElements.length,
    excludedForConfidence: [...new Set(excludedForConfidence)].slice(0, 30),
    categoryBreakdown: categoryBreakdown,
    needsReviewCount: reviewItems.length,
    reviewItems: reviewItems.slice(0, 30)
  };
}

function confidence_(value) {
  let confidence = Number(value);
  if (!Number.isFinite(confidence)) return 0.5;
  if (confidence > 1) confidence /= 100;
  return Math.max(0, Math.min(1, confidence));
}
function analyzeCombined_(payload) {
  if (!payload || !Array.isArray(payload.logs) || !payload.logs.length) {
    throw new Error('分析する制作ログがありません。');
  }
  const project = payload.project || {};
  // Musical / Agency / Resolution は1回のGeminiリクエストでまとめて取得する。
  const logAnalysis = analyzeProject_(payload);
  let betaAnalysis = null;
  try {
    betaAnalysis = analyzeBeta_({ project: project, logs: payload.logs });
  } catch (error) {
    // 5軸は補助軸なので、失敗しても判定保留として全体分析は返す。
    betaAnalysis = null;
  }
  const betaCalculation = betaAnalysis && betaAnalysis.calculation;
  const betaAvailable = !!(betaCalculation && betaCalculation.available);
  const betaTadaPercent = betaAvailable ? clamp_(Number(betaCalculation.tadaPercent), 0, 100, null) : null;

  const musicalAvailable = Number.isFinite(Number(logAnalysis.musicalA)) && confidence_(logAnalysis.musicalConfidence) >= MIN_AXIS_CONFIDENCE;
  const agencyAvailable = Number.isFinite(Number(logAnalysis.creativeAgencyA)) && confidence_(logAnalysis.agencyConfidence) >= MIN_AXIS_CONFIDENCE;
  const resolutionAvailable = Number.isFinite(Number(logAnalysis.creativeResolutionA)) && confidence_(logAnalysis.resolutionConfidence) >= MIN_AXIS_CONFIDENCE;

  // 利用できない軸はウェイトから外し、残りの軸を100%へ再正規化する。物量は常に利用可能。
  const combined = combineAxes_([
    { key: 'quantity', value: Number(logAnalysis.quantityA), weight: ANALYSIS_WEIGHTS.quantity, available: true },
    { key: 'musical', value: Number(logAnalysis.musicalA), weight: ANALYSIS_WEIGHTS.musical, available: musicalAvailable },
    { key: 'agency', value: Number(logAnalysis.creativeAgencyA), weight: ANALYSIS_WEIGHTS.agency, available: agencyAvailable },
    { key: 'resolution', value: Number(logAnalysis.creativeResolutionA), weight: ANALYSIS_WEIGHTS.resolution, available: resolutionAvailable },
    { key: 'fiveAxis', value: betaTadaPercent, weight: ANALYSIS_WEIGHTS.fiveAxis, available: betaAvailable }
  ]);

  if (project.id && betaAnalysis) {
    try {
      saveBetaAnalysis_({ projectId: project.id, projectTitle: project.title || 'Untitled Track', status: 'draft', analysis: betaAnalysis });
    } catch (error) {
      // ベータ保存に失敗しても統合分析自体は返す。
    }
  }

  return {
    metricVersion: 5,
    weights: ANALYSIS_WEIGHTS,
    quantityA: logAnalysis.quantityA,
    musicalA: musicalAvailable ? logAnalysis.musicalA : null,
    creativeAgencyA: agencyAvailable ? logAnalysis.creativeAgencyA : null,
    creativeResolutionA: resolutionAvailable ? logAnalysis.creativeResolutionA : null,
    betaTadaPercent: betaTadaPercent,
    musicalConfidence: confidence_(logAnalysis.musicalConfidence),
    agencyConfidence: confidence_(logAnalysis.agencyConfidence),
    resolutionConfidence: confidence_(logAnalysis.resolutionConfidence),
    recommendedA: combined.value,
    effectiveWeightTotal: combined.weightTotal,
    usedAxes: combined.usedAxes,
    quantityDetail: logAnalysis.quantityDetail,
    musicalDetail: logAnalysis.musicalDetail,
    agencyDetail: logAnalysis.agencyDetail,
    resolutionDetail: logAnalysis.resolutionDetail,
    betaSummary: betaAnalysis ? betaAnalysis.summary : '',
    beta: betaAnalysis,
    summary: logAnalysis.summary,
    evidence: logAnalysis.evidence,
    model: logAnalysis.model
  };
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
  const namedLogs = payload.logs.map(log => ({
    person: log.person === 'B' ? 'riku' : 'tada',
    type: log.type,
    name: log.name,
    count: log.count,
    scope: log.scope,
    effort: log.effort,
    contributionMode: contributionMode_(log.contributionMode),
    contributionModeLabel: CONTRIBUTION_MODES[contributionMode_(log.contributionMode)] || '',
    details: log.details
  }));
  const prompt = [
    'あなたは音楽制作コライトの貢献分析者です。品質や人物の優劣ではなく、完成曲の成立に対する音楽的中心性だけを評価してください。',
    'メインメロディー、曲構成、固有モチーフは高い比重。コード、ビート、ベース、ミックスは曲への影響範囲で評価。scopeは影響範囲、effortは制作負荷として扱い、単純なトラック追加やサンプル配置は物量が多くても音楽的比重を過大評価しないでください。',
    '担当者はtadaとrikuです。回答文では必ずこの名前を使い、記号や代替名で表現しないでください。',
    'tadaの音楽的比重を0〜100のtadaMusicalPercentで返してください。証拠はログに書かれた事実だけを使い、推測しないでください。',
    '物量の割合はシステムが本数・貢献範囲・制作負荷から算出済みです。割合を変更せず、quantityCommentには両者の物量差とその主な理由を、ログに出てくる具体的な作業名・本数・貢献範囲(1〜5)・制作負荷(1〜5)の値を複数挙げながら3〜5文程度で具体的に説明してください。抽象的な言い回しは避け、どのログがどれだけ物量に効いたかが分かるようにしてください。',
    'musicalDetailには、tadaとrikuそれぞれのどの作業が完成曲の音楽的な骨格(メロディー・構成・モチーフなど)にどの程度影響したかを、作業名とカテゴリに触れながら3〜5文程度で具体的に説明してください。',
    '',
    '同じ制作ログから、さらに2つの軸を評価してください。',
    '【Creative Agency / 創作主体性】その音楽要素に対してどの程度主体的な創作を行ったかを見ます。1=実装・再現・微調整、2=小規模な修正・追加、3=独自の発展・新規要素追加、4=大幅な再構築・創造的展開、5=核となる原案・0→1の創作。',
    'Agencyの強い根拠: 0から作った、考案した、原案を作った、新しいメロディー/コードを書いた、全面再構築した、元案から大幅に展開した。弱い根拠: 打ち込んだ、録音した、コピーした、微調整した、音量を直した、決まった内容を再現した。',
    'Agencyでは「作業時間」「本数」「制作負荷」「演奏技術の上手さ」を根拠にしてはいけません。それらは物量側で評価済みです。tadaのAgency割合を0〜100のtadaAgencyPercentで返してください。',
    '【Creative Resolution / 完成・収束寄与】その人の作業や判断で曲がどれだけ完成形へ前進したかを見ます。複数案の統合、構成の整理、不要部分の削除、サビの成立、コードとメロディーの矛盾解決、他者素材の組み合わせ、曲全体の方向決定などを拾います。1=自分の担当部分を処理しただけ、5=楽曲全体に関わる重要な判断・収束。',
    '単に大量に作ったことをResolutionの加点理由にしてはいけません。スケジュール管理やプロジェクト管理も対象外です。tadaのResolution割合を0〜100のtadaResolutionPercentで返してください。',
    'ログのcontributionMode(関わり方)は任意入力です。設定されていれば有力な手がかりとして使い、空欄でもdetailsから明確に読み取れる場合だけ判断してください。読み取れない場合は推測せず、該当軸のconfidenceを0.69以下にしてください。過去ログの修正をユーザーへ要求する文言は書かないでください。',
    'musicalConfidence / agencyConfidence / resolutionConfidence を0〜1で返してください。ログに明確な根拠がある場合だけ0.70以上にし、根拠が乏しい場合は必ず0.69以下にしてください。低confidenceの軸は最終計算から除外されます。',
    'agencyDetail / resolutionDetail には、それぞれの判断根拠をログの具体的な作業名を挙げながら2〜4文で書いてください。',
    'summaryには、物量・音楽的比重・創作主体性・完成寄与を踏まえた総合的な所見を3〜4文程度でまとめてください。割合の正解を決めるのではなく、話し合いの材料を示す姿勢で書いてください。',
    'JSONのみを返してください。形式: {"tadaMusicalPercent":number,"tadaAgencyPercent":number,"tadaResolutionPercent":number,"musicalConfidence":number,"agencyConfidence":number,"resolutionConfidence":number,"quantityComment":string,"musicalDetail":string,"agencyDetail":string,"resolutionDetail":string,"summary":string,"evidence":[string]}',
    `案件: ${JSON.stringify(payload.project || {})}`,
    `物量集計: ${JSON.stringify({ quantityTadaPercent: baseline.quantityA, detail: baseline.quantityDetail, evidence: baseline.evidence })}`,
    `制作ログ: ${JSON.stringify(namedLogs)}`
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
  const parsed = JSON.parse(extractJsonObject_(text));
  const quantityA = clamp_(Number(baseline.quantityA), 0, 100, 50);
  // 値が返らなかった軸はnullのまま（判定保留）とし、中立値50を作らない。
  const musicalA = clamp_(Number(parsed.tadaMusicalPercent), 0, 100, null);
  const agencyA = clamp_(Number(parsed.tadaAgencyPercent), 0, 100, null);
  const resolutionA = clamp_(Number(parsed.tadaResolutionPercent), 0, 100, null);
  return {
    metricVersion: 5,
    quantityA: quantityA,
    musicalA: musicalA,
    creativeAgencyA: agencyA,
    creativeResolutionA: resolutionA,
    musicalConfidence: confidence_(parsed.musicalConfidence),
    agencyConfidence: confidence_(parsed.agencyConfidence),
    resolutionConfidence: confidence_(parsed.resolutionConfidence),
    quantityDetail: parsed.quantityComment || baseline.quantityDetail || '本数・貢献範囲・制作負荷から機械算出',
    musicalDetail: parsed.musicalDetail || '制作ログの役割と採用範囲から評価',
    agencyDetail: parsed.agencyDetail || '',
    resolutionDetail: parsed.resolutionDetail || '',
    summary: parsed.summary || '',
    evidence: Array.isArray(parsed.evidence) ? parsed.evidence.slice(0, 6) : [],
    model: model
  };
}

function extractJsonObject_(text) {
  const start = text.indexOf('{');
  if (start < 0) throw new Error('AIの応答からJSONを取り出せませんでした。');
  let depth = 0;
  let inString = false;
  let escapeNext = false;
  for (let i = start; i < text.length; i += 1) {
    const char = text.charAt(i);
    if (inString) {
      if (escapeNext) escapeNext = false;
      else if (char === '\\') escapeNext = true;
      else if (char === '"') inString = false;
      continue;
    }
    if (char === '"') { inString = true; continue; }
    if (char === '{') depth += 1;
    else if (char === '}') {
      depth -= 1;
      if (depth === 0) return text.slice(start, i + 1);
    }
  }
  throw new Error('AIの応答のJSONが途中で切れています。');
}

function level_(value, fallback) {
  const level = Math.round(Number(value));
  return Number.isFinite(level) && level >= 1 && level <= 5 ? level : fallback;
}

// 未設定・未知の値は空文字（未選択）として扱う。古いログには列自体が無い。
function contributionMode_(value) {
  const mode = String(value == null ? '' : value).trim();
  return Object.prototype.hasOwnProperty.call(CONTRIBUTION_MODES, mode) ? mode : '';
}

/**
 * 利用可能な軸だけでウェイトを再正規化する汎用関数。
 * axes: [{ value:number, weight:number, available:boolean }]
 */
function combineAxes_(axes) {
  const usable = axes.filter(axis => axis.available && Number.isFinite(Number(axis.value)));
  const weightTotal = usable.reduce((sum, axis) => sum + Number(axis.weight || 0), 0);
  if (!usable.length || weightTotal <= 0) return { value: 50, weightTotal: 0, usedAxes: [] };
  const value = usable.reduce((sum, axis) => sum + Number(axis.value) * Number(axis.weight), 0) / weightTotal;
  return {
    value: clamp_(value, 0, 100, 50),
    weightTotal: weightTotal,
    usedAxes: usable.map(axis => axis.key)
  };
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

function deleteDataRows_(sheet, rowNumbers) {
  [...new Set(rowNumbers)].sort((a, b) => b - a).forEach(rowNumber => {
    if (sheet.getMaxRows() <= sheet.getFrozenRows() + 1) {
      sheet.getRange(rowNumber, 1, 1, sheet.getMaxColumns()).clearContent();
    } else {
      sheet.deleteRow(rowNumber);
    }
  });
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