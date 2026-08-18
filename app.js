    const categories = [
      { name: 'アイデア・構成', code: '01', tasks: [
        { id: 'reference', name: 'リファレンス選定', points: 3 },
        { id: 'structure', name: '全体構成', points: 7 }
      ]},
      { name: 'メロディ・コード', code: '02', tasks: [
        { id: 'chords', name: 'メインコード進行', points: 10 },
        { id: 'topline', name: 'トップライン', points: 10 },
        { id: 'counter', name: '対旋律', points: 5 }
      ]},
      { name: 'ビート・ベース', code: '03', tasks: [
        { id: 'beat', name: 'メインビート', points: 7 },
        { id: 'bass', name: 'ベースライン', points: 8 },
        { id: 'percussion', name: 'パーカッション', points: 5 }
      ]},
      { name: 'サウンド・編集', code: '04', tasks: [
        { id: 'main-sound', name: 'メイン音作り', points: 10 },
        { id: 'layers', name: '上モノ音作り', points: 5 },
        { id: 'fx', name: 'FX配置', points: 5 },
        { id: 'editing', name: '波形・ベロシティ編集', points: 5 }
      ]},
      { name: '録音・追加トラック', code: '05', tasks: [
        { id: 'guitar-recording', name: 'ギター録音', points: 6, adjustable: true },
        { id: 'instrument-tracks', name: '楽器トラック追加', points: 5, adjustable: true },
        { id: 'sample-tracks', name: 'サンプル・ループ追加', points: 4, adjustable: true }
      ]},
      { name: 'ミックス・納品', code: '06', tasks: [
        { id: 'level-pan', name: '音量・パン調整', points: 7 },
        { id: 'eq-comp', name: 'EQ・コンプ', points: 6 },
        { id: 'space', name: '空間系', points: 5 },
        { id: 'stems', name: 'ステム書き出し', points: 2 }
      ]}
    ];

    const taskList = document.getElementById('task-list');
    const allTasks = categories.flatMap(category => category.tasks.map(task => ({ ...task, category: category.name })));
    const maxPoints = allTasks.reduce((sum, task) => sum + task.points, 0);
    const PERSON_NAMES = { A: 'tada', B: 'riku' };
    const SCOPE_LEVELS = {
      1: { label: 'ワンポイント', description: '単発の差し込み・装飾' },
      2: { label: '1セクション', description: 'Aメロ・サビなど一部の構成' },
      3: { label: '複数セクション', description: '2〜3程度の構成にまたがる範囲' },
      4: { label: '曲の大部分', description: 'サビ以外にも複数箇所で使われる、曲の骨格に近い要素' },
      5: { label: '曲全体', description: '一曲を通して存在する要素' }
    };
    const EFFORT_LEVELS = {
      1: { label: '軽作業', description: '短時間でできる定型作業' },
      2: { label: 'やや工夫が必要', description: '手順の工夫が必要な作業' },
      3: { label: '標準的な負荷', description: '通常の制作負荷' },
      4: { label: '高負荷', description: '専門性や試行錯誤が多い作業' },
      5: { label: '非常に高負荷', description: '高度で継続的な作業' }
    };
    const STORAGE_KEY = 'splitlab_projects_v2';
    const API_URL_KEY = 'splitlab_apps_script_url';
    const DIRTY_PROJECTS_KEY = 'splitlab_dirty_projects_v1';
    const LAST_REMOTE_SYNC_KEY = 'splitlab_last_remote_sync_v1';
    const SYNC_SCHEMA_KEY = 'splitlab_sync_schema';
    const SYNC_SCHEMA_VERSION = 'delta-v1';
    // Bump to force every client to do one full fetch (see the backfill note below).
    const ANALYSIS_BACKFILL_KEY = 'splitlab_analysis_backfill';
    const ANALYSIS_BACKFILL_VERSION = 'v1';
    // Analyses produced by a superseded metric formula are discarded on load so
    // stale numbers never resurface. Keep this at or below the lowest version
    // still in use: analyzeCombined_ returns 4, the local calculation returns 2.
    const MIN_ANALYSIS_METRIC_VERSION = 2;
    let projects = loadProjects();
    let dirtyProjectIds = loadDirtyProjectIds();
    let lastRemoteSyncAt = localStorage.getItem(LAST_REMOTE_SYNC_KEY) || '';
    if (localStorage.getItem(SYNC_SCHEMA_KEY) !== SYNC_SCHEMA_VERSION) {
      projects.forEach(project => dirtyProjectIds.add(project.id));
      lastRemoteSyncAt = '';
      persistSyncState();
      localStorage.setItem(SYNC_SCHEMA_KEY, SYNC_SCHEMA_VERSION);
    }
    // One-time repair: analyses written to the sheet before the load-time guard was
    // fixed are stranded there, because syncDelta only returns projects whose
    // updatedAt is newer than the last sync — so they are never sent down again and
    // reloading cannot help. Clearing lastRemoteSyncAt forces one full fetch that
    // pulls them back. Deliberately does NOT mark projects dirty (unlike the schema
    // migration above): dirty local copies win over remote on a full sync, which
    // would push the missing analyses back over the good rows in the sheet.
    if (localStorage.getItem(ANALYSIS_BACKFILL_KEY) !== ANALYSIS_BACKFILL_VERSION) {
      lastRemoteSyncAt = '';
      persistSyncState();
      localStorage.setItem(ANALYSIS_BACKFILL_KEY, ANALYSIS_BACKFILL_VERSION);
    }
    let activeProjectId = null;
    let saveTimer;
    let autoSyncTimer;
    let syncInFlight = false;
    let syncQueued = false;

    function loadProjects() {
      try { return JSON.parse(localStorage.getItem(STORAGE_KEY)) || []; }
      catch { return []; }
    }

    function persistProjects() {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(projects));
    }

    function loadDirtyProjectIds() {
      try { return new Set(JSON.parse(localStorage.getItem(DIRTY_PROJECTS_KEY)) || []); }
      catch { return new Set(); }
    }

    function persistSyncState() {
      localStorage.setItem(DIRTY_PROJECTS_KEY, JSON.stringify([...dirtyProjectIds]));
      if (lastRemoteSyncAt) localStorage.setItem(LAST_REMOTE_SYNC_KEY, lastRemoteSyncAt);
      else localStorage.removeItem(LAST_REMOTE_SYNC_KEY);
    }

    function markProjectDirty(projectId) {
      if (!projectId) return;
      dirtyProjectIds.add(projectId);
      persistSyncState();
    }

    let fullSyncProgressTimer = null;
    let fullSyncHideTimer = null;

    function setFullSyncProgress(percent, phase) {
      const value = Math.max(0, Math.min(100, Math.round(percent)));
      document.getElementById('full-sync-percent').textContent = `${value}%`;
      document.getElementById('full-sync-remaining').textContent = `残り${100 - value}%`;
      document.getElementById('full-sync-bar').style.width = `${value}%`;
      if (phase) document.getElementById('full-sync-phase').textContent = phase;
    }

    function startFullSyncProgress() {
      clearInterval(fullSyncProgressTimer);
      clearTimeout(fullSyncHideTimer);
      const container = document.getElementById('full-sync-progress');
      container.classList.remove('hidden');
      setFullSyncProgress(8, 'ローカルの変更を整理しています');
      let progress = 8;
      fullSyncProgressTimer = setInterval(() => {
        progress = Math.min(82, progress + (progress < 45 ? 7 : 3));
        setFullSyncProgress(progress, progress < 45 ? 'Google Sheetsへ接続しています' : '案件と制作ログを取得しています');
      }, 450);
    }

    function finishFullSyncProgress(success) {
      clearInterval(fullSyncProgressTimer);
      fullSyncProgressTimer = null;
      const container = document.getElementById('full-sync-progress');
      if (success) {
        setFullSyncProgress(100, '全件取得が完了しました');
        fullSyncHideTimer = setTimeout(() => container.classList.add('hidden'), 900);
      } else {
        document.getElementById('full-sync-phase').textContent = '全件取得を完了できませんでした。自動再試行します';
        fullSyncHideTimer = setTimeout(() => container.classList.add('hidden'), 2400);
      }
    }

    function defaultTaskState() {
      return Object.fromEntries(allTasks.map(task => [task.id, {
        enabled: true,
        a: 50,
        amount: task.adjustable ? 100 : null
      }]));
    }

    function makeProject() {
      const now = new Date().toISOString();
      return {
        id: `project-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        title: 'Untitled Track',
        status: '制作中',
        client: '',
        deadline: '',
        bpm: '',
        duration: 180,
        tasks: defaultTaskState(),
        logs: [],
        splitA: 50,
        splitB: 50,
        finalA: 50,
        analysis: null,
        archived: false,
        archivedAt: '',
        createdAt: now,
        updatedAt: now
      };
    }

    function getActiveProject() {
      return projects.find(project => project.id === activeProjectId) || null;
    }

    function normalizeLevel(value, fallback = 3) {
      const level = Math.round(Number(value));
      return Number.isFinite(level) && level >= 1 && level <= 5 ? level : fallback;
    }

    function legacyScopeLevel(duration) {
      const seconds = Math.max(0, Number(duration) || 0);
      if (!seconds) return 3;
      if (seconds <= 10) return 1;
      if (seconds <= 45) return 2;
      if (seconds <= 90) return 3;
      if (seconds <= 240) return 4;
      return 5;
    }

    function legacyEffortLevel(events) {
      const count = Math.max(0, Number(events) || 0);
      if (!count) return 3;
      if (count <= 10) return 1;
      if (count <= 50) return 2;
      if (count <= 150) return 3;
      if (count <= 400) return 4;
      return 5;
    }

    function normalizeLog(log) {
      const fallbackDate = (log.id && typeof log.id === 'string' && log.id.startsWith('log-'))
        ? new Date(Number(log.id.split('-')[1]) || Date.now()).toISOString()
        : new Date().toISOString();
      return {
        ...log,
        createdAt: log.createdAt || fallbackDate,
        scope: normalizeLevel(log.scope, legacyScopeLevel(log.duration)),
        effort: normalizeLevel(log.effort, legacyEffortLevel(log.events))
      };
    }

    function normalizeProject(project) {
      const defaults = defaultTaskState();
      const archived = project.archived === true || project.status === 'アーカイブ';
      const status = project.status === 'アーカイブ' ? '完了' : (project.status || '制作中');
      return {
        ...makeProject(),
        ...project,
        status,
        archived,
        archivedAt: project.archivedAt || '',
        analysis: Number(project.analysis?.metricVersion) >= MIN_ANALYSIS_METRIC_VERSION ? project.analysis : null,
        tasks: Object.fromEntries(allTasks.map(task => [task.id, { ...defaults[task.id], ...(project.tasks?.[task.id] || {}) }])),
        logs: Array.isArray(project.logs) ? project.logs.map(normalizeLog) : []
      };
    }

    function escapeHtml(value) {
      return String(value ?? '').replace(/[&<>'"]/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' })[character]);
    }

    function scheduleProjectSave() {
      const project = getActiveProject();
      if (!project) return;
      captureProjectFields(project);
      renderProjectHeader(project);
      project.updatedAt = new Date().toISOString();
      markProjectDirty(project.id);
      clearTimeout(saveTimer);
      saveTimer = setTimeout(() => {
        persistProjects();
        scheduleAutoSync();
      }, 100);
    }

    function renderProjectHeader(project) {
      if (!project) return;
      document.getElementById('project-header-title').textContent = project.title || 'Untitled Track';
      document.getElementById('project-header-status').textContent = project.status || '制作中';
      document.getElementById('project-header-client').textContent = `クライアント: ${project.client || '—'}`;
      document.getElementById('project-header-bpm').textContent = `BPM ${project.bpm || '—'}`;
      document.getElementById('project-header-duration').textContent = `長さ ${project.duration ? `${project.duration}秒` : '—'}`;
      document.getElementById('project-header-deadline').textContent = `締切 ${project.deadline || '—'}`;
    }

    function captureProjectFields(project) {
      project.title = document.getElementById('project-title').value.trim() || 'Untitled Track';
      project.status = document.getElementById('project-status').value;
      project.bpm = document.getElementById('project-bpm').value;
      project.duration = Math.max(1, Number(document.getElementById('project-duration').value) || 180);
      project.client = document.getElementById('project-client').value.trim();
      project.deadline = document.getElementById('project-deadline').value;
      document.getElementById('song-title').value = project.title;
      project.finalA = Number(document.getElementById('final-slider').value);
      project.splitA = project.finalA;
      project.splitB = 100 - project.finalA;
      document.querySelectorAll('.task-row').forEach(row => {
        const amountSlider = row.querySelector('.amount-slider');
        project.tasks[row.dataset.taskId] = {
          enabled: row.querySelector('.task-enabled').checked,
          a: Number(row.querySelector('.task-slider').value),
          amount: amountSlider ? Number(amountSlider.value) : null
        };
      });
    }

    function projectCardMarkup(project, archived = false) {
      const splitA = Number(project.splitA ?? project.finalA ?? 50);
      const splitB = Number(project.splitB ?? (100 - splitA));
      return `
        <article class="flex flex-wrap items-center gap-x-4 gap-y-2 px-4 py-3 transition hover:bg-white/[.04] sm:flex-nowrap">
          <button type="button" data-open-project="${escapeHtml(project.id)}" class="group flex min-w-0 flex-1 items-center gap-3 text-left">
            <span class="shrink-0 rounded-full border border-line px-2 py-1 font-mono text-[10px] text-slate-400">${archived ? 'ARCHIVED' : escapeHtml(project.status)}</span>
            <span class="min-w-0">
              <span class="block truncate text-sm font-bold group-hover:text-acid">${escapeHtml(project.title)}</span>
              <span class="block truncate text-xs text-slate-600">${escapeHtml(project.client || 'NO CLIENT')}</span>
            </span>
          </button>
          <div class="flex shrink-0 items-center gap-4 font-mono text-xs">
            <span class="text-acid">tada ${splitA.toFixed(1)}%</span>
            <span class="text-violet-300">riku ${splitB.toFixed(1)}%</span>
            <span class="text-slate-500">${project.logs?.length || 0}件</span>
            <span class="hidden text-slate-500 sm:inline">${escapeHtml(project.deadline || '—')}</span>
            <span class="hidden text-slate-600 md:inline">${project.updatedAt ? new Date(project.updatedAt).toLocaleDateString('ja-JP') : '—'}</span>
          </div>
          <div class="flex shrink-0 items-center gap-2">
            <button type="button" data-${archived ? 'restore' : 'archive'}-project="${escapeHtml(project.id)}" class="btn btn-secondary btn-sm">${archived ? '復元' : 'アーカイブ'}</button>
            ${archived ? `<button type="button" data-delete-project="${escapeHtml(project.id)}" class="btn btn-danger btn-sm" aria-label="${escapeHtml(project.title)}を消去">消去</button>` : ''}
          </div>
        </article>`;
    }

    function renderDashboard() {
      const activeProjects = projects.filter(project => !project.archived);
      const archivedProjects = projects.filter(project => project.archived);
      const sortByUpdated = (a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt));
      document.getElementById('stat-projects').textContent = projects.length;
      document.getElementById('stat-progress').textContent = activeProjects.filter(project => project.status === '制作中').length;
      document.getElementById('stat-archived').textContent = archivedProjects.length;
      document.getElementById('stat-logs').textContent = projects.reduce((sum, project) => sum + (project.logs?.length || 0), 0);
      document.getElementById('empty-projects').classList.toggle('hidden', activeProjects.length > 0);
      document.getElementById('empty-archives').classList.toggle('hidden', archivedProjects.length > 0);
      document.getElementById('project-grid').innerHTML = activeProjects.slice().sort(sortByUpdated).map(project => projectCardMarkup(project)).join('');
      document.getElementById('archived-grid').innerHTML = archivedProjects.slice().sort(sortByUpdated).map(project => projectCardMarkup(project, true)).join('');
    }
    let suppressHistoryPush = false;
    function showDashboard() {
      const project = getActiveProject();
      if (project) {
        captureProjectFields(project);
        persistProjects();
        if (dirtyProjectIds.has(project.id)) scheduleAutoSync();
      }
      activeProjectId = null;
      document.body.classList.remove('is-project-view');
      document.getElementById('dashboard-view').classList.remove('hidden');
      document.getElementById('project-view').classList.add('hidden');
      renderDashboard();
      window.scrollTo({ top: 0, behavior: 'smooth' });
      if (!suppressHistoryPush) history.pushState({ view: 'dashboard' }, '', '#dashboard');
    }

    function openProject(projectId) {
      const project = projects.find(item => item.id === projectId);
      if (!project) return;
      activeProjectId = project.id;
      if (!suppressHistoryPush) history.pushState({ view: 'project', projectId }, '', `#project=${encodeURIComponent(projectId)}`);
      resetLogForm();
      document.body.classList.add('is-project-view');
      document.getElementById('dashboard-view').classList.add('hidden');
      document.getElementById('project-view').classList.remove('hidden');
      document.getElementById('project-title').value = project.title || '';
      document.getElementById('project-status').value = project.status || '制作中';
      document.getElementById('project-bpm').value = project.bpm || '';
      document.getElementById('project-duration').value = project.duration || 180;
      document.getElementById('project-client').value = project.client || '';
      document.getElementById('project-deadline').value = project.deadline || '';
      document.getElementById('song-title').value = project.title || '';
      renderProjectHeader(project);
      document.getElementById('final-slider').value = Number(project.finalA ?? project.splitA ?? project.analysis?.recommendedA ?? 50);
      document.querySelectorAll('.task-row').forEach(row => {
        const task = project.tasks[row.dataset.taskId] || defaultTaskState()[row.dataset.taskId];
        row.querySelector('.task-enabled').checked = task.enabled !== false;
        row.querySelector('.task-slider').value = Number(task.a ?? 50);
        const amountSlider = row.querySelector('.amount-slider');
        if (amountSlider) amountSlider.value = Number(task.amount ?? 100);
        syncTaskRow(row);
      });
      renderProductionLogs();
      if (project.analysis) renderAnalysis(project.analysis);
      calculateSplit();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    const logTypes = {
      melody: { label: 'メインメロディー', weight: 10 },
      structure: { label: '構成・展開', weight: 10 },
      motif: { label: 'モチーフ／フレーズ', weight: 9 },
      harmony: { label: 'コード／ハーモニー', weight: 8 },
      beat: { label: 'ビート', weight: 7 },
      bass: { label: 'ベース', weight: 7 },
      guitar: { label: 'ギター録音', weight: 6 },
      instrument: { label: '楽器トラック', weight: 5 },
      sound: { label: 'サウンドデザイン', weight: 6 },
      sample: { label: 'サンプル／ループ', weight: 3 },
      mix: { label: 'ミックス', weight: 8 },
      delivery: { label: '編集・納品', weight: 2 }
    };

    function ratioFor(a, b) {
      return a + b > 0 ? a / (a + b) * 100 : null;
    }

    function calculateLocalAnalysis() {
      const logs = getActiveProject()?.logs || [];
      const totals = { A: { count: 0, scope: 0, effort: 0, musical: 0 }, B: { count: 0, scope: 0, effort: 0, musical: 0 } };
      logs.forEach(rawLog => {
        const log = normalizeLog(rawLog);
        const person = log.person === 'B' ? 'B' : 'A';
        const count = Math.max(1, Number(log.count) || 1);
        totals[person].count += count;
        totals[person].scope += count * log.scope;
        totals[person].effort += count * log.effort;
        totals[person].musical += (logTypes[log.type]?.weight || 4) * log.scope;
      });
      const quantityRatios = [
        ratioFor(totals.A.count, totals.B.count),
        ratioFor(totals.A.scope, totals.B.scope),
        ratioFor(totals.A.effort, totals.B.effort)
      ].filter(value => value !== null);
      const quantityA = quantityRatios.length ? quantityRatios.reduce((sum, value) => sum + value, 0) / quantityRatios.length : 50;
      const musicalA = ratioFor(totals.A.musical, totals.B.musical) ?? 50;
      const recommendedA = quantityA * 0.4 + musicalA * 0.6;
      return {
        metricVersion: 2,
        quantityA,
        musicalA,
        recommendedA,
        summary: '記録された物量とカテゴリ別の音楽的中心性を分けて集計した基準値です。Google AI分析では、役割や採用箇所の説明も加味します。',
        evidence: [
          `物量 — tada: ${totals.A.count}本・範囲${totals.A.scope}pt・カロリー${totals.A.effort}pt / riku: ${totals.B.count}本・範囲${totals.B.scope}pt・カロリー${totals.B.effort}pt`,
          `音楽的比重スコア — tada: ${totals.A.musical} / riku: ${totals.B.musical}`
        ],
        quantityDetail: '本数・5段階の貢献範囲・制作負荷（カロリー）を同じ比重で集計',
        musicalDetail: 'メロディー・構成・モチーフなど曲の中心性に貢献範囲を加味',
        source: 'LOCAL BASELINE'
      };
    }

    // `persist` must only be set when `result` is a genuine analysis result. Rendering a
    // locally computed placeholder must never write it into project.analysis: that value
    // is synced, so it would overwrite the real AI analysis stored in the sheet.
    function renderAnalysis(result, { persist = false } = {}) {
      if (!result) return;
      const usePersonNames = value => String(value || '')
        .replace(/Person\s*A/gi, 'tada')
        .replace(/Person\s*B/gi, 'riku')
        .replace(/担当\s*A/g, 'tada')
        .replace(/担当\s*B/g, 'riku')
        .replace(/\bA(?=[はがのにをへもと、。])/g, 'tada')
        .replace(/\bB(?=[はがのにをへもと、。])/g, 'riku')
        .replace(/(^|[\s（(・:：,\/])A(?=$|[\s）)・:：,\/])/g, '$1tada')
        .replace(/(^|[\s（(・:：,\/])B(?=$|[\s）)・:：,\/])/g, '$1riku');
      const quantityA = Math.max(0, Math.min(100, Number(result.quantityA) || 0));
      const musicalA = Math.max(0, Math.min(100, Number(result.musicalA) || 0));
      const hasBeta = Number.isFinite(Number(result.betaTadaPercent));
      const betaTadaPercent = Math.max(0, Math.min(100, Number(result.betaTadaPercent) || 0));
      const recommendedA = Math.max(0, Math.min(100, Number(result.recommendedA) || 0));
      document.getElementById('analysis-empty').classList.add('hidden');
      document.getElementById('analysis-result').classList.remove('hidden');
      document.getElementById('quantity-score').textContent = `tada ${quantityA.toFixed(0)} / riku ${(100 - quantityA).toFixed(0)}`;
      document.getElementById('musical-score').textContent = `tada ${musicalA.toFixed(0)} / riku ${(100 - musicalA).toFixed(0)}`;
      document.getElementById('quantity-bar-a').style.width = `${quantityA}%`;
      document.getElementById('quantity-bar-b').style.width = `${100 - quantityA}%`;
      document.getElementById('musical-bar-a').style.width = `${musicalA}%`;
      document.getElementById('musical-bar-b').style.width = `${100 - musicalA}%`;
      document.getElementById('quantity-detail').textContent = usePersonNames(result.quantityDetail || '本数・貢献範囲・制作負荷から算出');
      document.getElementById('musical-detail').textContent = usePersonNames(result.musicalDetail || '曲の成立への中心性から算出');
      document.getElementById('beta-score').textContent = hasBeta ? `tada ${betaTadaPercent.toFixed(0)} / riku ${(100 - betaTadaPercent).toFixed(0)}` : '未実行';
      document.getElementById('beta-bar-a').style.width = hasBeta ? `${betaTadaPercent}%` : '50%';
      document.getElementById('beta-bar-b').style.width = hasBeta ? `${100 - betaTadaPercent}%` : '50%';
      document.getElementById('beta-detail').textContent = hasBeta ? usePersonNames(result.betaSummary || '要素ごとの5軸評価から算出しました。') : 'Google AIで分析すると、要素ごとの5軸評価から算出されます。';
      document.getElementById('analysis-recommendation').textContent = `tada ${recommendedA.toFixed(1)}% / riku ${(100 - recommendedA).toFixed(1)}%`;
      document.getElementById('analysis-summary').textContent = usePersonNames(result.summary || '');
      document.getElementById('analysis-evidence').innerHTML = (result.evidence || []).map(item => `<li>・${escapeHtml(usePersonNames(item))}</li>`).join('');
      document.getElementById('analysis-source').textContent = result.source || 'GOOGLE AI';
      document.getElementById('beta-detail-btn').disabled = !hasBeta;
      const project = getActiveProject();
      if (project && persist) project.analysis = { ...result, quantityA, musicalA, betaTadaPercent, recommendedA };
    }

    function renderBetaDetailModal() {
      const body = document.getElementById('beta-detail-body');
      const beta = getActiveProject()?.analysis?.beta;
      if (!beta || !Array.isArray(beta.elements) || !beta.elements.length) {
        body.innerHTML = '<p class="rounded-xl border border-dashed border-line p-8 text-center text-sm text-slate-600">まだ5軸音楽分析が実行されていません。「Google AIで分析」を実行すると詳細が表示されます。</p>';
        return;
      }
      const confidenceBadge = value => {
        const percent = Math.round(Math.max(0, Math.min(1, Number(value) || 0)) * 100);
        return `<span class="font-mono text-[9px] ${percent < 70 ? 'text-amber-400' : 'text-slate-600'}">確信度 ${percent}%${percent < 70 ? ' · 要確認' : ''}</span>`;
      };
      const categoryRows = (beta.categoryWeights || []).map(item => `
        <div class="flex items-start justify-between gap-3 border-b border-line py-2 text-xs last:border-0">
          <div class="min-w-0"><span class="font-semibold">${escapeHtml(logTypes[item.category]?.label || item.category)}</span><p class="mt-0.5 text-slate-500">${escapeHtml(item.reason || '')}</p></div>
          <div class="shrink-0 text-right"><span class="font-mono font-bold text-acid">${Number(item.weight).toFixed(1)}%</span><br>${confidenceBadge(item.confidence)}</div>
        </div>`).join('');
      const elementRows = beta.elements.map(element => {
        const contributors = (element.contributors || []).map(c => `
          <div class="rounded-lg border border-line bg-black/20 p-3">
            <div class="flex items-center justify-between gap-2"><span class="font-mono text-xs font-bold ${c.person === 'riku' ? 'text-violet-300' : 'text-acid'}">${escapeHtml(PERSON_NAMES[c.person === 'riku' ? 'B' : 'A'])} · ${escapeHtml(c.role || '')}</span>${confidenceBadge(c.confidence)}</div>
            <p class="mt-1 font-mono text-[10px] text-slate-500">役割${c.roleScore} ・ 採用度${c.adoptionScore} ・ 代替不可能性${c.irreplaceabilityScore}</p>
            <p class="mt-1 text-xs leading-5 text-slate-500">${escapeHtml(c.evidence || '')}</p>
          </div>`).join('');
        return `
        <div class="rounded-xl border border-line p-4">
          <div class="flex items-start justify-between gap-3"><div><span class="rounded bg-white/[.04] px-2 py-0.5 text-[9px] text-slate-500">${escapeHtml(logTypes[element.category]?.label || element.category)}</span><p class="mt-1 text-sm font-bold">${escapeHtml(element.name)}</p></div>${confidenceBadge(element.confidence)}</div>
          <p class="mt-1 font-mono text-[10px] text-slate-500">同一性${element.identityScore} ・ 影響範囲${element.scopeScore}</p>
          <p class="mt-2 text-xs leading-5 text-slate-500">${escapeHtml(element.evidence || '')}</p>
          <div class="mt-3 grid gap-2 sm:grid-cols-2">${contributors}</div>
        </div>`;
      }).join('');
      body.innerHTML = `
        <p class="text-sm leading-7 text-slate-400">${escapeHtml(beta.summary || '')}</p>
        <h3 class="mt-5 text-xs font-bold uppercase tracking-wide text-slate-500">カテゴリ重要度</h3>
        <div class="mt-2 rounded-xl border border-line p-4">${categoryRows}</div>
        <h3 class="mt-5 text-xs font-bold uppercase tracking-wide text-slate-500">音楽要素ごとの評価</h3>
        <div class="mt-2 space-y-3">${elementRows}</div>`;
    }

    function getLogDateKey(log) {
      if (!log) return '';
      const d = new Date(log.createdAt || 0);
      if (isNaN(d.getTime()) || d.getTime() === 0) return '';
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }

    function formatLogDateDivider(dateKey) {
      if (!dateKey) return '';
      const [y, m, d] = dateKey.split('-').map(Number);
      const targetDate = new Date(y, m - 1, d);
      if (isNaN(targetDate.getTime())) return dateKey;
      const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
      const weekday = weekdays[targetDate.getDay()];
      
      const now = new Date();
      const isToday = now.getFullYear() === y && (now.getMonth() + 1) === m && now.getDate() === d;
      
      if (isToday) return `${y}年${m}月${d}日(${weekday}) 今日`;
      return `${y}年${m}月${d}日(${weekday})`;
    }

    function formatLogTime(dateStr) {
      if (!dateStr) return '';
      const d = new Date(dateStr);
      if (isNaN(d.getTime()) || d.getTime() === 0) return '';
      return d.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' });
    }

    function renderProductionLogs() {
      const project = getActiveProject();
      if (!project) return;
      project.logs = project.logs.map(normalizeLog);
      const totals = project.logs.reduce((summary, log) => {
        const person = log.person === 'B' ? 'B' : 'A';
        const count = Math.max(1, Number(log.count) || 1);
        summary[person].count += count;
        summary[person].scope += count * log.scope;
        summary[person].effort += count * log.effort;
        return summary;
      }, { A: { count: 0, scope: 0, effort: 0 }, B: { count: 0, scope: 0, effort: 0 } });
      document.getElementById('log-summary').innerHTML = ['A', 'B'].map(person => `<div class="rounded-lg border border-line bg-black/20 px-3 py-2 text-xs leading-4" title="${PERSON_NAMES[person]}：${totals[person].count}本 / 範囲${totals[person].scope}pt / カロリー${totals[person].effort}pt"><span class="font-mono font-bold ${person === 'A' ? 'text-acid' : 'text-violet-300'}">${PERSON_NAMES[person]}</span><span class="ml-2 text-slate-400">${totals[person].count}本 · 範囲${totals[person].scope} · カロリー${totals[person].effort}</span></div>`).join('') + `<div class="rounded-lg border border-line bg-black/20 px-3 py-2 text-xs leading-4"><span class="font-mono font-bold text-slate-400">TOTAL</span><span class="ml-2 text-slate-400">${project.logs.length}ログ</span></div>`;
      const list = document.getElementById('log-list');
      document.getElementById('empty-logs').classList.toggle('hidden', project.logs.length > 0);

      // 上が最新（降順）にソート
      const sortedLogs = project.logs
        .map((log, index) => ({
          log,
          index,
          time: new Date(log.createdAt || 0).getTime() || index
        }))
        .sort((a, b) => b.time - a.time);

      let lastDateKey = null;
      const htmlParts = [];

      sortedLogs.forEach(({ log }) => {
        const dateKey = getLogDateKey(log);
        if (dateKey && dateKey !== lastDateKey) {
          htmlParts.push(`
            <div class="my-3.5 flex items-center justify-center gap-3">
              <span class="h-px flex-1 bg-line/60"></span>
              <span class="rounded-full border border-line bg-white/[.04] px-3.5 py-1 font-mono text-xs font-semibold text-slate-400 backdrop-blur-sm">
                ${escapeHtml(formatLogDateDivider(dateKey))}
              </span>
              <span class="h-px flex-1 bg-line/60"></span>
            </div>
          `);
          lastDateKey = dateKey;
        }

        const timeStr = formatLogTime(log.createdAt);

        htmlParts.push(`
          <div data-log-id="${escapeHtml(log.id)}" tabindex="0" title="${escapeHtml(log.details || 'ダブルクリックで編集')}" class="production-log-row flex items-center justify-between gap-4 rounded-xl border border-line bg-white/[.015] px-4 py-3 transition hover:border-slate-600 focus:outline-none focus:ring-1 focus:ring-acid/40">
            <div class="flex min-w-0 flex-1 items-center gap-3.5">
              <span class="grid h-8 w-8 shrink-0 place-items-center rounded-lg ${log.person === 'A' ? 'bg-acid/10 text-acid' : 'bg-violet/10 text-violet-300'} font-mono text-xs font-bold">${escapeHtml(PERSON_NAMES[log.person] || log.person)}</span>
              <div class="min-w-0 flex-1">
                <div class="flex flex-wrap items-center gap-2">
                  <p class="truncate text-sm font-bold text-slate-100">${escapeHtml(log.name)}</p>
                  <span class="shrink-0 rounded bg-white/[.05] px-2 py-0.5 text-[10px] text-slate-400">${escapeHtml(logTypes[log.type]?.label || log.type)}</span>
                </div>
                <p class="mt-1 truncate font-mono text-xs text-slate-400">${log.count}本 · 範囲${log.scope}:${SCOPE_LEVELS[log.scope].label} · カロリー${log.effort}:${EFFORT_LEVELS[log.effort].label}${timeStr ? ` · ${timeStr}` : ''}</p>
              </div>
            </div>
            <div class="flex shrink-0 items-center justify-end gap-2.5">
              <button type="button" data-edit-log="${escapeHtml(log.id)}" class="btn btn-ghost btn-sm px-2.5 py-1 text-xs">編集</button>
              <button type="button" data-remove-log="${escapeHtml(log.id)}" class="btn btn-danger btn-sm px-2.5 py-1 text-xs">削除</button>
            </div>
          </div>
        `);
      });

      list.innerHTML = htmlParts.join('');
      document.getElementById('analysis-empty').classList.toggle('hidden', project.logs.length > 0);
      document.getElementById('analyze-logs-btn').disabled = project.logs.length === 0;
      if (project.logs.length) renderAnalysis(project.analysis || calculateLocalAnalysis());
      else document.getElementById('analysis-result').classList.add('hidden');
    }

    let apiConnectionResolve = null;

    const API_MODAL_COPY = {
      normal: { title: 'Google Sheets / AI 接続設定', desc: 'GitHub Pagesでは初回のみ再設定が必要です。Apps Scriptは「実行ユーザー: 自分」「アクセスできるユーザー: 全員」でデプロイしてください。接続キーはタブを閉じると消去されます。' },
      gate: { title: 'ログインしてください', desc: '制作ログを閲覧するにはGoogle Sheets接続が必要です。接続キーはタブを閉じると消去されるため、次回開いたときも再入力が必要です。' }
    };
    function requestApiConnection(options = {}) {
      const gate = options.gate === true;
      const force = options.force === true;
      if (location.protocol === 'file:') {
        showToast('Google連携はGitHub PagesまたはローカルHTTPサーバーで利用してください');
        return Promise.resolve(null);
      }
      const savedUrl = localStorage.getItem(API_URL_KEY) || '';
      const sessionKey = sessionStorage.getItem('splitlab_api_key') || '';
      if (!force && savedUrl && sessionKey) return Promise.resolve({ apiUrl: savedUrl, apiKey: sessionKey });
      document.getElementById('api-url-input').value = savedUrl;
      document.getElementById('api-key-input').value = '';
      document.getElementById('api-modal-error').classList.add('hidden');
      const copy = gate ? API_MODAL_COPY.gate : API_MODAL_COPY.normal;
      document.getElementById('api-modal-title').textContent = copy.title;
      document.getElementById('api-modal-desc').textContent = copy.desc;
      document.getElementById('api-cancel').classList.toggle('hidden', gate);
      const modal = document.getElementById('api-modal');
      modal.dataset.gate = gate ? 'true' : 'false';
      modal.classList.remove('hidden');
      modal.classList.add('flex');
      setSyncStatus('disconnected');
      setTimeout(() => (savedUrl ? document.getElementById('api-key-input') : document.getElementById('api-url-input')).focus(), 0);
      return new Promise(resolve => { apiConnectionResolve = resolve; });
    }

    function closeApiModal(connection = null) {
      const modal = document.getElementById('api-modal');
      modal.classList.add('hidden');
      modal.classList.remove('flex');
      if (apiConnectionResolve) apiConnectionResolve(connection);
      apiConnectionResolve = null;
    }

    let pendingDeleteProjectId = null;

    function openDeleteModal(project) {
      const apiUrl = localStorage.getItem(API_URL_KEY) || '';
      if (!apiUrl) {
        showToast('先に「自動同期」からGoogle Sheets接続を設定してください');
        return;
      }
      pendingDeleteProjectId = project.id;
      document.getElementById('delete-project-name').textContent = `「${project.title}」`;
      document.getElementById('delete-api-key').value = '';
      document.getElementById('delete-modal-error').classList.add('hidden');
      const modal = document.getElementById('delete-modal');
      modal.classList.remove('hidden');
      modal.classList.add('flex');
      setTimeout(() => document.getElementById('delete-api-key').focus(), 0);
    }

    function closeDeleteModal() {
      const modal = document.getElementById('delete-modal');
      modal.classList.add('hidden');
      modal.classList.remove('flex');
      pendingDeleteProjectId = null;
    }
    let editingLogId = null;
    let parsedLogDrafts = [];

    function resetLogForm() {
      editingLogId = null;
      document.getElementById('log-name').value = '';
      document.getElementById('log-details').value = '';
      document.getElementById('log-count').value = '1';
      document.getElementById('log-scope').value = '3';
      document.getElementById('log-effort').value = '3';
      document.getElementById('log-submit-btn').textContent = '追加';
      document.getElementById('cancel-log-edit').classList.add('hidden');
    }

    function beginLogEdit(logId) {
      const log = getActiveProject()?.logs.find(item => String(item.id) === String(logId));
      if (!log) return;
      editingLogId = log.id;
      document.getElementById('log-person').value = log.person === 'B' ? 'B' : 'A';
      document.getElementById('log-type').value = logTypes[log.type] ? log.type : 'instrument';
      document.getElementById('log-name').value = log.name || '';
      document.getElementById('log-count').value = String(Math.max(1, Number(log.count) || 1));
      document.getElementById('log-scope').value = String(normalizeLevel(log.scope));
      document.getElementById('log-effort').value = String(normalizeLevel(log.effort));
      document.getElementById('log-details').value = log.details || '';
      document.getElementById('log-submit-btn').textContent = '変更を保存';
      document.getElementById('cancel-log-edit').classList.remove('hidden');
      openLogFormModal();
      setTimeout(() => document.getElementById('log-name').focus(), 250);
    }

    function openLogFormModal() {
      const modal = document.getElementById('log-form-modal');
      modal.classList.remove('hidden');
      modal.classList.add('flex');
    }
    function closeLogFormModal() {
      const modal = document.getElementById('log-form-modal');
      modal.classList.add('hidden');
      modal.classList.remove('flex');
    }

    function levelOptions(levels, selected) {
      return Object.entries(levels).map(([value, meta]) => `<option value="${value}" ${Number(value) === Number(selected) ? 'selected' : ''}>${value} · ${meta.label}</option>`).join('');
    }

    function renderLogDrafts() {
      const container = document.getElementById('log-draft-list');
      const typeOptions = Object.entries(logTypes).map(([value, meta]) => `<option value="${value}">${meta.label}</option>`).join('');
      container.innerHTML = parsedLogDrafts.map((draft, index) => `
        <div class="draft-row rounded-xl border border-line bg-black/25 p-4" data-draft-index="${index}">
          <div class="flex items-start gap-3">
            <input type="checkbox" class="draft-selected mt-1 h-4 w-4 accent-[#c7ff3d]" checked aria-label="このログ候補を登録">
            <div class="min-w-0 flex-1">
              <div class="grid gap-3 sm:grid-cols-[.55fr_1fr_1.6fr]">
                <label><span class="mb-1 block text-[9px] text-slate-600">担当</span><select class="draft-person w-full rounded-lg border border-line bg-[#0d1016] px-3 py-2 text-xs"><option value="A" ${draft.person === 'A' ? 'selected' : ''}>tada</option><option value="B" ${draft.person === 'B' ? 'selected' : ''}>riku</option></select></label>
                <label><span class="mb-1 block text-[9px] text-slate-600">カテゴリ</span><select class="draft-type w-full rounded-lg border border-line bg-[#0d1016] px-3 py-2 text-xs">${typeOptions}</select></label>
                <label><span class="mb-1 block text-[9px] text-slate-600">内容</span><input class="draft-name w-full rounded-lg border border-line bg-black/30 px-3 py-2 text-xs" value="${escapeHtml(draft.name)}"></label>
              </div>
              <div class="mt-3 grid gap-3 sm:grid-cols-3">
                <label><span class="mb-1 block text-[9px] text-slate-600">本数</span><input type="number" min="1" class="draft-count w-full rounded-lg border border-line bg-black/30 px-3 py-2 text-xs" value="${draft.count}"></label>
                <label><span class="mb-1 block text-[9px] text-slate-600">貢献範囲</span><select class="draft-scope w-full rounded-lg border border-line bg-[#0d1016] px-3 py-2 text-xs">${levelOptions(SCOPE_LEVELS, draft.scope)}</select></label>
                <label><span class="mb-1 block text-[9px] text-slate-600">制作負荷（カロリー）</span><select class="draft-effort w-full rounded-lg border border-line bg-[#0d1016] px-3 py-2 text-xs">${levelOptions(EFFORT_LEVELS, draft.effort)}</select></label>
              </div>
              <label class="mt-3 block"><span class="mb-1 block text-[9px] text-slate-600">役割・採用箇所</span><textarea rows="2" class="draft-details w-full resize-y rounded-lg border border-line bg-black/30 px-3 py-2 text-xs leading-5">${escapeHtml(draft.details)}</textarea></label>
              ${draft.uncertainFields?.length ? `<p class="mt-2 text-[10px] text-amber-400">要確認: ${escapeHtml(draft.uncertainFields.join('、'))}</p>` : ''}
            </div>
          </div>
        </div>`).join('');
      container.querySelectorAll('.draft-row').forEach((row, index) => {
        row.querySelector('.draft-type').value = parsedLogDrafts[index].type;
      });
      document.getElementById('reopen-drafts-btn').classList.toggle('hidden', parsedLogDrafts.length === 0);
    }

    function openLogDraftsModal() {
      const modal = document.getElementById('log-drafts-modal');
      modal.classList.remove('hidden');
      modal.classList.add('flex');
    }
    function closeLogDraftsModal() {
      const modal = document.getElementById('log-drafts-modal');
      modal.classList.add('hidden');
      modal.classList.remove('flex');
    }

    async function parseNarrativeWithGoogle(textareaId = 'log-narrative', buttonId = 'parse-narrative-btn', statusId = 'parse-status') {
      const textarea = document.getElementById(textareaId);
      const text = textarea.value.trim();
      const status = document.getElementById(statusId);
      if (!text) {
        status.textContent = '制作内容を文章で入力してください。';
        status.className = 'min-h-5 text-xs text-amber-400';
        return;
      }
      const connection = await requestApiConnection();
      if (!connection) return;
      const button = document.getElementById(buttonId);
      button.disabled = true;
      button.textContent = '分解中…';
      status.textContent = '文章から制作ログを抽出しています…';
      status.className = 'min-h-5 text-xs text-slate-500';
      try {
        const project = getActiveProject();
        const body = new URLSearchParams({ action: 'parseLogs', apiKey: connection.apiKey, payload: JSON.stringify({ text, project: { title: project?.title || '', duration: project?.duration || 180 } }) });
        const result = await postToAppsScript(connection.apiUrl, body);
        if (!result.ok) throw new Error(result.error || '文章を分解できませんでした。');
        parsedLogDrafts = (result.logs || []).map(log => ({
          person: log.person === 'B' ? 'B' : 'A',
          type: logTypes[log.type] ? log.type : 'instrument',
          name: String(log.name || '制作作業'),
          count: Math.max(1, Number(log.count) || 1),
          scope: normalizeLevel(log.scope),
          effort: normalizeLevel(log.effort),
          details: String(log.details || ''),
          uncertainFields: Array.isArray(log.uncertainFields) ? log.uncertainFields : []
        }));
        renderLogDrafts();
        status.textContent = `${parsedLogDrafts.length}件のログ候補に分解しました。内容を確認してください。`;
        status.className = 'min-h-5 text-xs text-acid';
        if (parsedLogDrafts.length) openLogDraftsModal();
      } catch (error) {
        parsedLogDrafts = [];
        renderLogDrafts();
        closeLogDraftsModal();
        status.textContent = `AI分解エラー: ${error.message}`;
        status.className = 'min-h-5 text-xs text-rose-400';
      } finally {
        button.disabled = false;
        button.textContent = 'AIで分解';
      }
    }
    async function analyzeLogsWithGoogle() {
      const project = getActiveProject();
      if (!project || !project.logs.length) return;
      const connection = await requestApiConnection();
      if (!connection) return;
      const { apiUrl, apiKey } = connection;
      const button = document.getElementById('analyze-logs-btn');
      button.disabled = true;
      button.textContent = '分析中…';
      document.getElementById('analysis-empty').classList.add('hidden');
      document.getElementById('analysis-result').classList.add('hidden');
      document.getElementById('analysis-loading').classList.remove('hidden');
      document.getElementById('analysis-loading').classList.add('flex');
      try {
        const payload = { project: { id: project.id, title: project.title, duration: project.duration }, logs: project.logs, baseline: calculateLocalAnalysis() };
        const body = new URLSearchParams({ action: 'analyzeCombined', apiKey, payload: JSON.stringify(payload) });
        const result = await postToAppsScript(apiUrl, body, { timeoutMs: 90000, timeoutLabel: 'AI分析' });
        if (!result.ok) {
          if (result.error === 'Unsupported action.') throw new Error('Apps Scriptが旧バージョンです。最新のCode.gsを新しいバージョンとして再デプロイすると5軸音楽分析も実行されます。');
          throw new Error(result.error || '分析に失敗しました。');
        }
        renderAnalysis({ ...result.analysis, source: 'GOOGLE GEMINI' }, { persist: true });
        scheduleProjectSave();
        showToast('Google AIによるログ分析が完了しました（5軸音楽分析を含む）');
      } catch (error) {
        renderAnalysis({ ...calculateLocalAnalysis(), source: 'LOCAL FALLBACK' });
        showToast(`AI分析エラー: ${error.message}`, { persistent: true });
      } finally {
        document.getElementById('analysis-loading').classList.add('hidden');
        document.getElementById('analysis-loading').classList.remove('flex');
        button.disabled = false;
        button.textContent = 'Google AIで分析';
      }
    }

    function getStoredApiConnection() {
      if (location.protocol === 'file:') return null;
      const apiUrl = localStorage.getItem(API_URL_KEY) || '';
      const apiKey = sessionStorage.getItem('splitlab_api_key') || '';
      return apiUrl && apiKey ? { apiUrl, apiKey } : null;
    }

    const SYNC_STATUS_META = {
      disconnected: { dotClass: 'is-disconnected', text: '未接続' },
      syncing: { dotClass: 'is-syncing', text: '同期中…' },
      synced: { dotClass: 'is-synced', text: '同期済み' },
      error: { dotClass: 'is-error', text: '同期エラー' }
    };
    function setSyncStatus(state, detail) {
      const meta = SYNC_STATUS_META[state] || SYNC_STATUS_META.disconnected;
      document.getElementById('sync-status-dot').className = `sync-dot ${meta.dotClass}`;
      document.getElementById('sync-status-label').textContent = detail || meta.text;
    }
    function refreshSyncStatus() {
      setSyncStatus(getStoredApiConnection() ? 'synced' : 'disconnected');
    }

    function applyTheme(theme) {
      document.documentElement.dataset.theme = theme;
      localStorage.setItem('splitlab_theme', theme);
      document.getElementById('theme-toggle').textContent = theme === 'dark' ? '☀️' : '🌙';
    }

    async function postToAppsScript(apiUrl, body, options = {}) {
      const timeoutMs = Math.max(1000, Number(options.timeoutMs) || 60000);
      const timeoutLabel = String(options.timeoutLabel || '処理');
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
      let response;
      let responseText;
      try {
        response = await fetch(apiUrl, { method: 'POST', body, signal: controller.signal });
        responseText = await response.text();
      } catch (error) {
        if (error && error.name === 'AbortError') {
          throw new Error(`${timeoutLabel}が${Math.round(timeoutMs / 1000)}秒以内に完了しなかったため中断しました。自動的に再試行します。`);
        }
        throw new Error('Apps Scriptへ接続できません。ウェブアプリを「実行ユーザー: 自分」「アクセスできるユーザー: 全員」で再デプロイし、/execで終わるURLを設定してください。');
      } finally {
        clearTimeout(timeoutId);
      }
      let result;
      try {
        result = JSON.parse(responseText);
      } catch {
        const looksLikeLoginPage = /<!doctype|<html|accounts\.google\.com|ServiceLogin/i.test(responseText);
        if (looksLikeLoginPage) {
          throw new Error('Apps Scriptがログイン画面を返しました。アクセスできるユーザーを「全員」にして新しいバージョンをデプロイしてください。');
        }
        throw new Error('Apps Scriptの応答を読み取れませんでした（HTTP ' + response.status + '）。Web App URLが/execで終わるか確認してください。');
      }
      if (!response.ok) throw new Error(result.error || 'Apps Script HTTP ' + response.status);
      return result;
    }

    function scheduleAutoSync(delay = 500) {
      if (!getStoredApiConnection()) return;
      clearTimeout(autoSyncTimer);
      autoSyncTimer = setTimeout(() => syncWithSheets({ silent: true }), delay);
    }

    function applyDeltaSync(remoteProjects, deletedProjectIds, full) {
      const deletedIds = new Set((deletedProjectIds || []).map(String));
      const localById = new Map(projects.map(project => [String(project.id), project]));
      const nextById = full
        ? new Map()
        : new Map([...localById].filter(([projectId]) => !deletedIds.has(projectId)));

      (remoteProjects || []).forEach(remoteProject => {
        const projectId = String(remoteProject.id);
        if (!projectId || deletedIds.has(projectId) || dirtyProjectIds.has(projectId)) return;
        nextById.set(projectId, normalizeProject(remoteProject));
      });

      if (full) {
        localById.forEach((localProject, projectId) => {
          if (!dirtyProjectIds.has(projectId) || deletedIds.has(projectId)) return;
          // A dirty local copy wins so unsynced edits are not lost, but it must not
          // drag the analysis back to nothing: keep the remote one when it has an
          // analysis and the local copy does not.
          const remote = nextById.get(projectId);
          nextById.set(projectId, remote && remote.analysis && !localProject.analysis
            ? { ...localProject, analysis: remote.analysis }
            : localProject);
        });
      }
      deletedIds.forEach(projectId => {
        nextById.delete(projectId);
        dirtyProjectIds.delete(projectId);
      });
      projects = [...nextById.values()];
    }

    async function syncWithSheets(options = {}) {
      const silent = options && options.silent === true;
      const connection = silent ? getStoredApiConnection() : await requestApiConnection();
      if (!connection) { setSyncStatus('disconnected'); return false; }
      if (syncInFlight) {
        if (!options.poll) syncQueued = true;
        return false;
      }
      setSyncStatus('syncing');

      const activeProject = getActiveProject();
      if (!silent && activeProject) {
        captureProjectFields(activeProject);
        persistProjects();
      }

      const full = options.full === true || !lastRemoteSyncAt || projects.length === 0;
      const attemptedProjects = projects.filter(project => dirtyProjectIds.has(project.id));
      const attemptedVersions = Object.fromEntries(attemptedProjects.map(project => [String(project.id), project.updatedAt || '']));
      syncInFlight = true;
      const { apiUrl, apiKey } = connection;
      const button = document.getElementById('sync-btn');
      if (!silent) {
        button.disabled = true;
        button.textContent = full ? '全件取得中…' : '同期中…';
      }
      if (full) startFullSyncProgress();

      try {
        if (full) setFullSyncProgress(18, '変更した案件を送信しています');
        const payload = {
          changes: attemptedProjects,
          since: full ? '' : lastRemoteSyncAt,
          full
        };
        const body = new URLSearchParams({ action: 'syncDelta', apiKey, payload: JSON.stringify(payload) });
        const result = await postToAppsScript(apiUrl, body, {
          timeoutMs: full ? 30000 : 15000,
          timeoutLabel: full ? '全件取得' : '同期'
        });
        if (!result.ok) {
          if (result.error === 'Unsupported action.') throw new Error('Apps Scriptが旧バージョンです。最新のCode.gsを新しいバージョンとして再デプロイしてください。');
          throw new Error(result.error || '同期に失敗しました。');
        }
        if (full) setFullSyncProgress(88, '取得したデータを反映しています');

        attemptedProjects.forEach(project => {
          const currentProject = projects.find(item => String(item.id) === String(project.id));
          if (!currentProject || String(currentProject.updatedAt || '') === String(attemptedVersions[String(project.id)] || '')) {
            dirtyProjectIds.delete(project.id);
          }
        });
        const remoteProjects = Array.isArray(result.projects) ? result.projects : [];
        const deletedProjectIds = Array.isArray(result.deletedProjectIds) ? result.deletedProjectIds : [];
        const receivedChanges = full || remoteProjects.length > 0 || deletedProjectIds.length > 0;
        applyDeltaSync(remoteProjects, deletedProjectIds, full);
        lastRemoteSyncAt = result.serverTime || new Date().toISOString();
        persistSyncState();
        persistProjects();
        if (receivedChanges) renderDashboard();
        // The open project holds its own rendered copy, so refresh it too — otherwise
        // an analysis that just arrived only appears after navigating away and back.
        if (receivedChanges && activeProjectId) {
          const refreshed = getActiveProject();
          if (refreshed) {
            renderProductionLogs();
            if (refreshed.analysis) renderAnalysis(refreshed.analysis);
          }
        }
        if (full) finishFullSyncProgress(true);
        if (!silent) showToast(full ? `${projects.length}件を全件取得しました` : `${remoteProjects.length}件の差分を同期しました`);
        setSyncStatus('synced');
        return true;
      } catch (error) {
        if (full) finishFullSyncProgress(false);
        if (!silent) showToast(`同期エラー: ${error.message}`);
        setSyncStatus('error');
        return false;
      } finally {
        syncInFlight = false;
        if (!silent) {
          button.disabled = false;
          button.textContent = '自動同期';
        }
        if (syncQueued) {
          syncQueued = false;
          scheduleAutoSync(150);
        }
      }
    }
    function taskTemplate(task) {
      return `
        <div class="task-row grid gap-4 px-5 py-5 transition sm:grid-cols-[minmax(180px,1fr)_minmax(250px,1.5fr)] sm:items-center sm:px-6" data-task-id="${task.id}" data-points="${task.points}" data-adjustable="${Boolean(task.adjustable)}">
          <div class="flex items-center gap-3">
            <label class="relative inline-flex cursor-pointer items-center" title="${task.name}を計算に含める">
              <input type="checkbox" class="task-enabled peer sr-only" checked aria-label="${task.name}を有効にする">
              <span class="h-5 w-9 rounded-full bg-slate-700 transition peer-checked:bg-acid/80 peer-focus-visible:ring-2 peer-focus-visible:ring-acid peer-focus-visible:ring-offset-2 peer-focus-visible:ring-offset-panel after:absolute after:left-[3px] after:top-[3px] after:h-3.5 after:w-3.5 after:rounded-full after:bg-white after:transition peer-checked:after:translate-x-4"></span>
            </label>
            <div class="min-w-0">
              <p class="truncate text-sm font-semibold text-slate-200">${task.name}</p>
              <p class="mt-1 font-mono text-[10px] text-slate-600">${String(task.points).padStart(2, '0')} ${task.adjustable ? 'MAX POINTS' : 'POINTS'}</p>
            </div>
          </div>
          <div>
            <div class="mb-3 flex items-center justify-between font-mono text-[11px] font-semibold">
              <span class="text-acid">tada <strong class="value-a ml-1">50%</strong></span>
              <span class="text-violet-300"><strong class="value-b mr-1">50%</strong> riku</span>
            </div>
            <input type="range" min="0" max="100" value="50" step="1" class="range task-slider w-full" aria-label="${task.name}のtadaの貢献割合">
            ${task.adjustable ? `
              <div class="mt-4 flex items-center gap-3 border-t border-line/70 pt-3">
                <span class="shrink-0 font-mono text-[10px] text-slate-500">制作量</span>
                <input type="range" min="0" max="100" value="100" step="10" class="amount-slider h-1 w-full cursor-pointer appearance-none rounded-full bg-slate-700 accent-slate-300" aria-label="${task.name}の制作量">
                <span class="amount-value w-10 text-right font-mono text-[10px] font-semibold text-slate-400">100%</span>
              </div>` : ''}
          </div>
        </div>`;
    }

    categories.forEach(category => {
      const section = document.createElement('section');
      section.className = 'glass overflow-hidden rounded-2xl border border-line';
      section.innerHTML = `
        <div class="flex items-center justify-between border-b border-line bg-white/[.015] px-5 py-4 sm:px-6">
          <div class="flex items-center gap-3">
            <span class="font-mono text-[10px] text-slate-600">${category.code}</span>
            <h3 class="text-sm font-bold">${category.name}</h3>
          </div>
          <span class="font-mono text-[10px] text-slate-600">${category.tasks.reduce((sum, task) => sum + task.points, 0)} PT</span>
        </div>
        <div class="divide-y divide-line">${category.tasks.map(taskTemplate).join('')}</div>`;
      taskList.appendChild(section);
    });

    function calculateSplit() {
      const slider = document.getElementById('final-slider');
      const a = Math.max(0, Math.min(100, Number(slider.value) || 0));
      const b = 100 - a;
      slider.style.setProperty('--fill', `${a}%`);
      document.getElementById('agreement-a').textContent = `${a.toFixed(1)}%`;
      document.getElementById('agreement-b').textContent = `${b.toFixed(1)}%`;
      const project = getActiveProject();
      if (project) {
        project.finalA = a;
        project.splitA = a;
        project.splitB = b;
      }
      renderSplit(a, b, 100);
      return { a, b, totalPoints: 100 };
    }

    function renderSplit(a, b, totalPoints) {
      const analysis = getActiveProject()?.analysis;
      const pointsText = analysis
        ? `分析推奨 tada ${Number(analysis.recommendedA).toFixed(1)}% / 最終合意を優先`
        : 'ログ分析は参考値 / 最終合意を優先';
      
      const barA = document.getElementById('bar-a');
      const barB = document.getElementById('bar-b');
      if (barA) barA.style.width = `${a}%`;
      if (barB) barB.style.width = `${b}%`;

      const labelA = document.getElementById('split-percent-a');
      const labelB = document.getElementById('split-percent-b');
      if (labelA) labelA.textContent = `${a.toFixed(1)}%`;
      if (labelB) labelB.textContent = `${b.toFixed(1)}%`;

      const oldLabelA = document.getElementById('bar-a-label');
      const oldLabelB = document.getElementById('bar-b-label');
      if (oldLabelA) oldLabelA.innerHTML = a >= 12 ? `tada&nbsp; ${a.toFixed(1)}%` : a >= 5 ? `${a.toFixed(0)}%` : '';
      if (oldLabelB) oldLabelB.innerHTML = b >= 12 ? `riku&nbsp; ${b.toFixed(1)}%` : b >= 5 ? `${b.toFixed(0)}%` : '';

      document.getElementById('active-points').textContent = pointsText;
    }

    function syncTaskRow(row) {
      const slider = row.querySelector('.task-slider');
      const a = Number(slider.value);
      slider.style.setProperty('--fill', `${a}%`);
      row.querySelector('.value-a').textContent = `${a}%`;
      row.querySelector('.value-b').textContent = `${100 - a}%`;
      const amountSlider = row.querySelector('.amount-slider');
      if (amountSlider) row.querySelector('.amount-value').textContent = `${amountSlider.value}%`;
      row.classList.toggle('is-disabled', !row.querySelector('.task-enabled').checked);
    }

    function formatPoints(points) {
      return Number.isInteger(points) ? String(points) : points.toFixed(1);
    }

    taskList.addEventListener('input', event => {
      const row = event.target.closest('.task-row');
      if (!row) return;
      syncTaskRow(row);
      calculateSplit();
      scheduleProjectSave();
    });
    taskList.addEventListener('change', event => {
      const row = event.target.closest('.task-row');
      if (!row) return;
      syncTaskRow(row);
      calculateSplit();
      scheduleProjectSave();
    });

    function buildCopyText() {
      const result = calculateSplit();
      const project = getActiveProject();
      const analysis = project?.analysis;
      const logLines = (project?.logs || []).map(rawLog => {
        const log = normalizeLog(rawLog);
        return `${PERSON_NAMES[log.person] || log.person} / ${logTypes[log.type]?.label || log.type} / ${log.name} / ${log.count}本 / 貢献範囲 ${log.scope}:${SCOPE_LEVELS[log.scope].label} / 制作負荷 ${log.effort}:${EFFORT_LEVELS[log.effort].label} / ${log.details || '詳細なし'}`;
      });
      return [
        `SPLIT DATA — ${project?.title || 'Untitled Track'}`,
        '================================',
        '',
        '[制作ログ]',
        ...(logLines.length ? logLines : ['記録なし']),
        '',
        '[分析結果]',
        analysis ? `物量: tada ${Number(analysis.quantityA).toFixed(1)}% / riku ${(100 - Number(analysis.quantityA)).toFixed(1)}%` : '未分析',
        analysis ? `音楽的比重: tada ${Number(analysis.musicalA).toFixed(1)}% / riku ${(100 - Number(analysis.musicalA)).toFixed(1)}%` : '',
        analysis ? `推奨値: tada ${Number(analysis.recommendedA).toFixed(1)}% / riku ${(100 - Number(analysis.recommendedA)).toFixed(1)}%` : '',
        '',
        '--------------------------------',
        `FINAL AGREEMENT: tada ${result.a.toFixed(1)}% / riku ${result.b.toFixed(1)}%`
      ].filter(line => line !== '').join('\n');
    }

    document.querySelectorAll('.copy-split-btn').forEach(btn => btn.addEventListener('click', async () => {
      const text = buildCopyText();
      try {
        await navigator.clipboard.writeText(text);
        showToast('スプリットデータをコピーしました');
      } catch {
        const area = document.createElement('textarea');
        area.value = text;
        area.style.position = 'fixed';
        area.style.opacity = '0';
        document.body.appendChild(area);
        area.select();
        document.execCommand('copy');
        area.remove();
        showToast('スプリットデータをコピーしました');
      }
    }));

    let toastTimer;
    function hideToast() {
      document.getElementById('toast').classList.add('opacity-0', 'translate-y-3');
    }
    function showToast(message, options = {}) {
      const persistent = options.persistent === true;
      const toast = document.getElementById('toast');
      document.getElementById('toast-message').textContent = message;
      document.getElementById('toast-close').classList.toggle('hidden', !persistent);
      toast.style.pointerEvents = persistent ? 'auto' : '';
      toast.classList.remove('opacity-0', 'translate-y-3');
      clearTimeout(toastTimer);
      if (!persistent) toastTimer = setTimeout(hideToast, 2200);
    }
    document.getElementById('toast-close').addEventListener('click', () => {
      clearTimeout(toastTimer);
      hideToast();
    });

    document.getElementById('api-form').addEventListener('submit', event => {
      event.preventDefault();
      const apiUrl = document.getElementById('api-url-input').value.trim();
      const apiKey = document.getElementById('api-key-input').value;
      if (!/^https:\/\/script\.google\.com\//.test(apiUrl)) {
        const error = document.getElementById('api-modal-error');
        error.textContent = '有効なApps Script Web App URLを入力してください。';
        error.classList.remove('hidden');
        return;
      }
      localStorage.setItem(API_URL_KEY, apiUrl);
      sessionStorage.setItem('splitlab_api_key', apiKey);
      refreshSyncStatus();
      closeApiModal({ apiUrl, apiKey });
    });
    document.getElementById('api-cancel').addEventListener('click', () => closeApiModal(null));
    document.getElementById('api-modal').addEventListener('click', event => {
      if (event.target.id === 'api-modal' && document.getElementById('api-modal').dataset.gate !== 'true') closeApiModal(null);
    });
    document.getElementById('theme-toggle').textContent = document.documentElement.dataset.theme === 'dark' ? '☀️' : '🌙';
    document.getElementById('theme-toggle').addEventListener('click', () => {
      applyTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark');
    });
    document.getElementById('sync-status-pill').addEventListener('click', async () => {
      const connection = await requestApiConnection({ force: true });
      if (connection) syncWithSheets();
    });
    document.getElementById('delete-cancel').addEventListener('click', closeDeleteModal);
    document.getElementById('delete-modal').addEventListener('click', event => {
      if (event.target.id === 'delete-modal') closeDeleteModal();
    });
    function openProjectInfoModal() {
      const modal = document.getElementById('project-info-modal');
      modal.classList.remove('hidden');
      modal.classList.add('flex');
      setTimeout(() => document.getElementById('project-title').focus(), 0);
    }
    function closeProjectInfoModal() {
      const modal = document.getElementById('project-info-modal');
      modal.classList.add('hidden');
      modal.classList.remove('flex');
    }
    document.getElementById('edit-project-btn').addEventListener('click', openProjectInfoModal);
    document.getElementById('project-info-close').addEventListener('click', closeProjectInfoModal);
    document.getElementById('project-info-done').addEventListener('click', closeProjectInfoModal);
    document.getElementById('project-info-modal').addEventListener('click', event => {
      if (event.target.id === 'project-info-modal') closeProjectInfoModal();
    });
    function toggleModal(id, show) {
      const modal = document.getElementById(id);
      modal.classList.toggle('hidden', !show);
      modal.classList.toggle('flex', show);
    }
    document.getElementById('axis-explainer-btn').addEventListener('click', () => toggleModal('axis-explainer-modal', true));
    document.getElementById('axis-explainer-close').addEventListener('click', () => toggleModal('axis-explainer-modal', false));
    document.getElementById('axis-explainer-modal').addEventListener('click', event => {
      if (event.target.id === 'axis-explainer-modal') toggleModal('axis-explainer-modal', false);
    });
    document.getElementById('beta-detail-btn').addEventListener('click', () => {
      renderBetaDetailModal();
      toggleModal('beta-detail-modal', true);
    });
    document.getElementById('beta-detail-close').addEventListener('click', () => toggleModal('beta-detail-modal', false));
    document.getElementById('beta-detail-modal').addEventListener('click', event => {
      if (event.target.id === 'beta-detail-modal') toggleModal('beta-detail-modal', false);
    });
    document.getElementById('delete-form').addEventListener('submit', async event => {
      event.preventDefault();
      const projectId = pendingDeleteProjectId;
      const apiUrl = localStorage.getItem(API_URL_KEY) || '';
      const apiKey = document.getElementById('delete-api-key').value;
      const errorElement = document.getElementById('delete-modal-error');
      const submitButton = document.getElementById('delete-submit');
      if (!projectId || !apiUrl || !apiKey) return;
      submitButton.disabled = true;
      submitButton.textContent = '消去中…';
      errorElement.classList.add('hidden');
      try {
        const body = new URLSearchParams({ action: 'deleteProject', apiKey, payload: JSON.stringify({ projectId }) });
        const result = await postToAppsScript(apiUrl, body, { timeoutMs: 30000, timeoutLabel: '消去' });
        if (!result.ok) throw new Error(result.error || '案件を消去できませんでした。');
        projects = projects.filter(project => String(project.id) !== String(projectId));
        dirtyProjectIds.delete(projectId);
        if (result.serverTime) lastRemoteSyncAt = result.serverTime;
        persistSyncState();
        persistProjects();
        renderDashboard();
        closeDeleteModal();
        showToast(`案件を消去し、${Number(result.archivedLogCount) || 0}件のログを履歴へ保管しました`);
      } catch (error) {
        errorElement.textContent = error.message === 'Unauthorized.' ? 'API_KEYが一致しません。' : `消去エラー: ${error.message}`;
        errorElement.classList.remove('hidden');
      } finally {
        submitButton.disabled = false;
        submitButton.textContent = '消去する';
      }
    });
    document.getElementById('new-project-btn').addEventListener('click', () => {
      const project = makeProject();
      projects.push(project);
      markProjectDirty(project.id);
      persistProjects();
      scheduleAutoSync();
      openProject(project.id);
      document.getElementById('project-title').focus();
      document.getElementById('project-title').select();
    });

    document.getElementById('dashboard-view').addEventListener('click', async event => {
      const archiveButton = event.target.closest('[data-archive-project]');
      const restoreButton = event.target.closest('[data-restore-project]');
      const deleteButton = event.target.closest('[data-delete-project]');
      if (deleteButton) {
        const project = projects.find(item => item.id === deleteButton.dataset.deleteProject);
        if (project) openDeleteModal(project);
        return;
      }
      if (archiveButton) {
        const project = projects.find(item => item.id === archiveButton.dataset.archiveProject);
        if (!project || !confirm('「' + project.title + '」をアーカイブしますか？')) return;
        project.archived = true;
        project.archivedAt = new Date().toISOString();
        project.status = '完了';
        project.updatedAt = project.archivedAt;
        markProjectDirty(project.id);
        persistProjects();
        renderDashboard();
        const synced = getStoredApiConnection() ? await syncWithSheets({ silent: true }) : false;
        showToast(synced ? '案件をアーカイブし、シートへ保存しました' : '案件をアーカイブしました。接続後にシートへ同期してください');
        return;
      }
      if (restoreButton) {
        const project = projects.find(item => item.id === restoreButton.dataset.restoreProject);
        if (!project) return;
        project.archived = false;
        project.archivedAt = '';
        project.updatedAt = new Date().toISOString();
        markProjectDirty(project.id);
        persistProjects();
        renderDashboard();
        const synced = getStoredApiConnection() ? await syncWithSheets({ silent: true }) : false;
        showToast(synced ? '案件を復元し、シートへ保存しました' : '案件を一覧へ復元しました。接続後に同期してください');
        return;
      }
      const openButton = event.target.closest('[data-open-project]');
      if (openButton) openProject(openButton.dataset.openProject);
    });
    document.getElementById('toggle-archive-btn').addEventListener('click', () => {
      const content = document.getElementById('archive-content');
      const button = document.getElementById('toggle-archive-btn');
      const willOpen = content.classList.contains('hidden');
      content.classList.toggle('hidden', !willOpen);
      button.setAttribute('aria-expanded', String(willOpen));
      document.getElementById('archive-toggle-label').textContent = willOpen ? '閉じる' : '表示する';
    });
    document.getElementById('logo-home').addEventListener('click', event => {
      event.preventDefault();
      showDashboard();
    });
    document.getElementById('nav-dashboard').addEventListener('click', showDashboard);
    document.getElementById('back-dashboard').addEventListener('click', showDashboard);
    document.getElementById('sync-btn').addEventListener('click', syncWithSheets);
    document.querySelectorAll('.project-field').forEach(field => field.addEventListener('input', scheduleProjectSave));

    document.getElementById('parse-narrative-btn').addEventListener('click', () => parseNarrativeWithGoogle('log-narrative', 'parse-narrative-btn', 'parse-status'));
    document.getElementById('parse-narrative-btn-inline').addEventListener('click', () => parseNarrativeWithGoogle('log-narrative-inline', 'parse-narrative-btn-inline', 'parse-status-inline'));
    document.getElementById('open-log-form-btn').addEventListener('click', () => {
      resetLogForm();
      openLogFormModal();
      setTimeout(() => document.getElementById('log-narrative').focus(), 250);
    });
    document.getElementById('log-form-modal-close').addEventListener('click', closeLogFormModal);
    document.getElementById('log-form-modal').addEventListener('click', event => {
      if (event.target.id === 'log-form-modal') closeLogFormModal();
    });
    document.getElementById('reopen-drafts-btn').addEventListener('click', openLogDraftsModal);
    document.getElementById('log-drafts-close').addEventListener('click', closeLogDraftsModal);
    document.getElementById('log-drafts-cancel').addEventListener('click', closeLogDraftsModal);
    document.getElementById('log-drafts-modal').addEventListener('click', event => {
      if (event.target.id === 'log-drafts-modal') closeLogDraftsModal();
    });
    document.getElementById('register-drafts-btn').addEventListener('click', () => {
      const project = getActiveProject();
      if (!project) return;
      const selectedLogs = [];
      document.querySelectorAll('#log-draft-list .draft-row').forEach(row => {
        if (!row.querySelector('.draft-selected').checked) return;
        selectedLogs.push({
          id: `log-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
          person: row.querySelector('.draft-person').value,
          type: row.querySelector('.draft-type').value,
          name: row.querySelector('.draft-name').value.trim() || '制作作業',
          count: Math.max(1, Number(row.querySelector('.draft-count').value) || 1),
          scope: normalizeLevel(row.querySelector('.draft-scope').value),
          effort: normalizeLevel(row.querySelector('.draft-effort').value),
          details: row.querySelector('.draft-details').value.trim(),
          createdAt: new Date().toISOString()
        });
      });
      if (!selectedLogs.length) {
        showToast('登録するログ候補を選択してください');
        return;
      }
      project.logs.push(...selectedLogs);
      project.analysis = null;
      parsedLogDrafts = [];
      document.getElementById('log-narrative').value = '';
      document.getElementById('parse-status').textContent = '';
      document.getElementById('log-narrative-inline').value = '';
      document.getElementById('parse-status-inline').textContent = '';
      renderLogDrafts();
      closeLogDraftsModal();
      renderProductionLogs();
      scheduleProjectSave();
      showToast(`${selectedLogs.length}件の制作ログを登録しました`);
    });
    document.getElementById('log-form').addEventListener('submit', event => {
      event.preventDefault();
      const project = getActiveProject();
      if (!project) return;
      const values = {
        person: document.getElementById('log-person').value,
        type: document.getElementById('log-type').value,
        name: document.getElementById('log-name').value.trim(),
        count: Math.max(1, Number(document.getElementById('log-count').value) || 1),
        scope: normalizeLevel(document.getElementById('log-scope').value),
        effort: normalizeLevel(document.getElementById('log-effort').value),
        details: document.getElementById('log-details').value.trim()
      };
      const editingIndex = editingLogId === null ? -1 : project.logs.findIndex(log => String(log.id) === String(editingLogId));
      if (editingIndex >= 0) {
        project.logs[editingIndex] = { ...project.logs[editingIndex], ...values };
      } else {
        project.logs.push({
          id: `log-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
          ...values,
          createdAt: new Date().toISOString()
        });
      }
      const wasEditing = editingIndex >= 0;
      project.analysis = null;
      resetLogForm();
      renderProductionLogs();
      scheduleProjectSave();
      showToast(wasEditing ? '制作ログを更新しました' : '制作ログを追加しました');
    });

    document.getElementById('cancel-log-edit').addEventListener('click', () => {
      resetLogForm();
      showToast('ログ編集をキャンセルしました');
    });

    document.getElementById('log-list').addEventListener('dblclick', event => {
      if (event.target.closest('button, input, select, textarea')) return;
      const row = event.target.closest('[data-log-id]');
      if (row) beginLogEdit(row.dataset.logId);
    });

    document.getElementById('log-list').addEventListener('keydown', event => {
      if (event.key !== 'Enter' || event.target.closest('button')) return;
      const row = event.target.closest('[data-log-id]');
      if (row) beginLogEdit(row.dataset.logId);
    });

    document.getElementById('log-list').addEventListener('click', event => {
      const editButton = event.target.closest('[data-edit-log]');
      if (editButton) {
        beginLogEdit(editButton.dataset.editLog);
        return;
      }
      const removeButton = event.target.closest('[data-remove-log]');
      const project = getActiveProject();
      if (!removeButton || !project || !confirm('この制作ログを削除しますか？')) return;
      project.logs = project.logs.filter(log => String(log.id) !== String(removeButton.dataset.removeLog));
      if (String(editingLogId) === String(removeButton.dataset.removeLog)) resetLogForm();
      project.analysis = null;
      renderProductionLogs();
      scheduleProjectSave();
    });

    function openAgreementModal() {
      const project = getActiveProject();
      if (!project) return;
      const analysis = project.analysis;
      const hintText = document.getElementById('agreement-hint-text');
      const recBtn = document.getElementById('agreement-apply-rec-btn');
      if (analysis && Number.isFinite(Number(analysis.recommendedA))) {
        const recA = Number(analysis.recommendedA);
        hintText.textContent = `AI推奨値: tada ${recA.toFixed(1)}% / riku ${(100 - recA).toFixed(1)}%`;
        recBtn.classList.remove('hidden');
      } else {
        hintText.textContent = 'AI分析を実行すると推奨値が表示されます';
        recBtn.classList.add('hidden');
      }
      document.getElementById('agreement-modal').classList.remove('hidden');
      document.getElementById('agreement-modal').classList.add('flex');
    }

    function closeAgreementModal() {
      document.getElementById('agreement-modal').classList.add('hidden');
      document.getElementById('agreement-modal').classList.remove('flex');
    }

    document.getElementById('open-agreement-modal-btn')?.addEventListener('click', openAgreementModal);
    document.getElementById('agreement-modal-close')?.addEventListener('click', closeAgreementModal);
    document.getElementById('agreement-modal-done')?.addEventListener('click', () => {
      closeAgreementModal();
      showToast('スプリット合意を保存しました');
    });
    document.getElementById('agreement-modal')?.addEventListener('click', event => {
      if (event.target.id === 'agreement-modal') closeAgreementModal();
    });
    document.getElementById('agreement-apply-rec-btn')?.addEventListener('click', () => {
      const analysis = getActiveProject()?.analysis;
      if (!analysis || !Number.isFinite(Number(analysis.recommendedA))) return;
      document.getElementById('final-slider').value = Math.round(Number(analysis.recommendedA) * 2) / 2;
      calculateSplit();
      scheduleProjectSave();
      showToast('AI推奨値をスライダーに適用しました');
    });

    document.getElementById('analyze-logs-btn').addEventListener('click', analyzeLogsWithGoogle);
    document.getElementById('apply-analysis-btn').addEventListener('click', () => {
      const analysis = getActiveProject()?.analysis;
      if (!analysis) return;
      document.getElementById('final-slider').value = Math.round(Number(analysis.recommendedA) * 2) / 2;
      calculateSplit();
      scheduleProjectSave();
      openAgreementModal();
      showToast('分析推奨値を話し合いの初期値にしました');
    });
    document.getElementById('final-slider').addEventListener('input', () => {
      calculateSplit();
      scheduleProjectSave();
    });

    projects = projects.map(normalizeProject);
    persistProjects();
    document.querySelectorAll('.task-row').forEach(syncTaskRow);
    calculateSplit();

    function finishBoot() {
      renderDashboard();
      refreshSyncStatus();
      scheduleAutoSync(100);
      const initialProjectMatch = location.hash.match(/^#project=(.+)$/);
      if (initialProjectMatch) {
        suppressHistoryPush = true;
        openProject(decodeURIComponent(initialProjectMatch[1]));
        suppressHistoryPush = false;
      } else if (!location.hash) {
        history.replaceState({ view: 'dashboard' }, '', '#dashboard');
      }
    }
    window.addEventListener('popstate', () => {
      suppressHistoryPush = true;
      const match = location.hash.match(/^#project=(.+)$/);
      if (match) openProject(decodeURIComponent(match[1]));
      else showDashboard();
      suppressHistoryPush = false;
    });
    (async function bootGate() {
      if (location.protocol === 'file:' || getStoredApiConnection()) {
        finishBoot();
        return;
      }
      setSyncStatus('disconnected');
      const connection = await requestApiConnection({ gate: true });
      if (connection) finishBoot();
    })();

    setInterval(() => {
      if (document.visibilityState === 'visible') syncWithSheets({ silent: true, poll: true });
    }, 10000);
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible') scheduleAutoSync(100);
    });
