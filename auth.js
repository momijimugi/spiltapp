/* ==========================================================================
   SPLITLAB — Googleログイン（Firebase Authentication）＋ 接続設定のクラウド保存

   既存のアプリ（app.js）の外側に認証の層をかぶせるモジュール。
   app.js の案件データ・同期処理には手を触れず、このファイルは
     ・本体を見せるかどうか
     ・ヘッダーにアカウントを出すか
     ・Firestore（users/{uid}/appSettings/splitapp）の読み書き
   だけを受け持つ。ワークスペースの実体は app.js 側の既存実装のまま。

   状態は <html data-auth>。CSSはこの属性だけを見る。
     checking  … 認証状態の確認中。本体もログイン画面も出さない
     out       … 未ログイン。ログイン画面だけを出す
     restoring … 保存された接続設定を復元中。まだ本体を出さない
     in        … ログイン済み。本体を出す
     error     … 認証を確立できなかった。本体は開かない（fail-closed）

   LYRICLAB の初期実装と違い、firebase-config.js に設定が入っている状態で
   SDK読み込み／Firebase初期化／Auth初期化に失敗した場合は本体を開かない。
   Firestore だけが落ちた場合は、ログイン済みユーザーに限り
   既存localStorageの接続設定での手動接続へフォールバックする。
   ========================================================================== */

import { firebaseConfig, isFirebaseConfigured } from './firebase-config.js?v=12';

// Firebase公式のブラウザ向けES Modules。ビルド環境（npm/Vite）は使わない。
const SDK = 'https://www.gstatic.com/firebasejs/10.14.1';

/**
 * Firestore に置く接続設定の場所。uid ごとに1件。
 *   users/{uid}/appSettings/splitapp
 * LYRICLAB（.../lyricsplit）とはドキュメントを分けているので互いに干渉しない。
 * セキュリティルールで request.auth.uid == uid のときだけ読み書きできる前提。
 * Gemini等の外部APIキーはここに入れない（Apps Script の Script Properties 側で持つ）。
 */
const SETTINGS_COLLECTION = 'appSettings';
const SETTINGS_DOC = 'splitapp';

const $ = (id) => document.getElementById(id);

/**
 * 復元中の進捗表示。何を待っているのかが分かるように、段階と割合を出す。
 * 進むだけで戻らない（percentは単調増加）ようにしておく。
 */
let restorePercent = 0;
function setRestoreProgress(percent, phase) {
  restorePercent = Math.max(restorePercent, Math.min(100, Math.round(percent)));
  const bar = $('auth-restore-bar');
  const label = $('auth-restore-percent');
  const text = $('auth-restore-phase');
  if (bar) bar.style.width = `${restorePercent}%`;
  if (label) label.textContent = `${restorePercent}%`;
  if (phase && text) text.textContent = phase;
}

const setAuthState = (state) => {
  document.documentElement.dataset.auth = state;
  // 起動が止まっていないことをウォッチドッグ（index.html）へ知らせる。
  window.__SPLITLAB_AUTH_BOOTED = true;
};

const listeners = [];
window.SPLITLAB_AUTH = {
  enabled: false,
  user: null,
  get uid() { return this.user ? this.user.uid : null; },
  get displayName() { return this.user ? this.user.displayName : null; },
  get email() { return this.user ? this.user.email : null; },
  /**
   * 接続設定の保存先。app.js から呼ぶ。
   * 未ログイン・Firestore未使用のときは available=false を返すだけで、
   * 呼び出し側は今まで通り localStorage / sessionStorage だけで動く。
   */
  settings: unavailableSettings(),
  /** 復元中の進捗表示。app.js の復元処理から段階を進める。 */
  progress: setRestoreProgress,
  onChange(fn) {
    if (typeof fn !== 'function') return () => {};
    listeners.push(fn);
    try { fn(this.user); } catch (e) { console.error(e); }
    return () => {
      const i = listeners.indexOf(fn);
      if (i >= 0) listeners.splice(i, 1);
    };
  }
};

function publish(user) {
  window.SPLITLAB_AUTH.user = user;
  listeners.forEach((fn) => { try { fn(user); } catch (e) { console.error(e); } });
}

function showError(message) {
  const el = $('auth-error');
  if (!el) return;
  el.textContent = message || '';
  el.classList.toggle('hidden', !message);
}

/** 復旧できない失敗。本体は開かず、原因を出したまま止める。 */
function failClosed(message, detail) {
  const el = $('auth-fatal-detail');
  if (el) el.textContent = [message, detail].filter(Boolean).join('\n');
  console.error('[SPLITLAB] 認証を確立できませんでした:', message, detail || '');
  setAuthState('error');
}

/** ヘッダーのアカウント表示。名前とアイコン、ログアウトボタン。 */
function renderAccount(user) {
  const wrap = $('auth-user');
  if (!wrap) return;
  wrap.classList.toggle('hidden', !user);
  if (!user) return;

  const name = user.displayName || user.email || 'ログイン中';
  $('auth-user-name').textContent = name;
  wrap.title = [user.displayName, user.email].filter(Boolean).join(' · ');

  const img = $('auth-user-photo');
  const initial = $('auth-user-initial');
  if (user.photoURL) {
    img.src = user.photoURL;
    img.alt = name;
    img.classList.remove('hidden');
    initial.classList.add('hidden');
  } else {
    // 画像が無いアカウントもあるので、頭文字で代用する。
    img.classList.add('hidden');
    initial.textContent = name.slice(0, 1);
    initial.classList.remove('hidden');
  }
}

/** firebase-config.js が空のときだけ通る、認証なしの従来動作。 */
function runWithoutAuth(reason) {
  console.warn('[SPLITLAB] Googleログインは無効です:', reason);
  window.SPLITLAB_AUTH.enabled = false;
  renderAccount(null);
  publish(null);
  setAuthState('in');
  startApp();
}

/** app.js の起動（接続ゲート込み）。二重に呼ばれても一度しか走らない。 */
let appStarted = false;
function startApp() {
  if (appStarted) return;
  const app = window.SPLITLAB_APP;
  if (!app || typeof app.start !== 'function') return;
  appStarted = true;
  try { app.start(); } catch (e) { console.error('[SPLITLAB] 起動に失敗', e); }
}

async function boot() {
  if (!isFirebaseConfigured()) {
    runWithoutAuth('firebase-config.js が未設定です。');
    return;
  }

  let app, auth, provider, signInWithPopup, signOut;
  try {
    const [{ initializeApp }, authMod] = await Promise.all([
      import(`${SDK}/firebase-app.js`),
      import(`${SDK}/firebase-auth.js`)
    ]);
    app = initializeApp(firebaseConfig);
    auth = authMod.getAuth(app);
    provider = new authMod.GoogleAuthProvider();
    signInWithPopup = authMod.signInWithPopup;
    signOut = authMod.signOut;

    // ブラウザを閉じてもログイン状態を保つ。
    await authMod.setPersistence(auth, authMod.browserLocalPersistence);

    window.SPLITLAB_AUTH.enabled = true;
    // 読み込みだけ先に始めておく。ログイン確定後にまとめて待たないで済む。
    loadFirestore(app).catch(() => {});

    authMod.onAuthStateChanged(auth, (user) => {
      showError('');
      renderAccount(user);
      publish(user);
      if (!user) {
        detachSettings();
        if (window.SPLITLAB_APP && window.SPLITLAB_APP.reset) window.SPLITLAB_APP.reset();
        setAuthState('out');
        return;
      }
      attachSettings(app, user.uid);
      // 保存された接続設定があれば、本体を出す前に復元と接続確認を済ませる。
      // 接続画面が一瞬出てから消える、という見え方にしないため。
      restorePercent = 0;
      setRestoreProgress(12, 'ログインを確認しました');
      setAuthState('restoring');
      restoreConnection()
        .catch((e) => console.error('[SPLITLAB] 接続設定の復元に失敗', e))
        .finally(() => {
          setRestoreProgress(100, '完了しました');
          setAuthState('in');
          startApp();
        });
    }, (err) => {
      // 認証状態そのものが確認できない。本体は開かない。
      failClosed('ログイン状態を確認できませんでした。', (err && err.message) || String(err));
    });
  } catch (e) {
    // SDKが読めない・Firebase初期化に失敗、といった段階の失敗。
    // 設定が入っている本番環境なので、素通りさせずここで止める。
    failClosed(
      'Firebaseに接続できませんでした。通信環境を確認して再読み込みしてください。',
      (e && e.message) || String(e)
    );
    return;
  }

  const btn = $('auth-signin');
  if (btn) {
    btn.addEventListener('click', async () => {
      showError('');
      btn.disabled = true;
      try {
        await signInWithPopup(auth, provider);
        // 画面の切り替えは onAuthStateChanged 側で行う。
      } catch (e) {
        showError(signInErrorMessage(e));
      } finally {
        btn.disabled = false;
      }
    });
  }

  const out = $('auth-signout');
  if (out) {
    out.addEventListener('click', async () => {
      if (window.SPLITLAB_APP && window.SPLITLAB_APP.confirmSignOut && !window.SPLITLAB_APP.confirmSignOut()) return;
      try {
        await signOut(auth);
        // ログアウト後は状態を持ち越さないよう読み込み直す。
        location.reload();
      } catch (e) {
        showError((e && e.message) || String(e));
      }
    });
  }

  const retry = $('auth-fatal-retry');
  if (retry) retry.addEventListener('click', () => location.reload());
}

/* ------------------------------------------------------- 接続設定の保存 */

let firestore = null;
let settingsDocRef = null;
let firestoreMod = null;

function unavailableSettings() {
  const no = async () => ({ ok: false, reason: 'unavailable' });
  return {
    available: false,
    load: no,
    upsertWorkspace: no,
    renameWorkspace: no,
    removeWorkspace: no,
    setActive: no,
    mergeWorkspaces: no,
    remove: no
  };
}

function detachSettings() {
  firestore = null;
  settingsDocRef = null;
  window.SPLITLAB_AUTH.settings = unavailableSettings();
}

/** 文字列以外が入っていても落ちないように、期待する型だけを通す。 */
function normalizeWorkspaceMap(raw) {
  const out = {};
  const src = raw && typeof raw === 'object' ? raw : {};
  Object.keys(src).forEach((id) => {
    const w = src[id];
    if (!w || typeof w !== 'object') return;
    out[id] = {
      label: typeof w.label === 'string' ? w.label : '',
      apiUrl: typeof w.apiUrl === 'string' ? w.apiUrl : '',
      apiKey: typeof w.apiKey === 'string' ? w.apiKey : '',
      // 既定ワークスペースかどうか。localStorageのキー割り当てに使う。
      primary: w.primary === true,
      updatedAt: typeof w.updatedAt === 'string' ? w.updatedAt : ''
    };
  });
  return out;
}

/**
 * ログインしたuid専用の設定ドキュメントを用意する。
 * 読み書きするのは users/{uid}/appSettings/splitapp の1件だけで、
 * 他のユーザーの領域には触れない。
 */
function attachSettings(app, uid) {
  window.SPLITLAB_AUTH.settings = {
    available: true,

    async load() {
      try {
        const m = await loadFirestore(app);
        const snap = await m.getDoc(ref(m, uid));
        if (!snap.exists()) return { ok: true, data: null };
        const d = snap.data() || {};
        return {
          ok: true,
          data: {
            activeWorkspaceId: typeof d.activeWorkspaceId === 'string' ? d.activeWorkspaceId : null,
            workspaces: normalizeWorkspaceMap(d.workspaces)
          }
        };
      } catch (e) {
        console.error('[SPLITLAB] 接続設定の読み込みに失敗', e);
        return { ok: false, reason: 'firestore', error: e };
      }
    },

    /**
     * 接続先を1件だけ追加・更新する。
     * workspaces マップ全体は書き換えず、workspaces.{id} だけを更新するので、
     * 別端末がその間に追加した接続先が巻き戻ることはない。
     * 同じ接続先を2端末から触った場合は updatedAt が新しい方を残す。
     */
    async upsertWorkspace(id, entry) {
      return mutate(app, uid, (current, m) => {
        const value = normalizeWorkspaceMap({ [id]: entry })[id];
        const existing = current && current.workspaces[id];
        // 相手側のほうが新しければ何もしない（自分の古い値で上書きしない）。
        if (existing && existing.updatedAt && value.updatedAt && existing.updatedAt > value.updatedAt) {
          return null;
        }
        return { fields: { [`workspaces.${id}`]: value } };
      });
    },

    /** 表示名だけを差し替える。IDと接続情報、他の接続先には触れない。 */
    async renameWorkspace(id, label, updatedAt) {
      return mutate(app, uid, (current) => {
        const existing = current && current.workspaces[id];
        // Firestore側に無い接続先は、名前だけ書いても意味が無いので触らない。
        if (!existing) return null;
        if (existing.updatedAt && updatedAt && existing.updatedAt > updatedAt) return null;
        return {
          fields: {
            [`workspaces.${id}.label`]: String(label || ''),
            [`workspaces.${id}.updatedAt`]: String(updatedAt || new Date().toISOString())
          }
        };
      });
    },

    /** 接続先を1件だけ消す。他の接続先は残す。 */
    async removeWorkspace(id, nextActiveId) {
      return mutate(app, uid, (current, m) => {
        if (!current || !current.workspaces[id]) return null;
        const fields = { [`workspaces.${id}`]: m.deleteField() };
        // 消したものがactiveだったときだけ、activeも同時に付け替える。
        if (current.activeWorkspaceId === id) {
          fields.activeWorkspaceId = (nextActiveId && current.workspaces[nextActiveId]) ? nextActiveId : null;
        }
        return { fields };
      });
    },

    /** activeWorkspaceId だけを変える。 */
    async setActive(id) {
      return mutate(app, uid, (current) => {
        if (current && current.activeWorkspaceId === id) return null;
        return { fields: { activeWorkspaceId: id || null } };
      });
    },

    /**
     * 初回移行など、複数件をまとめて入れる場面。
     * Firestore側の現在値をトランザクション内で読み、既にある接続先は残したまま
     * 足りないものだけを足す。既存のほうが新しければそちらを優先する。
     */
    async mergeWorkspaces(entries, activeId) {
      return mutate(app, uid, (current) => {
        const incoming = normalizeWorkspaceMap(entries);
        const fields = {};
        Object.keys(incoming).forEach((id) => {
          const existing = current && current.workspaces[id];
          if (existing && existing.updatedAt && incoming[id].updatedAt
            && existing.updatedAt > incoming[id].updatedAt) return;
          fields[`workspaces.${id}`] = incoming[id];
        });
        // activeは、Firestore側にまだ無いときだけこちらの値を入れる。
        if (activeId && !(current && current.activeWorkspaceId)) fields.activeWorkspaceId = activeId;
        if (!Object.keys(fields).length) return null;
        return { fields };
      });
    },

    async remove() {
      try {
        const m = await loadFirestore(app);
        await m.deleteDoc(ref(m, uid));
        return { ok: true };
      } catch (e) {
        console.error('[SPLITLAB] 接続設定の削除に失敗', e);
        return { ok: false, reason: 'firestore', error: e };
      }
    }
  };
}

function ref(m, uid) {
  if (!settingsDocRef) {
    settingsDocRef = m.doc(firestore, 'users', uid, SETTINGS_COLLECTION, SETTINGS_DOC);
  }
  return settingsDocRef;
}

async function loadFirestore(app) {
  if (!firestoreMod) firestoreMod = await import(`${SDK}/firebase-firestore.js`);
  if (!firestore) firestore = firestoreMod.getFirestore(app);
  return firestoreMod;
}

/**
 * 接続設定の部分更新。
 *
 * workspaces マップ全体を書き戻すと、別端末がその間に追加・改名した接続先を
 * 古い手元の状態で巻き戻してしまう。そこでトランザクションの中で
 *   1. Firestoreの現在値を読む
 *   2. 更新する「フィールドだけ」を決める（workspaces.{id} などのフィールドパス）
 *   3. tx.update で、そのフィールドだけを書く
 * という手順にしている。触っていない接続先は読み書きの対象にならないので消えない。
 *
 * plan() が null を返した場合は「書く必要が無い」とみなして何もしない。
 * ドキュメントがまだ無いときだけ、tx.set で新規作成する。
 */
async function mutate(app, uid, plan) {
  try {
    const m = await loadFirestore(app);
    const docRef = ref(m, uid);
    const result = await m.runTransaction(firestore, async (tx) => {
      const snap = await tx.get(docRef);
      const current = snap.exists()
        ? {
            activeWorkspaceId: typeof (snap.data() || {}).activeWorkspaceId === 'string'
              ? snap.data().activeWorkspaceId : null,
            workspaces: normalizeWorkspaceMap((snap.data() || {}).workspaces)
          }
        : null;

      const step = plan(current, m);
      if (!step || !step.fields || !Object.keys(step.fields).length) return 'skipped';

      if (!current) {
        // 初回作成。フィールドパスは使えないので、ここだけ組み立てて書く。
        const created = { activeWorkspaceId: null, workspaces: {}, updatedAt: new Date().toISOString() };
        Object.keys(step.fields).forEach((path) => {
          const value = step.fields[path];
          if (path === 'activeWorkspaceId') { created.activeWorkspaceId = value; return; }
          const id = path.split('.')[1];
          if (!id) return;
          // 新規作成時に「消す」指示が来ることは無いので無視してよい。
          if (path === `workspaces.${id}`) created.workspaces[id] = value;
        });
        tx.set(docRef, created);
        return 'created';
      }

      tx.update(docRef, { ...step.fields, updatedAt: new Date().toISOString() });
      return 'updated';
    });
    return { ok: true, result };
  } catch (e) {
    console.error('[SPLITLAB] 接続設定の更新に失敗', e);
    return { ok: false, reason: 'firestore', error: e };
  }
}

/**
 * ログイン直後の復元。実際のワークスペース操作・接続確認は app.js 側が行う。
 * ここは「Firestoreを読んで渡す」だけ。Firestoreが落ちていても
 * 既存localStorageの接続設定での手動接続へ落ちられるようにする。
 */
async function restoreConnection() {
  const app = window.SPLITLAB_APP;
  if (!app || typeof app.restoreConnection !== 'function') return;

  setRestoreProgress(28, '保存された接続先を読み込んでいます');
  const res = await window.SPLITLAB_AUTH.settings.load();
  setRestoreProgress(55, '接続先を確認しています');
  if (!res.ok) {
    // Firestoreが読めない。Apps Scriptの失敗とは区別して伝える。
    app.restoreFailed('firestore', (res.error && res.error.message) || '');
    return;
  }
  await app.restoreConnection(res.data);
}

/** よくある失敗は、原因が分かる日本語にして画面に出す。 */
function signInErrorMessage(e) {
  const code = (e && e.code) || '';
  if (code === 'auth/popup-blocked') {
    return 'ブラウザにポップアップを塞がれました。このサイトのポップアップを許可してから、もう一度お試しください。';
  }
  if (code === 'auth/popup-closed-by-user' || code === 'auth/cancelled-popup-request') {
    return '';
  }
  if (code === 'auth/unauthorized-domain') {
    return 'このドメインはFirebaseで許可されていません。Firebase Console の Authentication > Settings > 承認済みドメイン に、このサイトのドメインを追加してください。';
  }
  if (code === 'auth/operation-not-allowed') {
    return 'Googleログインが有効になっていません。Firebase Console の Authentication > Sign-in method で Google を有効にしてください。';
  }
  if (code === 'auth/network-request-failed') {
    return 'ネットワークに接続できませんでした。通信状況を確認してください。';
  }
  return 'ログインできませんでした。' + ((e && e.message) || '');
}

boot();
