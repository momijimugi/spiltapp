/**
 * Firebase の設定値。
 *
 * LYRICLAB と同じ Firebase プロジェクト（soldb-ebb27）を使う。
 * Authentication / Firestore / セキュリティルールは共通で、
 * SPLITLAB の設定だけ users/{uid}/appSettings/splitapp に分けて保存する。
 *
 * ここに入る値は公開してよいもの（ブラウザに配られる識別子）で、秘密鍵ではない。
 * アクセス制御は Firebase Console 側の「承認済みドメイン」とセキュリティルールで行う。
 *
 * ※ SPLITLAB では、この設定が入っている限り認証を素通りさせない（fail-closed）。
 *    SDKの読み込みや初期化に失敗した場合はエラー画面を出し、本体は開かない。
 */
export const firebaseConfig = {
  apiKey: 'AIzaSyDGuejFpTXjblX8sCFMDaBxgGKde2p7VKs',
  authDomain: 'soldb-ebb27.firebaseapp.com',
  projectId: 'soldb-ebb27',
  storageBucket: 'soldb-ebb27.firebasestorage.app',
  messagingSenderId: '941289825290',
  appId: '1:941289825290:web:e19fbbecb559317f9febf3',
  // Analytics は使っていないので参照されないが、Console の値をそのまま残しておく。
  measurementId: 'G-TRCDD1WNVF'
};

/** 設定が入っているか。空のときだけ認証なしで動く（開発用の逃げ道）。 */
export function isFirebaseConfigured() {
  return Boolean(firebaseConfig.apiKey && firebaseConfig.authDomain && firebaseConfig.projectId);
}
