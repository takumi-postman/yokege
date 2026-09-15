/**
 * よけゲー 順位表 ─ Google Apps Script
 *
 * ■ 入れ替えるときの手順（URLを変えないこと）
 *  1. スプレッドシートを開く → 拡張機能 → Apps Script
 *  2. 中のコードを全部消して、このファイルの中身を貼り付ける
 *  3. 右上「デプロイ」→「デプロイを管理」
 *  4. 出ている行の右の　鉛筆マーク　を押す
 *  5. 「バージョン」を　新バージョン　に変える
 *  6. 「デプロイ」を押す
 *
 *  ※「新しいデプロイ」を選ぶと URL が変わってしまい、ゲームから
 *    順位表が見えなくなります。かならず「デプロイを管理」→ 鉛筆 で。
 *
 * ■ 記録の置き場所
 *  「scores2」シートに1行ずつ溜まります。
 *  ひとつ前の「scores」シートには、終わりが無かった頃の記録が残っています。
 *  見えないだけで消してはいないので、要らなければシートごと削除してください。
 *
 *  変な名前や、明らかに嘘くさい記録は、その行を消すだけで順位表から消えます。
 */

const SHEET_NAME = "scores2";   // ← 名前を変えると、そこから新しい順位表が始まる
const MIN_SEC  = 5;       // これ未満の記録は受け付けない
const MAX_SEC  = 86400;   // 24時間より長い記録は嘘とみなす
const MAX_NAME = 12;      // なまえの長さ上限
const MAX_LIVES = 3;      // ゲーム側の体力と同じ数

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["日時", "なまえ", "秒", "タイムコード", "残り体力", "踏破"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* なまえの掃除。改行や制御文字を落として、長さを切る。 */
function clean_(v) {
  const s = String(v == null ? "" : v)
    .replace(/[\x00-\x1f\x7f]/g, "")
    .trim()
    .slice(0, MAX_NAME);
  return s || "ななし";
}

function tc_(sec) {
  const p = function (n) { return ("0" + n).slice(-2); };
  return p(Math.floor(sec / 3600)) + ":" + p(Math.floor(sec / 60) % 60) + ":" +
         p(Math.floor(sec) % 60)   + ":" + p(Math.floor(sec * 30) % 30);
}

/**
 * どちらが上か。a が b より上なら負の数を返す。
 *
 * 時間だけで並べると、終わりまで行けた人が全員ほぼ同じ秒数になってしまい、
 * 順位が付かない。だから「踏破したか」→「体力が何個残ったか」を先に見て、
 * 時間は最後の同点決着にだけ使う。
 * 踏破していない人どうしは、今までどおり長く粘った方が上。
 */
function better_(a, b) {
  if (a.cleared !== b.cleared) return b.cleared - a.cleared;
  if (a.cleared) {
    if (a.lives !== b.lives) return b.lives - a.lives;
    return a.sec - b.sec;
  }
  return b.sec - a.sec;
}

/* ゲームから結果が届いたとき */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    let sec = Number(d.sec);
    if (!isFinite(sec) || sec < MIN_SEC || sec > MAX_SEC) {
      return ContentService.createTextOutput("skip");
    }
    sec = Math.round(sec * 100) / 100;

    let lives = Math.round(Number(d.lives));
    if (!isFinite(lives) || lives < 0) lives = 0;
    if (lives > MAX_LIVES) lives = MAX_LIVES;
    const cleared = (Number(d.cleared) === 1 && lives >= 1) ? 1 : 0;

    sheet_().appendRow([new Date(), clean_(d.name), sec, tc_(sec), lives, cleared]);
    return ContentService.createTextOutput("ok");
  } catch (err) {
    return ContentService.createTextOutput("err");
  }
}

/* 順位表を返すとき。callback が付いていれば JSONP で返す。 */
function doGet(e) {
  const top = Math.min(50, Math.max(1, Number(e.parameter.top) || 10));
  const sh = sheet_();
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 2, last - 1, 5).getValues() : [];

  // 同じなまえは、いちばん良い1回だけ残す
  const best = {};
  let plays = 0;
  rows.forEach(function (r) {
    const name = String(r[0]);
    const sec = Number(r[1]);
    if (!name || !isFinite(sec)) return;
    plays++;
    const one = {
      name: name,
      sec: sec,
      lives: isFinite(Number(r[3])) ? Number(r[3]) : 0,
      cleared: Number(r[4]) === 1 ? 1 : 0
    };
    if (!(name in best) || better_(one, best[name]) < 0) best[name] = one;
  });

  const ranked = Object.keys(best).map(function (k) { return best[k]; }).sort(better_);

  // 呼び出してきた本人の順位。ゲーム側が name= を付けてくる。
  const me = clean_(e.parameter.name);
  let you = null;
  for (let i = 0; i < ranked.length; i++) {
    if (ranked[i].name === me) {
      you = { rank: i + 1, sec: ranked[i].sec,
              lives: ranked[i].lives, cleared: ranked[i].cleared };
      break;
    }
  }

  const payload = JSON.stringify({
    list: ranked.slice(0, top),
    you: you,
    players: ranked.length,
    plays: plays,
    clears: ranked.filter(function (r) { return r.cleared; }).length
  });

  const cb = e.parameter.callback;
  if (cb && /^[A-Za-z_$][\w$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + "(" + payload + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}
