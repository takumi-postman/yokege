/**
 * よけゲー 順位表 ─ Google Apps Script
 *
 * 使い方
 *  1. Googleスプレッドシートを新規作成する
 *  2. 拡張機能 → Apps Script を開く
 *  3. 最初から入っているコードを全部消して、このファイルの中身を貼り付ける
 *  4. 右上の「デプロイ」→「新しいデプロイ」
 *       種類                 : ウェブアプリ
 *       次のユーザーとして実行 : 自分
 *       アクセスできるユーザー : 全員     ← ここが「全員」でないと動きません
 *  5. 出てきた「ウェブアプリのURL」をコピーして渡してください
 *
 * 記録は同じスプレッドシートの「scores」シートに1行ずつ溜まります。
 * 変な名前や、明らかに嘘くさい記録は、その行を消すだけで順位表から消えます。
 */

const SHEET_NAME = "scores";
const MIN_SEC  = 5;       // これ未満の記録は受け付けない
const MAX_SEC  = 86400;   // 24時間より長い記録は嘘とみなす
const MAX_NAME = 12;      // なまえの長さ上限

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["日時", "なまえ", "秒", "タイムコード"]);
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

/* ゲームから結果が届いたとき */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    let sec = Number(d.sec);
    if (!isFinite(sec) || sec < MIN_SEC || sec > MAX_SEC) {
      return ContentService.createTextOutput("skip");
    }
    sec = Math.round(sec * 100) / 100;
    sheet_().appendRow([new Date(), clean_(d.name), sec, tc_(sec)]);
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
  const rows = last > 1 ? sh.getRange(2, 2, last - 1, 2).getValues() : [];

  // 同じなまえは、一番いい記録だけ残す
  const best = {};
  let plays = 0;
  rows.forEach(function (r) {
    const name = String(r[0]);
    const sec = Number(r[1]);
    if (!name || !isFinite(sec)) return;
    plays++;
    if (!(name in best) || sec > best[name]) best[name] = sec;
  });

  const ranked = Object.keys(best)
    .map(function (name) { return { name: name, sec: best[name] }; })
    .sort(function (a, b) { return b.sec - a.sec; });

  // 呼び出してきた本人の順位。ゲーム側が name= を付けてくる。
  const me = clean_(e.parameter.name);
  let you = null;
  for (let i = 0; i < ranked.length; i++) {
    if (ranked[i].name === me) { you = { rank: i + 1, sec: ranked[i].sec }; break; }
  }

  const payload = JSON.stringify({
    list: ranked.slice(0, top),
    you: you,
    players: ranked.length,
    plays: plays
  });

  const cb = e.parameter.callback;
  if (cb && /^[A-Za-z_$][\w$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + "(" + payload + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}
