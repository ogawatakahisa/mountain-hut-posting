// main.js
function getMailAndWrite(label) {
  // 既存の関数内容はそのまま
  // ただし、定数やユーティリティ関数は別ファイルから参照
  // 例: getGmail() -> gmailUtils.getGmail()
  // 例: addLabelToEmail() -> gmailUtils.addLabelToEmail()
  // 例: getSpreadsheetByName() -> spreadsheetUtils.getSpreadsheetByName()
  // 例: getDateFromSubject() -> messageExtractors.getDateFromSubject()
  // など
  // これらの関数呼び出し部分はそのままにしておきます
}
function onButtonClickUsuall() {
  getMailAndWrite('usuall');
}
function onButtonClickLimit() {
  getMailAndWrite('予約上限');
}
function onButtonClickNotyet() {
  getMailAndWrite('未作成のブック');
}
function onButtonClickError() {
  getMailAndWrite('想定外エラー');
}
function onButtonClickNotyetSheet() {
  getMailAndWrite('未作成のシート');
}
function onButtonClicknotName() {
  getMailAndWrite('not氏名');
}
// 他のボタンも同様
