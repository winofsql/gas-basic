// グローバル変数
var cur_ui;

function gas_basic_01() {

  cur_ui = SpreadsheetApp.getUi();

  var book = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(book.getUrl());

  var sheet = book.getSheets()[0];
  Logger.log(sheet.getName());

  cur_ui.alert('こんにちは 世界');

  console.log( check_prompt() );
}

function check_prompt() {

  var result = cur_ui.prompt(
      'メッセージ部分',
      'お名前を入力してください',
      cur_ui.ButtonSet.OK_CANCEL);

  var text = result.getResponseText();
  var button = result.getSelectedButton();
  if (button == cur_ui.Button.OK) {
    console.log('入力値 :' + text);
    return text;
  }
  if (button == cur_ui.Button.CANCEL) {
    console.log('CANCEL');
    return "CANCEL"
  }
  if (button == cur_ui.Button.CLOSE) {
    console.log('閉じる');
    return "CLOSE"
  }

}
