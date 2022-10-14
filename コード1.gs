// グローバル変数
var cur_ui;

function gas_basic_01() {

  // 現在のスプレッドシートの Ui インスタンスを取得
  cur_ui = SpreadsheetApp.getUi();

  // 現在のスプレッドシートインスタンスを取得
  var book = SpreadsheetApp.getActiveSpreadsheet();
  // 現在のスプレッドシートの URL を取得
  Logger.log(book.getUrl());
  // ID を保存しておくと、URL からすぐ目的のスプレッドシートが開く
  Logger.log(book.getId() );

  // 現在のスプレットシートが持つシート一覧の先頭のインスタンス( Sheet )を取得
  var sheet = book.getSheets()[0];
  // シート名の表示
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
