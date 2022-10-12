function gas_basic_02() {
  
  var cur_ui = SpreadsheetApp.getUi();

  var result = cur_ui.alert(
    '送信してよろしいですか?',
    cur_ui.ButtonSet.YES_NO);

  if (result == cur_ui.Button.YES) {
  }
  if (result == cur_ui.Button.NO) {
    return;
  }
  if (result == cur_ui.Button.CLOSE) {
    return;
  }

  console.log("メール送信処理開始");

  var sheet = SpreadsheetApp.getActiveSheet();
  var targetRange;
  var targetMessage = "";
  var lf2 = "\n\n";

  // 科目名
  targetRange = sheet.getRange('B2');
  targetData = targetRange.getValue().toString();
  targetMessage += "【科目名】\n" + targetData + lf2;

  // 講師名
  targetRange = sheet.getRange('C2');
  targetData = targetRange.getValue().toString();
  targetMessage += "【講師名】\n" + targetData + lf2;

  // 実施日
  targetRange = sheet.getRange('D2');
  targetData = targetRange.getDisplayValue();
  targetMessage += "【実施日】\n" + targetData + lf2;

  var targetMail = "メールアドレス";
  var targetSubject = "件名";

  GmailApp.sendEmail(targetMail, targetSubject, targetMessage );


}
