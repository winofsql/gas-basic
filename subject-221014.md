# マクロコードの確認
```javascript
function myMacro1() {

  // Excel で言うところの Book を取得
  var spreadsheet = SpreadsheetApp.getActive();

  // 現在操作中の シートを取得( Excel では worksheet )
  var sheet = spreadsheet.getActiveSheet();

  // getRange で対象範囲を指定する( 左上のクリックの処理 )
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();

  // セル内の対象範囲のデータをすべて削除
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
};
```

### セルへの書き込み
```javascript
function cellAction() {

  // データのクリア
  myMacro1();

  // 対象のスプレッドシートの ID を設定
  var id = "1Zx6ylOhCwQDGK-UBtQE_i061umoZwBZz3Gj_Rie6eJI";

  // 対象のスプレッドシート
  var spreadsheet = SpreadsheetApp.openById(id);

  // シート名
  var sheetName = "こんにちは";

  // 対象のシート
  var sheet = spreadsheet.getSheetByName(`${sheetName}`);


  for( var i = 1; i <= 10; i++ ) {
    // 範囲の指定
    var range = sheet.getRange("B" + i );

    // セルに値をセット
    range.setValue("日本語" + i );

  }
  
}
```
