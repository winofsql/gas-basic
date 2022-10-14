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
