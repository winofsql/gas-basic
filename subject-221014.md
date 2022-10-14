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

### セルへの書き込みと他のセルからのデータの取得 )
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

  for( var i = 1; i <= 10; i++ ) {
    // 範囲の指定
    var range = sheet.getRange("B" + i );
    var range2 = sheet.getRange("D" + i );
    var range3 = sheet.getRange("F" + i );

    // セルに値をセット
    range2.setValue( range.getValue().toString() );
    range3.setValue( range.getDisplayValue() );

  }
 
}
```

### Excel VBA 処理
```vb
Sub Test1()

    Dim range As range
    
    Set range = Sheet1.range("A1")
    
    range.Value = ThisWorkbook.Name
    
    Set range = Sheet1.range("A2")

    range.Value = Sheet1.Name

    Set range = Sheet1.range("A3")

    range.Value = "あいうえお"
    
    For I = 1 To 10
    
        Set range = Sheet1.range("B" & I)
        range.Value = "日本語" & I
    
    Next

End Sub
```
