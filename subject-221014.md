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
![image](https://user-images.githubusercontent.com/1501327/195759891-bbea6f75-a5a7-405e-b7df-2f074fc91cd3.png)\
![image](https://user-images.githubusercontent.com/1501327/195759967-a21ef019-a198-4d86-93d9-dbfa74118f1f.png)

### スプレッドシート同士のデータ転送
```javascript
function myFunction() {

  clearAll();

  var spreadsheet1 = SpreadsheetApp.getActive();
  var sheet1 = spreadsheet1.getActiveSheet();

  var spreadsheet2 = SpreadsheetApp.openById("1Zx6ylOhCwQDGK-UBtQE_i061umoZwBZz3Gj_Rie6eJI");
  var sheet2 = spreadsheet2.getSheetByName("社員マスタ");

  var i = 1;
  while( true ) {

    var range1 = sheet1.getRange("A" + i);
    var data = range1.getDisplayValue();
    if ( data == "" ) {
      break;
    }

    console.log(data);

    var range2 = sheet2.getRange("A" + i);
    range2.setValue(data);

    range2 = sheet2.getRange("B" + i);
    range2.setValue( sheet1.getRange("B" + i).getDisplayValue() );

    range2 = sheet2.getRange("C" + i);
    range2.setValue( sheet1.getRange("C" + i).getDisplayValue() );

    range2 = sheet2.getRange("D" + i);
    range2.setValue( sheet1.getRange("D" + i).getDisplayValue() );

    range2 = sheet2.getRange("E" + i);
    range2.setValue( sheet1.getRange("E" + i).getDisplayValue() );

    range2 = sheet2.getRange("F" + i);
    range2.setValue( sheet1.getRange("F" + i).getDisplayValue() );

    range2 = sheet2.getRange("G" + i);
    range2.setValue( sheet1.getRange("G" + i).getDisplayValue() );

    range2 = sheet2.getRange("H" + i);
    range2.setValue( sheet1.getRange("H" + i).getDisplayValue() );

    range2 = sheet2.getRange("I" + i);
    range2.setValue( sheet1.getRange("I" + i).getDisplayValue() );

    range2 = sheet2.getRange("J" + i);
    range2.setValue( sheet1.getRange("J" + i).getDisplayValue() );

    range2 = sheet2.getRange("K" + i);
    range2.setValue( sheet1.getRange("K" + i).getDisplayValue() );

    i++;
  }
  
}

function clearAll() {

  // Excel で言うところの Book を取得
  var spreadsheet = SpreadsheetApp.openById("1Zx6ylOhCwQDGK-UBtQE_i061umoZwBZz3Gj_Rie6eJI");

  // 現在操作中の シートを取得( Excel では worksheet )
  var sheet = spreadsheet.getSheetByName("社員マスタ");

  // getRange で対象範囲を指定する( 左上のクリックの処理 )
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();

  // セル内の対象範囲のデータをすべて削除
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
}
```
