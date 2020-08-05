Attribute VB_Name = "Constant"
Option Explicit


'''シート検索文字列
Function SheetSearchStr() As Range
    Set SheetSearchStr = ThisWorkbook.Sheets("設定").Cells(5, 3)
End Function

'''複数セル一括入力文字列
Function MultiInputString() As Range
    Set MultiInputString = ThisWorkbook.Sheets("設定").Cells(7, 3)
End Function

'''複数セル一括入力上書き設定
Function MultiInputOverwrite() As Range
    Set MultiInputOverwrite = ThisWorkbook.Sheets("設定").Cells(8, 3)
End Function


'''置換辞書ブック
Function ReplaceDicBook() As Range
    Set ReplaceDicBook = ThisWorkbook.Sheets("設定").Cells(11, 3)
End Function

'''置換辞書シート
Function ReplaceDicSht() As Range
    Set ReplaceDicSht = ThisWorkbook.Sheets("設定").Cells(12, 3)
End Function

'''置換辞書開始行
Function ReplaceDicStartRow() As Range
    Set ReplaceDicStartRow = ThisWorkbook.Sheets("設定").Cells(13, 3)
End Function

'''置換辞書検索列
Function ReplaceDicSearchCol() As Range
    Set ReplaceDicSearchCol = ThisWorkbook.Sheets("設定").Cells(14, 3)
End Function

'''置換辞書置換列
Function ReplaceDicRepCol() As Range
    Set ReplaceDicRepCol = ThisWorkbook.Sheets("設定").Cells(15, 3)
End Function

'''グリッド範囲
Function SelectionGriding() As Range
    Set SelectionGriding = ThisWorkbook.Sheets("設定").Cells(17, 3)
End Function

'''PCL判断フラグ
Function pclMark() As Range
    Set pclMark = ThisWorkbook.Sheets("設定").Cells(18, 3)
End Function

'''正規表現でセルをクリアする設定
Function ClearByRegRepSetting() As Range
    Set ClearByRegRepSetting = ThisWorkbook.Sheets("設定").Cells(20, 3)
End Function
