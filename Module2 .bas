Attribute VB_Name = "Module2"
Option Explicit
Sub 勘定科目ごとの表の金額を集計し総計表に転記する()
    Dim X As Long
    Dim Y As Long
    Dim Z As Range
    
    For X = 5 To Sheets.Count '勘定科目ごとのシートの数だけ次の処理をする
        Sheets(X).Activate
        Sheets(X).Cells(2, 11) = WorksheetFunction.Sum(Sheets(X).Range(Range("G3"), Range("G3").End(xlDown)))          'Sセルの合計を同じシートの「総計」欄に表示
    Next X
    
    For Y = 5 To Sheets.Count '勘定科目ごとのシートの数だけ次の処理をする
        Sheets(Y).Activate
        Sheets(4).Cells(Y, 2) = Sheets(Y).Name         '未処理の勘定科目シートのうち、左端のものの名前を「総計」シートの名前欄にコピー
        Sheets(4).Cells(Y, 3) = Sheets(Y).Range("K2")  '未処理の勘定科目シートのうち、左端のものの「総計」セルを「総計」シートの名前欄にコピー
    Next Y
    
    Sheets(4).Activate
        Set Z = Sheets(4).Range(Range("C5"), Range("C5").End(xlDown)) '変数Zに転記された科目ごとの合計セルの列を代入
        Sheets(4).Cells(18, 6) = Application.WorksheetFunction.Sum(Z)  'Zセルの総計を同じシートの「総計」欄に表示
    
    MsgBox "総計の入力が完了しました。"
End Sub

