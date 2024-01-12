Attribute VB_Name = "Module1"
Option Explicit
Sub 勘定科目ごとに別シート抽出()

Dim Kanjoukamoku_List As Worksheet
Dim Data As Worksheet
Dim Genshi As Worksheet
Dim i As Long
Dim J As Long
Dim RowCnt  As Long
Dim LastRow  As Long
Dim List_Cnt As Long
Dim ShtName As String

    '各シートを変数代入
    Set Kanjoukamoku_List = Sheets("勘定科目リスト")
    Set Data = Sheets("小口現金出納帳")
    Set Genshi = Sheets("原紙")

    '「勘定科目」 最終行
    List_Cnt = Kanjoukamoku_List.Cells(Rows.Count, 1).End(xlUp).Row

    '「データ」 最終行
    LastRow = Data.Cells(Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False

        '「勘定科目リスト」をもとに新規シート作成
        For i = 2 To List_Cnt
            Genshi.Copy after:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = Kanjoukamoku_List.Cells(i, 1)
        Next i

        '得意先の数をループ
        For i = 2 To List_Cnt

            Data.Select

            '貼り付け開始行
            RowCnt = 3

            'シート名取得
            ShtName = Kanjoukamoku_List.Cells(i, 1)

            '「データ」 2～最終行までループ
            For J = 6 To LastRow

                '「データ」 に 「勘定科目リスト」 と同じ名称があったら
                If Kanjoukamoku_List.Cells(i, 1) = Data.Cells(J, 3) Then

                    '「データ」  A～D列の値を、該当シートに貼り付け
                    Data.Range(Cells(J, 1), Cells(J, 7)).Copy Sheets(ShtName).Cells(RowCnt, 1)

                    '貼り付け開始行を更新
                    RowCnt = RowCnt + 1
               
                End If
            Next J
        Next i

    Application.ScreenUpdating = True
    
    MsgBox "勘定科目ごとの表が作成されました。各シートの内容をチェックし、加筆修正をおねがいします。"

End Sub


