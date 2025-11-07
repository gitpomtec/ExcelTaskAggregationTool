Option Explicit

'==================================================================
' dataシートA4以降の「hhmm-hhmm タスク名」を集計し
' resultシートを作成（既存あれば削除）→最後尾に出力
'==================================================================

Private Function ParseLine(ByVal s As String, _
                           ByRef t0 As Date, ByRef t1 As Date, ByRef task As String) As Boolean
    Dim h1 As Long, m1 As Long, h2 As Long, m2 As Long

    s = Trim$(s)
    If Len(s) < 11 Then Exit Function                 ' "hhmm-hhmm タスク"
    If Mid$(s, 5, 1) <> "-" Then Exit Function        ' 5文字目がハイフン
    If Mid$(s, 10, 1) <> " " Then Exit Function       ' 10文字目がスペース
    If Not IsNumeric(Left$(s, 4)) Then Exit Function
    If Not IsNumeric(Mid$(s, 6, 4)) Then Exit Function

    h1 = CLng(Left$(s, 2))
    m1 = CLng(Mid$(s, 3, 2))
    h2 = CLng(Mid$(s, 6, 2))
    m2 = CLng(Mid$(s, 8, 2))

    If h1 > 23 Or h2 > 23 Or m1 > 59 Or m2 > 59 Then Exit Function

    t0 = TimeSerial(h1, m1, 0)
    t1 = TimeSerial(h2, m2, 0)
    task = Trim$(Mid$(s, 11))
    ParseLine = (Len(task) > 0)
End Function


'==================================================================
' メイン処理
'==================================================================
Public Sub RunAggregation()
    Dim wsData As Worksheet, wsResult As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long
    Dim dict As Object, key As Variant
    Dim s As String, task As String
    Dim t0 As Date, t1 As Date
    Dim dmin As Long

    Set wsData = ThisWorkbook.Worksheets("data")

    ' 既存result削除
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("result").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' resultを最後尾に作成
    Set wsResult = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResult.Name = "result"

    Set dict = CreateObject("Scripting.Dictionary")

    ' dataシートA4以降を処理
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    For r = 4 To lastRow
        s = CStr(wsData.Cells(r, "A").Value2)
        If ParseLine(s, t0, t1, task) Then
            dmin = DateDiff("n", t0, t1)
            If dmin < 0 Then dmin = dmin + 1440
            If dict.Exists(task) Then
                dict(task) = dict(task) + dmin
            Else
                dict.Add task, dmin
            End If
        End If
    Next r

    ' 結果出力
    wsResult.Range("A1:C1").Value = Array("Task", "Total Time(min)", "Total Time (hh:mm)")
    outRow = 2
    For Each key In dict.Keys
        wsResult.Cells(outRow, 1).Value = key
        wsResult.Cells(outRow, 2).Value = dict(key)
        wsResult.Cells(outRow, 3).Value = dict(key) / 1440#
        outRow = outRow + 1
    Next key

    ' 書式とソート
    If outRow > 2 Then
        wsResult.Range("C2:C" & outRow - 1).NumberFormatLocal = "[h]:mm"
        wsResult.Range("A1:C" & outRow - 1).Sort _
            Key1:=wsResult.Range("B2:B" & outRow - 1), _
            Order1:=2, Header:=xlYes
    End If

    wsResult.Columns("A:C").AutoFit
End Sub


'==================================================================
' ボタン1クリックイベント（ボタンに割当済みの場合）
'==================================================================
Sub ボタン1_Click()
    Application.ScreenUpdating = False
    RunAggregation
    Application.ScreenUpdating = True
End Sub

Sub ボタン2_Click()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("data")
    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 4 Then
        ws.Range("A4:A" & lastRow).ClearContents   ' A4以下の内容を消去（書式は保持）
    End If

    Application.ScreenUpdating = True
End Sub
