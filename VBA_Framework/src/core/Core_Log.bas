Option Explicit

Private Const LOG_SHEET As String = "LOG"

Public Sub Log_Init()
    If Not SheetExists(LOG_SHEET) Then
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = LOG_SHEET
        ActiveSheet.Range("A1:D1").Value = Array("日時", "レベル", "処理", "内容")
    End If
End Sub

Public Sub Log_Write(msg As String, Optional level As String = "INFO", Optional proc As String = "")
    Dim ws As Worksheet: Set ws = Sheets(LOG_SHEET)
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = level
    ws.Cells(r, 3).Value = proc
    ws.Cells(r, 4).Value = msg
End Sub
