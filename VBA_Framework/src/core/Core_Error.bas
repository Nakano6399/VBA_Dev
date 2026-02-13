Option Explicit

Public Sub Error_Handle(errNo As Long, msg As String)

    Call Log_Write(msg, "ERROR", "SYSTEM")

    MsgBox "エラーが発生しました。" & vbCrLf & msg, vbCritical

End Sub


