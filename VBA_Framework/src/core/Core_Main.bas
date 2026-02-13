Option Explicit

Public Sub FW_Execute()
    On Error GoTo ErrHandler

    Log_Init
    Log_Write "処理開始", "INFO", "SYSTEM"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call Main_Process

    Log_Write "正常終了", "INFO", "SYSTEM"

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Call Error_Handle(Err.Number, Err.Description)
    Resume CleanExit
End Sub

