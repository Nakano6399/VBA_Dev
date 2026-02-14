Option Explicit

Public Sub FW_Execute()
    On Error GoTo ErrHandler

    Log_Init
    Log_Write "Processing started", "INFO", "SYSTEM"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call Main_Process

    Log_Write "Processing completed successfully", "INFO", "SYSTEM"

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Call Error_Handle(Err.Number, Err.Description)
    Resume CleanExit
End Sub

