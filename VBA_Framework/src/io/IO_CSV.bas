Option Explicit

Public Function CSV_LoadFast(path As String) As Variant

    Dim f As Integer, buf As String
    Dim data(), tmp
    Dim r As Long

    f = FreeFile
    Open path For Input As #f

    Do Until EOF(f)
        Line Input #f, buf
        tmp = Split(buf, ",")
        ReDim Preserve data(r)
        data(r) = tmp
        r = r + 1
    Loop

    Close #f
    CSV_LoadFast = data

End Function
