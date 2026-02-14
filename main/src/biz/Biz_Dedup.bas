Option Explicit

Public Function Dedup_ByKey(arr, keyIndex() As Long) As Object

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, k As String, j As Long

    For i = 1 To UBound(arr)
        k = ""
        For j = LBound(keyIndex) To UBound(keyIndex)
            k = k & "|" & arr(i)(keyIndex(j))
        Next

        If Not dict.exists(k) Then dict(k) = arr(i)
    Next

    Set Dedup_ByKey = dict

End Function
