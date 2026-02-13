Option Explicit

Public Function Aggregate_Sum(dictSrc As Object, groupIndex As Long, valueIndex As Long) As Object

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim k, arr, gKey As String

    For Each k In dictSrc.keys
        arr = dictSrc(k)
        gKey = arr(groupIndex)

        If Not dict.exists(gKey) Then dict(gKey) = 0
        dict(gKey) = dict(gKey) + CDbl(arr(valueIndex))
    Next

    Set Aggregate_Sum = dict

End Function
