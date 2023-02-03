Sub CountUniqueValues()

Dim ws As Worksheet
Dim rng As Range
Dim dict As Object
Dim newSheet As Worksheet
Dim lastRow As Long

Set dict = CreateObject("Scripting.Dictionary")

For Each ws In ThisWorkbook.Worksheets
    If ws.Name Like "RIO 0*" Then
        Set rng = ws.Range("H1:H" & ws.Cells(ws.Rows.Count, "H").End(xlUp).Row)
        For Each cell In rng
            If Not dict.Exists(cell.Value) Then
                dict.Add cell.Value, 1
            Else
                dict(cell.Value) = dict(cell.Value) + 1
            End If
        Next cell
    End If
Next ws

Set newSheet = ThisWorkbook.Worksheets.Add
newSheet.Name = "SUMM"
newSheet.Range("A1").Value = "Source Worksheet"
newSheet.Range("B1").Value = "Unique Value"
newSheet.Range("C1").Value = "Count"

lastRow = 2
For Each Key In dict.Keys
    newSheet.Range("B" & lastRow).Value = Key
    newSheet.Range("C" & lastRow).Value = dict(Key)
    lastRow = lastRow + 1
Next Key

End Sub
