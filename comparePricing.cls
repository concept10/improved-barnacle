' Excel macro used to create a report of price differences between two sources.
' To use, ensure that the developer options are enabled in Excel.  Ensure that VBA is enabled for the workspace you want to use this in.
' Then, open the VBA editor and import this file. Run the ComparePrices() macro.

#Region 

Sub ComparePrices()

Dim ws As Worksheet
Dim rngA As Range, rngE As Range
Dim cellA As Range, cellE As Range
Dim dictA As Object, dictE As Object
Dim priceA As Double, priceE As Double
Dim delta As Double
Dim row As Long

Dim reportOutputWorksheet As String = "Report5000"

#End Region

Set dictA = CreateObject("Scripting.Dictionary")
Set dictE = CreateObject("Scripting.Dictionary")

For Each ws In ThisWorkbook.Worksheets
If ws.Name Like "Sheet1" Then
Set rngA = ws.Range("A1:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
Set rngE = ws.Range("E1:E" & ws.Cells(ws.Rows.Count, "E").End(xlUp).row)

    For Each cellA In rngA
        If Not dictA.Exists(cellA.Value) Then
            If IsNumeric(ws.Cells(cellA.row, 2)) Then
                priceA = ws.Cells(cellA.row, 2)
            Else
                priceA = 0
            End If
            dictA.Add cellA.Value, priceA
        End If
    Next cellA
    
    For Each cellE In rngE
        If Not dictE.Exists(cellE.Value) Then
            If IsNumeric(ws.Cells(cellE.row, 6)) Then
                priceE = ws.Cells(cellE.row, 6)
            Else
                priceE = 0
            End If
            dictE.Add cellE.Value, priceE
        End If
    Next cellE
    
    For Each Key In dictA.Keys
        If dictE.Exists(Key) Then
            delta = dictA(Key) - dictE(Key)
            row = row + 1
           
            ThisWorkbook.Sheets(reportOutputWorksheet).Cells(row, 1).Value = Key
            ThisWorkbook.Sheets(reportOutputWorksheet).Cells(row, 2).Value = delta
            ThisWorkbook.Sheets(reportOutputWorksheet).Cells(row, 3).Value = "Source A: " & dictA(Key) & "  Source E: " & dictE(Key)
            
        End If
    Next Key
    
End If

Next ws

End Sub
