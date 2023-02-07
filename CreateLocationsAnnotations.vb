Sub CreateXML()

Dim i As Long
Dim lastRow As Long
Dim location As String
Dim layerName As String
Dim layerColor As String

'Find the last row with data in column H
lastRow = Cells(Rows.Count, 8).End(xlUp).Row

'Loop through the list of locations in column H
For i = 1 To lastRow

    location = Cells(i, 8).Value
    layerName = "Layer" & location
    layerColor = RGB(255, 0, 0)
    
    'Build the XML string
    XMLString = XMLString & "<Annotation>" & vbCrLf & _
                            "  <Layer>" & layerName & "</Layer>" & vbCrLf & _
                            "  <LayerColor>" & layerColor & "</LayerColor>" & vbCrLf & _
                            "</Annotation>" & vbCrLf

Next i

'Write the XML string to a new file
Open "C:\output.xml" For Output As #1
Print #1, XMLString
Close #1

End Sub
