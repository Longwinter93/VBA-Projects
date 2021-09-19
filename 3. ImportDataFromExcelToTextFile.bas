Attribute VB_Name = "Module4"
Sub DataTransferWithComma()
Dim FileName As String
Dim i As Integer
Dim j As Integer
Dim CurrDate As Date

Dim lineText As String
Dim my_range As Range
CurrDate = Date

FileName = ThisWorkbook.Path & "\Data From Excel " & Date & ".txt"

Open FileName For Output As #1
Set my_range = Worksheets(1).Range("A1:E12")

For i = 1 To 12
    For j = 1 To 5
    lineText = IIf(j = 1, "", lineText & ",") & my_range.Cells(i, j)
    
    
    Next j
    Print #1, lineText
Next i



Close #1

MsgBox ("Data Transfer is completed " & Date)

End Sub
