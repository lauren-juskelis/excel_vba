Attribute VB_Name = "Module1"
Sub DeleteColumns()

' Source 1: https://www.extendoffice.com/documents/excel/3086-excel-delete-columns-based-on-header.html
' Source 2: https://www.thespreadsheetguru.com/the-code-vault/a-vba-for-loop-in-reverse-order

Dim i As Long

For i = 489 To 1 Step -1
    If ActiveSheet.Cells(2, i).Value <> "01" Then ActiveSheet.Cells(2, i).EntireColumn.Delete
Next i

End Sub
