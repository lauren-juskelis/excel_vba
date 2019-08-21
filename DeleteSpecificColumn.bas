Attribute VB_Name = "Module1"
Sub DeleteSpecifcColumn()
    Set MR = Range("I2:RU2")
    For Each cell In MR
        If cell.Value = "02" Then cell.EntireColumn.Delete
    Next
End Sub

