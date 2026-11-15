Attribute VB_Name = "deleteSheet"
Sub deleteSheet()
    MsgBox "DeleteSheet started"
    Dim name As String
    name = InputBox("Enter a sheet name to delete: ")
    Set sheet = ThisWorkbook.Sheets(name)
    sheet.Delete
    
End Sub
