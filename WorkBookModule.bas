Attribute VB_Name = "WorkBookModule"
Dim outputTemplate As String
Dim Workbook As Workbook

Function workbookTemplate() As Workbook
    outputTemplate = "C:\Users\Bui Dinh BACH\Downloads\Macro\2_fb_out.xlsx"
    'Set workbookTemplate = Workbooks.Add(outputTemplate)
End Function

Sub closeWorkBook(wkb As Workbook)
    wkb.Close
End Sub


