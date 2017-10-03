Attribute VB_Name = "ExcelModule"
Function OpenFileName(TitleNM As String, JOB As Integer) As Variant
    Dim Xls As Application
    Set Xls = Excel.Application
    If JOB = 1 Then
        OpenFileName = Xls.GetOpenFilename(fileFilter:="Excelファイル(*.xlsx),*.xlsx", Title:=TitleNM, MultiSelect:=False)
    End If

    If OpenFileName = False Then
        MsgBox "キャンセルしました。", , TitleNM
    End If
    Set Xls = Nothing
End Function

Sub SaveExcel(wkb As Workbook, sFileSaveName As Variant)
    sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=InitialName, fileFilter:="Excel Files (*.xlsx), *.xlsm")
    cellOutRange = sFileSaveName
    If sFileSaveName <> False Then
        wkb.SaveAs sFileSaveName
    End If
End Sub
