Attribute VB_Name = "ExportSheet"
Option Explicit

'======================================================
'Экспорт прочих листов в .pdf или .xlsx
'======================================================

Private Const xlFiles = "Excel Files (*.xlsx), *.xlsx"
Private Const pdFiles = "PDF Files (*.pdf), *.pdf"

Private SH As Worksheet

'Функция экспорта листов из книги
Public Sub ExportActiveSheet(ByRef bFormat As Boolean, Optional ByRef sFileName$)
    Dim sFile$, sName$
    
    Set SH = ActiveSheet
    sName = SH.Name & "_" & Format(Date - 1, "yyyy-mm-dd")
    
    On Error GoTo en
    If bFormat Then
        sFile = ShowFileDialog(sName & ".pdf", pdFiles)
        If sFile = "" Then Exit Sub
        to_pdf SH, sFile
    Else
        sFile = ShowFileDialog(sName & ".xlsx", xlFiles)
        If sFile = "" Then Exit Sub
        to_excel SH, sFile
    End If
en:
End Sub

'Функция выбора каталога и файла сохранения
Private Function ShowFileDialog$(ByRef sNameFiles$, ByRef TypeFile$)
    Const sTitle$ = "Сохранения файла"
    Dim tmp
    tmp = Application.GetSaveAsFilename(sNameFiles, FileFilter:=TypeFile, Title:=sTitle, ButtonText:="Экспорт")
    If tmp = False Then Exit Function
    ShowFileDialog = tmp
End Function

'Экспорт в pdf
Private Sub to_pdf(ByRef SH As Worksheet, ByRef sFile$)
    SH.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=sFile, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=True, _
    OpenAfterPublish:=True
End Sub

'Экспорт в xlsx
Private Sub to_excel(ByRef SH As Worksheet, ByRef sFile$)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    SH.Copy
    ActiveWorkbook.SaveAs Filename:=sFile, FileFormat:=xlWorkbookDefault
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub


