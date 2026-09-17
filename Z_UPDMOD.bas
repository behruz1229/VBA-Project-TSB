Attribute VB_Name = "Z_UPDMOD"
Option Explicit

Private Const sFileSetting$ = "macrosetting.ini"
Private Const sPartSet$ = "Macro setting for update"

Private cntrow&

'Экспорт модулей кода
Private Sub ExportModulesAndClasses()
    Dim vbComp As Object, sFolder$
    
    cntrow = 0  'Сброс счетчика колличества строк
    
    'Каталог для экспорта, проверка/создание
    sFolder = ThisWorkbook.Path & "\Macro"
    If Dir(sFolder, vbDirectory) = "" Then CreateDir sFolder
        
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type 'Экпорт стандартных модулей, классов, форм и код в ЭтаКнига
            Case Is = 1, 2, 3, (100 And vbComp.Name = "ЭтаКнига")
                ExportCodeModule vbComp, sFolder
        End Select
    Next

    Debug.Print "Общее колличество выгруженных строк кода в данной книге: " & cntrow
    
End Sub

'Экспорт модулей кода
Private Sub ExportCodeModule(ByRef vbComp As Object, ByRef sFolder$)

    If vbComp.Type = 100 Then
        ExportCodeToTXT vbComp, sFolder
    Else
        vbComp.Export sFolder & "\" & vbComp.Name & Choose(vbComp.Type, ".bas", ".cls", ".frm")
    End If
    logExportCode vbComp
    
End Sub

'Экспорт кода из модуля ЭтаКнига в текстовый файл sFile
Private Sub ExportCodeToTXT(ByRef vbComp As Object, ByRef sFolder$)
    Dim q&, cnt_line&, Line

    With vbComp.CodeModule
        cnt_line = .CountOfLines
        Open sFolder & "\ЭтаКнига.txt" For Output As 1
            For q = 1 To cnt_line
                 Print #1, .Lines(q, 1)
            Next
        Close #1
    End With
    
End Sub

'Логирование процесса, суммирование общего кол-во строк кода
Private Sub logExportCode(ByRef vbComp As Object)
    Dim tmpstr$: tmpstr = "SAVED: " & vbComp.Name & vbTab
    tmpstr = tmpstr & "TYPE: " & vbComp.Type & vbTab
    tmpstr = tmpstr & "CNT_LINES: " & vbComp.CodeModule.CountOfLines
    cntrow = cntrow + vbComp.CodeModule.CountOfLines
    Debug.Print tmpstr
End Sub
