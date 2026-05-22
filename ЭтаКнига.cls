VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Dim back_up As New CL_BACKUP
    If Success And Not ThisWorkbook.ReadOnly Then back_up.GetBackUpTable: Set back_up = Nothing
End Sub

Private Sub Workbook_Deactivate()
    ResetCommandBar
End Sub

Private Sub Workbook_Open()
    Application.EnableEvents = True
End Sub

Private Sub Workbook_SheetChange(ByVal SH As Object, ByVal Target As Range)
    
    Dim RG As Range
    
    If SH Is Лист1 Then
        
        If Target.Rows.Count = Rows.Count Then Exit Sub
        If Target.Columns.Count = Columns.Count Then Exit Sub
        If Application.CutCopyMode Then Exit Sub
        
        Application.EnableEvents = False
        
        For Each RG In Target
            If RG.Column <> idx_insp_smu Then GoTo continue
            If RG.Value = "" Then GoTo continue
            SearchInspData RG.Value, RG.Row
continue:
        Next
        Application.EnableEvents = True
    End If
End Sub



Private Sub Workbook_SheetBeforeRightClick(ByVal SH As Object, ByVal Target As Range, Cancel As Boolean)
    Dim IndexCmd%, q&, mass
    
    ResetCommandBar 'Default меню правой кнопки мышки
    
    If Target.Rows.Count = Rows.Count Then Exit Sub
    If Target.Columns.Count = Columns.Count Then Exit Sub
    If Application.CutCopyMode Then Exit Sub
    
    IndexCmd = IndexCommandBar 'индекс контекстстного меню
    If IndexCmd = 0 Then Exit Sub
    
    'Отделение пользовательских меню группой
    Application.CommandBars(IndexCmd).Controls(1).BeginGroup = True
    
    'Создание пункта меню обновления таблицы Ведомости
    If KeyPressed(VK_CONTROL) And (SH Is Лист1 Or SH Is Лист2) Then
        
        UpdateTable.SetUpdMode = IIf(SH Is Лист1, 1, 2)
        
        With Application.CommandBars(IndexCmd).Controls.Add _
            (msoControlButton, Before:=1, Temporary:=True)
            .Caption = "Обновить: " & SH.Name
            .FaceId = 2608
            .OnAction = "'UpdateTable.Prog001'"
        End With
    
    ElseIf SH Is Лист1 Or SH Is Лист2 Then
        
        UpdateTable.SetUpdMode = 0
        
        With Application.CommandBars(IndexCmd).Controls.Add _
            (msoControlButton, Before:=1, Temporary:=True)
            .Caption = "Обновить таблицы"
            .FaceId = 2608
            .OnAction = "'UpdateTable.Prog001'"
        End With
        
        ' ------------------- новые пункты RFI (только для листа "База данных по элементам") -------------------
        If SH.Name = "База данных по элементам" Then
            ' Добавляем сначала "Обновить статус RFI", чтобы после обратного порядка он оказался ниже
            With Application.CommandBars(IndexCmd).Controls.Add _
                (msoControlButton, Before:=2, Temporary:=True)
                .Caption = "Обновить статус RFI"
                .FaceId = 2608
                .OnAction = "Update_status_RFI.Update_status_RFI"
            End With
            ' Теперь "Обновить исх. № RFI" встанет на вторую позицию, выше
            With Application.CommandBars(IndexCmd).Controls.Add _
                (msoControlButton, Before:=2, Temporary:=True)
                .Caption = "Обновить исх. № RFI"
                .FaceId = 2608
                .OnAction = "Update_SMU_MK_RFI.Update_SMU_MK_RFI"
            End With
        End If
      
    End If
    
    'Создание пунктов меню списка выгружаемых документов
    If SH Is Лист1 Then
        
        'Создание пункта меню Приложение к АОСР п.3
        On Error GoTo skip_step_1
        mass = UpdateTable.MassUniqueVal(column_aosr)
        If IsArray(mass) Then
            With Application.CommandBars(IndexCmd).Controls.Add _
                (msoControlPopup, Before:=1, Temporary:=True)
                .Caption = "&Приложение к АООК п.4"
                .Enabled = True
                For q = LBound(mass) To UBound(mass)    'Создание под-пунктов меню
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = mass(q)
                        .FaceId = 593
                        .OnAction = "'UpdateTable.RegP03 """ & mass(q) & """'"
                    End With
                Next
            End With
        End If
skip_step_1:
        
        'Создание пункта меню Спецификация для NanoCAD
        On Error GoTo skip_step_2
        mass = UpdateTable.MassUniqueVal(column_diagram)
        If IsArray(mass) Then
            With Application.CommandBars(IndexCmd).Controls.Add _
                (msoControlPopup, Before:=1, Temporary:=True)
                .Caption = "&Спецификация для САПР"
                For q = LBound(mass) To UBound(mass)    'Создание под-пунктов меню
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = mass(q)
                        .FaceId = 593
                        .OnAction = "'UpdateTable.CadSpec """ & mass(q) & """'"
                    End With
                Next
            End With
        End If
skip_step_2:

    End If
    
    'Создание пункта меню экспорта листа в xslx\pdf
    If Лист2 Is SH Or Not SH Is Лист1 Then
        If KeyPressed(VK_CONTROL) Or ActiveWindow.View = xlPageBreakPreview Then
        
            With Application.CommandBars(IndexCmd).Controls.Add _
                (msoControlButton, Before:=1, Temporary:=True)
                .Caption = "Экспорт листа в .pdf"
                .FaceId = 268
                .OnAction = "'ExportSheet.ExportActiveSheet True'"
            End With
            
        Else
        
            With Application.CommandBars("Cell").Controls.Add _
                (msoControlButton, Before:=1, Temporary:=True)
                .Caption = "Экспорт листа в .xlsx"
                .FaceId = 263
                .OnAction = "'ExportSheet.ExportActiveSheet False'"
            End With
            
        End If
    End If
    
End Sub

'Контекстное меню в разном режиме страницы (индексы меню, имена Cell)
Private Function IndexCommandBar%()
    Dim q As Byte, cmd As CommandBar
    
    For Each cmd In Application.CommandBars
        If cmd.Name = "Cell" Then
            q = q + 1
            Debug.Print "Name: " & cmd.Name, "Index: " & cmd.Index
            If q = ActiveWindow.View Then IndexCommandBar = cmd.Index: Exit Function
        End If
    Next
End Function

'Восстановление меню на правую кнопку
Private Sub ResetCommandBar()
    Dim cmb As CommandBar
    On Error Resume Next
    For Each cmb In Application.CommandBars
        cmb.Reset
    Next
End Sub
