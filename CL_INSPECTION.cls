VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_INSPECTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Настройки файла:
Private Const mask_file$ = "инспекции"
Private Const sFolder$ = "\\vls.lan\ULVZG-DFS\ПТС\1.14. ТСБ и МОТ_Исполнительная документация КМ\!Ведомость элементов"

'Настройки таблицы:
Private Const sheet_name$ = "Инспекции"
Private Const start_row As Byte = 2

'Столбцы с данными
Private Const L_ID$ = "A" 'Служебный столбец для счета последней строки таблицы
Private Const L_Insp_NumAis$ = "C"    'Номера инспекций
Private Const L_Insp_NumSmu$ = "D"    'Исходящий номер инспекий сму

Private Const L_Insp_StatGSP$ = "AG"  'Статус RFI ГСП
Private Const L_Insp_StatSK$ = "AH"   'Статус RFI СК
Private Const L_Insp_ResltGSP$ = "AS" 'Заключение RFI ГСП
Private Const L_Insp_ResltSK$ = "AT"  'Заключение RFI СК

Private Const L_Insp_Date$ = "L"      'Столбец дата подачи

'Переменные класса
Private sFileName$
Private mass_data()
Private dict_data As Object
Private WB As Excel.Workbook

'Настройки словаря
Private Enum CompareMode
    BinaryCompare
    TextCompare
End Enum

'Конструктор и деструктор класса:
Private Sub Class_Initialize()
    Set dict_data = CreateObject("Scripting.Dictionary")
    dict_data.CompareMode = TextCompare
End Sub
Private Sub Class_Terminate()
    Set dict_data = Nothing
End Sub


'Номера столбцов таблицы
Private Property Get cNumID%()
    cNumID = Columns(L_ID).Column
End Property
Private Property Get cNumInsp_NumAis%()
    cNumInsp_NumAis = Columns(L_Insp_NumAis).Column
End Property
Private Property Get cNumInsp_NumSmu%()
    cNumInsp_NumSmu = Columns(L_Insp_NumSmu).Column
End Property
Private Property Get cNumInsp_StatGSP%()
    cNumInsp_StatGSP = Columns(L_Insp_StatGSP).Column
End Property
Private Property Get cNumInsp_StatSK%()
    cNumInsp_StatSK = Columns(L_Insp_StatSK).Column
End Property
Private Property Get cNumInsp_ResltGSP%()
    cNumInsp_ResltGSP = Columns(L_Insp_ResltGSP).Column
End Property
Private Property Get cNumInsp_ResltSK%()
    cNumInsp_ResltSK = Columns(L_Insp_ResltSK).Column
End Property

Private Property Get cNumDateSubmiss%()
    cNumDateSubmiss = Columns(L_Insp_Date).Column
End Property


'Область табличных данных
Private Property Get GetRangeTable(ByRef L_Column$) As Range
    Set GetRangeTable = WB.Worksheets(sheet_name).Range(L_ID & start_row).CurrentRegion. _
        Offset(start_row - 1).Resize(cnt_row - start_row + 1)
End Property

'Загрузка табличных данных в массив
Private Function GetMassInTable()
    GetMassInTable = GetRangeTable(L_ID).Value
End Function

'Последняя строка данных в таблице
Private Property Get cnt_row&()
    cnt_row = WB.Worksheets(sheet_name).Cells(Rows.Count, cNumID).End(xlUp).Row
End Property

'Словарь с данными инспекций
Public Property Get DictData() As Object
    Set DictData = dict_data
End Property

'Загрузка данных в словарь из таблицы
Public Sub LoadData()

    Dim q As Byte, sStatusBar$, tmpstr$, OPWB As New CL_OPENFILE
    OPWB.ListFolder.Add sFolder
    
    For q = 0 To 29
                
        'Информация о поиске файла
        sStatusBar = "Поиск файла: " & mask_file & " на " & GetLastDate(q) & " число..."
        If InStr(1, tmpstr, sStatusBar & vbTab, vbTextCompare) = 0 Then
            tmpstr = tmpstr & sStatusBar & vbTab
            Application.StatusBar = sStatusBar
            Debug.Print Application.StatusBar: DoEvents
        End If
        
        OPWB.SetMaskFile = mask_file & "*" & GetLastDate(q) & ".xls?"
        sFileName = OPWB.GetFileName
        
        If sFileName <> "" Then Exit For
    Next
    
    If sFileName = "" Then
        OPWB.SetMaskFile = mask_file
        OPWB.ShowFileDialog = True
        sFileName = OPWB.GetFileName
    End If
    
    If sFileName = "" Then
        If MsgBox("Файл с ниспекциями не был найден." & Chr(10) & String(30, "-") & Chr(10) & _
         "Да - продолжить" & Chr(10) & "Нет - остановить выполнение макроса", _
         vbExclamation + vbDefaultButton1 + vbYesNo, "Действие пользователя") = vbNo Then
            Err.Raise 510, mask_file
        End If
        Exit Sub
    End If
    
    'Способ загрузки данных из файла через ADO подключение
    Dim ODBC As New CL_ODBC
    If Not ODBC.SetXlFile(OPWB.GetFullFileName) Then Exit Sub
    
    Application.StatusBar = "Загрузка данных из таблицы: " & sFileName
    Debug.Print Application.StatusBar: DoEvents
    
    ODBC.FROM = sheet_name
    
    mass_data = ODBC.Execute
    
    Set OPWB = Nothing
    Set ODBC = Nothing
    
    LoadDataInDict
    Application.StatusBar = False
    
Exit Sub
    'Способ загрузки данных из файла через открытие книги
    Set WB = OPWB.GetWorkBook

    Application.StatusBar = "Загрузка данных из таблицы: " & sFileName
    Debug.Print Application.StatusBar: DoEvents

    mass_data = GetMassInTable
    LoadDataInDict

    Set WB = Nothing
    Set OPWB = Nothing
    
    Application.StatusBar = False
    
End Sub
Private Property Get GetLastDate$(ByRef Decrement As Byte)
    GetLastDate = Format(Date - Decrement, "dd.mm.yy")
End Property

'Загрузка данных в словарь
Private Sub LoadDataInDict()
    Dim q&, sKey$
    
    For q = LBound(mass_data) To UBound(mass_data)
        
        If Not IsDate(mass_data(q, cNumDateSubmiss)) Then GoTo continue   'Пропуск инспекций без даты
        If Not CDate(mass_data(q, cNumDateSubmiss)) > CDate("01.06.2023") Then GoTo continue 'Пропуск инспекций раньше от 01.06.2023
        
        sKey = Trim(mass_data(q, cNumInsp_NumSmu))    'исходящий номер сму
        If sKey = "" Then GoTo continue     'Пропуск пустых инспекций
        
        If Not dict_data.exists(sKey) Then
            dict_data.Add key:=sKey, Item:=CreateObject("Scripting.Dictionary")
            With dict_data(sKey)
                .Add key:=.Count, Item:=Trim(mass_data(q, cNumInsp_NumSmu))
                .Add key:=.Count, Item:=Trim(mass_data(q, cNumInsp_NumAis))
                .Add key:=.Count, Item:=Trim(mass_data(q, cNumInsp_StatGSP))
                .Add key:=.Count, Item:=Trim(mass_data(q, cNumInsp_StatSK))
                .Add key:=.Count, Item:=Trim(mass_data(q, cNumInsp_ResltGSP))
                .Add key:=.Count, Item:=Trim(mass_data(q, cNumInsp_ResltSK))
            End With
        End If

continue:
    Next
    
End Sub


