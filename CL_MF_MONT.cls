VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_MF_MONT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



'Настройки таблицы мастер файла:
Private Const mask_file$ = "ТСБ - Монтаж МК - Мастер файл"
Private Const sFolder$ = "~\!Ведомость элементов"

'Столбцы мастер файла
Private Const N_MF_DP% = 1          'DP
Private Const N_MF_Mark% = 4        'Марка элементов
Private Const N_MF_Name% = 6        'Наименование элементов
Private Const N_MF_Unit% = 7        'Единица измерения
Private Const N_MF_Pos% = 8         'Порядковый номер элементов
Private Const N_MF_Wght% = 9        'Масса элемента
Private Const N_MF_NumDlvr% = 14    'Номер поставки эемента
Private Const N_MF_DataDlvr% = 15   'Дата поставки элемента
Private Const N_MF_NumMont% = 20    'Номер монтажа элемента
Private Const N_MF_DataMont% = 21   'Дата монтажа элемента
Private Const N_MF_Note% = 45       'Примечание


'Переменные класса:
Private sTitle$
Private sFileName$
Private mass_data()
Private dict_data As Object
Private dict_dp As Object
Private WB As Excel.Workbook
Private SH As Excel.Worksheet

'Конструктор и деструктор класса
Private Sub Class_Initialize()
    Set dict_data = CreateObject("Scripting.Dictionary")
    Set dict_dp = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    Set dict_data = Nothing
    Set dict_dp = Nothing
End Sub


'Словарь с элементами
Public Property Get DictData() As Object
    Set DictData = dict_data
End Property

'Словарь для сводной
Public Property Get dictDataDP() As Object
    Set dictDataDP = dict_dp
End Property

'Поиск открытие таблицы с данными
Public Sub LoadData()
    
    Dim OPWB As New CL_OPENFILE
    OPWB.SetMaskFile = mask_file
    OPWB.ListFolder.Add sFolder
    OPWB.ShowFileDialog = True
    
    Set WB = OPWB.GetWorkBook
    sFileName = OPWB.GetFileName
    
    If sFileName = "" Then Err.Raise 510, mask_file
    
    ReadWorksheets
    
    Set WB = Nothing
    Set OPWB = Nothing
    
End Sub

'Считывание данных из таблицы: Master file
Private Sub ReadWorksheets()
      
    'Перебор всех листов и считывание данных умных таблиц в словарь
    For Each SH In WB.Worksheets
        If Not SH.ListObjects.Count = 1 Then GoTo continue
        If SH.Name = "X-YY0000" Then GoTo continue
        
        ReadListObjects SH
        
continue:
    Next
    
    Application.StatusBar = False
    
End Sub

'Считывание данных из smart-таблиц
Private Sub ReadListObjects(ByRef SH As Excel.Worksheet)
    Dim LO As Excel.ListObject
    
    For Each LO In SH.ListObjects
        If LO.DataBodyRange.Columns.Count < N_MF_Note Then GoTo continue
        
        sTitle = Split(SH.Name)(0)
        
        Application.StatusBar = "Загрузка листа: " & SH.Name & _
            ", Строк данных таблицы: " & LO.DataBodyRange.Rows.Count
        Debug.Print Application.StatusBar: DoEvents
        
        mass_data = LO.DataBodyRange.Value
        LoadMassToDict
    
continue:
    Next
    
End Sub

'Загрузка данных в словарь
Private Sub LoadMassToDict()
    Dim q&, sKey$
    
    For q = LBound(mass_data) To UBound(mass_data)
        If Not IsNumeric(mass_data(q, N_MF_Wght)) Then GoTo continue
        If mass_data(q, N_MF_Wght) = 0 Then GoTo continue
        
        LoadDictData q          'Элементы для базы данных
        LoadDictDataFromPivot q    'Данные для сводной таблицы
        
continue:
    Next

End Sub

'Данные для сводной таблицы
Private Sub LoadDictDataFromPivot(ByRef q&)
    Dim sKey$, coll_keys As New Collection
    
    '1. Лист таблицы в мастер файле
    sKey = Trim_(SH.Name)
    coll_keys.Add sKey
    
    '2. Номер DP
    sKey = Trim_(mass_data(q, N_MF_DP))
    If sKey = "" Then sKey = "NO_DP"
    coll_keys.Add sKey
    
    InsertLevelDict q, dict_dp, coll_keys
    
End Sub

'Формирование уровней вложения в словарь
Private Sub InsertLevelDict(ByRef q&, ByRef dict As Object, ByRef coll_keys As Collection, Optional ByVal level% = 0)
    Dim sKey$
    
    level = level + 1       'Счетчик уровней вложения
    sKey = coll_keys(level)  'Ключ словаря из коллекции
    
    If level = coll_keys.Count Then  'Вставление значения счетчика (шт.) на последнем ключе
        If Not dict.exists(sKey) Then dict.Add key:=sKey, Item:=CreateObject("Scripting.Dictionary")
        AddPivotData q, dict(sKey)
    Else
        If Not dict.exists(sKey) Then dict.Add key:=sKey, Item:=CreateObject("Scripting.Dictionary")
        InsertLevelDict q, dict(sKey), coll_keys, level
    End If
End Sub

'Формирование данных словаря для сводной таблицы
Private Sub AddPivotData(ByRef q&, ByRef dict As Object)
    
    '1. Всего
    dict(idx_pv_all) = dict(idx_pv_all) + mass_data(q, N_MF_Wght)
    
    '2. Поставлено
    If Not mass_data(q, N_MF_NumDlvr) = Empty And Not mass_data(q, N_MF_NumDlvr) = False Then
        dict(idx_pv_dlvr) = dict(idx_pv_dlvr) + mass_data(q, N_MF_Wght)
    Else
        dict(idx_pv_dlvr) = dict(idx_pv_dlvr) + 0
    End If
        
    '3. Смонтированно
    If IsDate(mass_data(q, N_MF_DataMont)) Then
        dict(idx_pv_mount) = dict(idx_pv_mount) + mass_data(q, N_MF_Wght)
    Else
        dict(idx_pv_mount) = dict(idx_pv_mount) + 0
    End If
    
    '4. Остаток от поставленного
    dict(idx_pv_remain) = dict(idx_pv_dlvr) - dict(idx_pv_mount)
    
End Sub

'Элементы для базы данных по элементам для ведомости
Private Sub LoadDictData(ByRef q&)
    Dim sKey$
    
    'Проверка наличия номера поставки элемента
    If mass_data(q, N_MF_NumDlvr) = Empty Then Exit Sub
    If mass_data(q, N_MF_NumDlvr) = False Then Exit Sub
    
    'Ключ (Наименование листа, номер dp, марка элемента, наименование элемента, вес)
    sKey = GetKeyElement(q)
    
    'Заполение словаря данными из мастер файла
    If Not dict_data.exists(sKey) Then dict_data.Add key:=sKey, Item:=New Collection
    dict_data(sKey).Add tmpdict(q)
    
End Sub

'Данные для базы данных материалов части мастер файла
Public Function tmpdict(ByRef q&) As Object
    Set tmpdict = CreateObject("Scripting.Dictionary")
    
    With tmpdict
        
        .Add key:=.Count + 1, Item:=sTitle
        .Add key:=.Count + 1, Item:=Trim_(SH.Name)
        
        .Add key:=.Count + 1, Item:=Trim_(mass_data(q, N_MF_DP))
        .Add key:=.Count + 1, Item:=Trim_(mass_data(q, N_MF_Mark))
        .Add key:=.Count + 1, Item:=Trim_(mass_data(q, N_MF_Name))
        
        .Add key:=.Count + 1, Item:=mass_data(q, N_MF_Pos)
        .Add key:=.Count + 1, Item:=mass_data(q, N_MF_Wght)
        .Add key:=.Count + 1, Item:=Trim_(mass_data(q, N_MF_Note))
        
        If IsDate(mass_data(q, N_MF_DataMont)) Then
            .Add key:=.Count + 1, Item:=CDate(mass_data(q, N_MF_DataMont))
        Else
            .Add key:=.Count + 1, Item:=Empty
        End If
        
    End With
    
End Function

'Ключ для словаря с элементами
Private Property Get GetKeyElement$(ByRef q&)
    Dim sKey$
    sKey = sTitle & vbTab
    sKey = sKey & Trim_(mass_data(q, N_MF_DP)) & vbTab
    sKey = sKey & Trim_(mass_data(q, N_MF_Mark)) & vbTab
    sKey = sKey & Trim_(mass_data(q, N_MF_Name)) & vbTab
    sKey = sKey & mass_data(q, N_MF_Wght)
    GetKeyElement = sKey
End Property

'Функция убрать пробелы вначале и конце и замена 10, 13, 160 символов на пробел (33)
Private Property Get Trim_$(ByVal sTXT)
    If IsError(sTXT) Then Exit Property
    If InStr(1, sTXT, Chr(10), vbBinaryCompare) > 0 Then sTXT = Replace(sTXT, Chr(10), Chr(32))
    If InStr(1, sTXT, Chr(13), vbBinaryCompare) > 0 Then sTXT = Replace(sTXT, Chr(13), Chr(32))
    If InStr(1, sTXT, Chr(160), vbBinaryCompare) > 0 Then sTXT = Replace(sTXT, Chr(160), Chr(32))
    Trim_ = Trim(sTXT)
End Property
