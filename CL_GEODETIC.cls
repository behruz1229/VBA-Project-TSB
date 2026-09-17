VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_GEODETIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Настройки таблицы входного контроля
Private Const mask_file$ = "геодезия сводная"
Private Const sFolder$ = "\\vls.lan\ULVZG-DFS\ПТС\1.14. ТСБ и МОТ_Исполнительная документация КМ\!Ведомость элементов"

Private Const sheet_name$ = "Сводная"       'Имя листа на которой таблица
Private Const start_row As Byte = 2

'Столбцы таблицы геодезия сводная:
Private Const L_Title$ = "A"     'Титул установка
Private Const L_DP$ = "B"     'Номер DP
Private Const L_WghtDvlp$ = "C"  'Масса поставленного материала
Private Const L_WghtMont$ = "D"  'Масса смонтированного материала

Private Const L_GeoSurvey$ = "E" 'Геодезическая съемка
Private Const L_GeoPnt$ = "F"    'ИГС отрисована
Private Const L_GeoAgrd$ = "G"   'ИГС на согласовании
Private Const L_GeoAgrmnt$ = "H" 'ИГС согласованы

'Переменные класса:
Private sFileName$
Private mass_data()
Private dict_data As Object
Private WB As Excel.Workbook


'Конструктор и деструктор класса
Private Sub Class_Initialize()
    Set dict_data = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    Set dict_data = Nothing
End Sub


'Номера столбцов
Private Property Get cNumTitle%()
    cNumTitle = Columns(L_Title).Column
End Property
Private Property Get cNumDP%()
    cNumDP = Columns(L_DP).Column
End Property
Private Property Get cNumWghtDvlp%()
    cNumWghtDvlp = Columns(L_WghtDvlp).Column
End Property
Private Property Get cNumWghtMont%()
    cNumWghtMont = Columns(L_WghtMont).Column
End Property
Private Property Get cNumGeoSurvey%()
    cNumGeoSurvey = Columns(L_GeoSurvey).Column
End Property
Private Property Get cNumGeoPnt%()
    cNumGeoPnt = Columns(L_GeoPnt).Column
End Property
Private Property Get cNumGeoAgrd%()
    cNumGeoAgrd = Columns(L_GeoAgrd).Column
End Property
Private Property Get cNumGeoAgrmnt%()
    cNumGeoAgrmnt = Columns(L_GeoAgrmnt).Column
End Property


'Словарь сводной таблицы
Public Property Get DictData() As Object
    Set DictData = dict_data
End Property
Public Property Set DictData(dict As Object)
    Set dict_data = dict
End Property


'Загрузка данных в словарь данных из таблицы мастер файла ADO
Public Sub LoadData2()
        
    Dim ODBC As New CL_ODBC
    Dim OPWB As New CL_OPENFILE
    
    OPWB.SetMaskFile = mask_file
    OPWB.ListFolder.Add sFolder
    
    sFileName = OPWB.GetFileName
    If sFileName = "" Then Exit Sub 'Err.Raise 513, mask_file
    
    Application.StatusBar = "Загрузка данных из файла: " & sFileName
    Debug.Print Application.StatusBar: DoEvents
    
    If Not ODBC.SetXlFile(OPWB.GetFullFileName) Then Exit Sub
    
    ODBC.FROM = sheet_name
    
    mass_data = ODBC.Execute
    
    Set OPWB = Nothing
    Set ODBC = Nothing
    
    LoadDataInDict  'Загрузка данных из массива в словарь
    
    Application.StatusBar = False
    
End Sub

'Загрузка данных в словарь данных из таблицы мастер файла
Public Sub LoadData()

    Dim OPWB As New CL_OPENFILE
    OPWB.SetMaskFile = mask_file
    OPWB.ListFolder.Add sFolder
    
    
    sFileName = OPWB.GetFileName
    If sFileName = "" Then Exit Sub 'Err.Raise 513, mask_file
    
    Set WB = OPWB.GetWorkBook
    
    Application.StatusBar = "Загрузка данных из файла: " & sFileName
    Debug.Print Application.StatusBar: DoEvents
    
    mass_data = GetMassData  'Загрузка данных из файла в массив
    
    Set WB = Nothing
    Set OPWB = Nothing
    
    LoadDataInDict  'Загрузка данных из массива в словарь
    
    Application.StatusBar = False
    
End Sub

'Считывание массива данных из диапазона таблицы
Private Function GetMassData()
    With WB.Worksheets(sheet_name)
        If .FilterMode Then .ShowAllData
        GetMassData = .Range(L_Title & start_row).CurrentRegion. _
            Offset(start_row - 1).Resize(cnt_row - start_row + 1).Value
    End With
End Function

'Последняя строка данных в таблице
Private Property Get cnt_row&()
    cnt_row = WB.Worksheets(sheet_name).Cells(Rows.Count, cNumTitle).End(xlUp).Row
End Property

'Загрузка данных из массива в словарь
Private Sub LoadDataInDict()
    Dim q&, sKey$, mass_keys$(), dict As Object
    
    For q = LBound(mass_data) To UBound(mass_data)
        
        sKey = Trim(mass_data(q, cNumTitle))
        If sKey = "" Then GoTo continue
        
        If Not dict_data.exists(sKey) Then GoTo continue
        Set dict = dict_data(sKey)
        
        sKey = Trim(mass_data(q, cNumDP))
        If Not dict.exists(sKey) Then GoTo continue
        Set dict = dict(sKey)
        
        'Суммы, столбцы: Геодезическая съемка, ИГС отрисована, ИГС на согласовании, ИГС согласованы
        dict(idx_pv_geo_survey) = dict(idx_pv_geo_survey) + GetGeoSurvey(q)
        dict(idx_pv_geo_isdrawn) = dict(idx_pv_geo_isdrawn) + GetGeoPnt(q)
        dict(idx_pv_geo_underapp) = dict(idx_pv_geo_underapp) + GetGeoAgrd(q)
        dict(idx_pv_geo_agreed) = dict(idx_pv_geo_agreed) + GetGeoAgrmnt(q)

continue:
    Next
    
    'Расчет, столбцы: Ведется СМР, Готовая физика под геодезию (съемка)
    DoCalculationFromPivot
    
End Sub

'Расчет значений по геодезии
Private Sub DoCalculationFromPivot()
    Dim i: For Each i In dict_data.items
        calc_dp i
    Next
End Sub
Private Sub calc_dp(ByRef dict)
    Dim i: For Each i In dict.items
        i(idx_pv_geo_readphis) = GetComletedPhisic(i)
        i(idx_pv_geo_incomesmr) = GetWorkIsUnderway(i)
    Next
End Sub

'Ведется СМР '"F" => "=D5-G5-ЕСЛИ(H5="н/п";D5;H5)"
Private Property Get GetWorkIsUnderway#(ByRef dict)
    Dim val#
    val = dict(idx_pv_mount) - dict(idx_pv_geo_readphis) - dict(idx_pv_geo_survey)
    GetWorkIsUnderway = IIf(val < 0, 0, val)
End Property

'Готовая физика под геодезию '"G" => "=ЕСЛИ(D5/C5>0,8;D5-ЕСЛИ(H5="н/п";D5;H5);0)"
Private Property Get isDelivDivMont(ByRef dict) As Boolean
    If dict(idx_pv_dlvr) = 0 Then Exit Property
    If dict(idx_pv_mount) / dict(idx_pv_dlvr) > 0.8 Then isDelivDivMont = True
End Property
Private Property Get GetComletedPhisic#(ByRef dict)
    Dim val#
    If isDelivDivMont(dict) Then
        val = dict(idx_pv_mount) - dict(idx_pv_geo_survey)
        GetComletedPhisic = IIf(val < 0, 0, val)
    Else
        GetComletedPhisic = 0
    End If
End Property

'Данные весов из таблицы
Private Property Get GetGeoSurvey#(ByRef q&)
    If Not IsNumeric(mass_data(q, cNumGeoSurvey)) Then Exit Property
    GetGeoSurvey = mass_data(q, cNumGeoSurvey) * 1000
End Property
Private Property Get GetGeoPnt#(ByRef q&)
    If Not IsNumeric(mass_data(q, cNumGeoPnt)) Then Exit Property
    GetGeoPnt = mass_data(q, cNumGeoPnt) * 1000
End Property
Private Property Get GetGeoAgrd#(ByRef q&)
    If Not IsNumeric(mass_data(q, cNumGeoAgrd)) Then Exit Property
    GetGeoAgrd = mass_data(q, cNumGeoAgrd) * 1000
End Property
Private Property Get GetGeoAgrmnt#(ByRef q&)
    If Not IsNumeric(mass_data(q, cNumGeoAgrmnt)) Then Exit Property
    GetGeoAgrmnt = mass_data(q, cNumGeoAgrmnt) * 1000
End Property
