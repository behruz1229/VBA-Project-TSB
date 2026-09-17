VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_INPUTCONTROLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'======================================================
' Класс для поиска файлов журнала входного контроля
' Все настройки данных входного контроля находятся здесь
'======================================================

'Каталог с файлами входного контроля
Private Const sFolder$ = "\\vls.lan\ULVZG-DFS\ПТС\1.14. ТСБ и МОТ_Исполнительная документация КМ\!Ведомость элементов\Входной контроль\db"

'Столбцы таблицы
Private Const L_Mark$ = "B"         'Марка
Private Const L_Name$ = "C"         'Наименование
Private Const L_Pos$ = "D"          'Позиция

Private Const L_Wght_MF$ = "E"      'Вес по мастер файлу
Private Const L_Wght_VK$ = "F"      'Вес по входному контролю

Private Const L_Insp_Num$ = "G"     'Номер инспекции
Private Const L_Aorpi_Name$ = "H"   'Наименование АоРПИ
Private Const L_Aorpi_Date$ = "I"   'Дата АоРПИ
Private Const L_Sert_Name$ = "J"    'Паспорт качества изделия
Private Const L_Sert_Date$ = "K"    'Дата паспорта

'Настройка таблицы
Private Const start_row As Byte = 4 'Стартовая строка данных в таблице
Private Const sPrefixFile$ = "REG_"


'Переменные класса
Private sFileName$
Private mass_data()
Private mass_loadtitles()
Private list_load_dp$
Private dict_data As Object

Private sWB_Name$
Private sSH_Name$

'Индексы в словаре с данными
Public Enum data_inputcontrols
    idx_ic_mark
    idx_ic_name
    idx_ic_pos
    idx_ic_wght_mf
    idx_ic_wght_vk
    idx_ic_insp
    idx_ic_aorpi_name
    idx_ic_aorpi_date
    idx_ic_sert_name
    idx_ic_sert_date
End Enum



'конструктор и деструктор класса
Private Sub Class_Initialize()
    Set dict_data = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    Set dict_data = Nothing
End Sub


'Номера столбцов таблицы
Private Property Get cNumMark%()
    cNumMark = Columns(L_Mark).Column
End Property
Private Property Get cNumName%()
    cNumName = Columns(L_Name).Column
End Property
Private Property Get cNumPos%()
    cNumPos = Columns(L_Pos).Column
End Property
Private Property Get cNumWght_MF%()
    cNumWght_MF = Columns(L_Wght_MF).Column
End Property
Private Property Get cNumWght_VK%()
    cNumWght_VK = Columns(L_Wght_VK).Column
End Property

Private Property Get cNumInsp_Num%()
    cNumInsp_Num = Columns(L_Insp_Num).Column
End Property
Private Property Get cNumAorpi_Name%()
    cNumAorpi_Name = Columns(L_Aorpi_Name).Column
End Property
Private Property Get cNumAorpi_Date%()
    cNumAorpi_Date = Columns(L_Aorpi_Date).Column
End Property
Private Property Get cNumSert_Name%()
    cNumSert_Name = Columns(L_Sert_Name).Column
End Property
Private Property Get cNumSert_Date%()
    cNumSert_Date = Columns(L_Sert_Date).Column
End Property


'Словарь с данными для сводной таблицы
Public Property Get DictData() As Object
    Set DictData = dict_data
End Property
Public Property Set DictData(dict As Object)
    Dim i, sTitle$, tmpdict As Object
    Set tmpdict = CreateObject("Scripting.Dictionary")
    For Each i In dict
        sTitle = Split(i)(0)
        If sTitle = "" Then GoTo continue
        If Not tmpdict.exists(sTitle) Then tmpdict.Add key:=sTitle, Item:=Empty
continue:
    Next
    mass_loadtitles = tmpdict.keys
    Set tmpdict = Nothing
End Property

'Установка списка загружаемых титулов через строку
Public Property Let SetMassLoaded_Title(ByRef mass_title)
    mass_loadtitles = mass_title
End Property

'Установка загружаемых dp разделитель vbTab
Public Property Let SetListLoaded_DP(ByRef mass_dp)
    list_load_dp = Join(mass_dp, vbTab) & vbTab
End Property

'Считывание массива данных из диапазона таблицы
Private Function GetMassData(ByRef SH As Excel.Worksheet)
    GetMassData = SH.Range(L_Mark & start_row).CurrentRegion. _
        Offset(start_row - 1).Resize(cnt_row(SH) - start_row + 1).Value
End Function

'Последняя строка данных в таблице
Private Property Get cnt_row&(ByRef SH As Excel.Worksheet)
    cnt_row = SH.Cells(Rows.Count, cNumMark).End(xlUp).Row
End Property

'Проверка наличия листа в книге
Private Function shExists(ByRef WB As Excel.Workbook, ByRef shName$) As Boolean
    On Error Resume Next
    shExists = WB.Worksheets(shName).Name = shName
End Function

'Загрузка данных через ADO
Public Sub LoadData2()
    Dim q&, ODBC As New CL_ODBC, SH, mass
    
    For q = LBound(mass_loadtitles) To UBound(mass_loadtitles)
        
        sFileName = sFolder & "\" & sPrefixFile & mass_loadtitles(q) & ".xlsx"
        
        If Not ODBC.SetXlFile(sFileName) Then GoTo continue
        
        sWB_Name = mass_loadtitles(q)
        
        Application.StatusBar = "Загрузка данных входного контроля из файла: " _
            & sWB_Name & " | (" & q + 1 & " из " & UBound(mass_loadtitles) + 1 & ")"
        Debug.Print Application.StatusBar: DoEvents
        
        For Each SH In ODBC.GetCollSheets
            sSH_Name = SH
            ODBC.FROM = SH
            mass = ODBC.Execute
            If Not IsEmpty(mass) Then
                mass_data = mass
                LoadDataInDict
            End If
        Next
        
continue:
    Next
    
    Set ODBC = Nothing
    Application.StatusBar = False
    
End Sub

'Основной метод загрузки данных по входному контролю
Public Sub LoadData()
    Dim q&, OPWB As CL_OPENFILE, WB As Excel.Workbook
    
    For q = LBound(mass_loadtitles) To UBound(mass_loadtitles)
        
        sFileName = sPrefixFile & mass_loadtitles(q)
        
        Set OPWB = New CL_OPENFILE
        OPWB.ListFolder.Add sFolder
        OPWB.SetMaskFile = sFileName
        If OPWB.GetFileName = "" Then GoTo continue
        
        Set WB = OPWB.GetWorkBook
        
        sWB_Name = WB.Name
        
        Application.StatusBar = "Загрузка данных входного контроля из файла: " _
            & sWB_Name & " | (" & q + 1 & " из " & UBound(mass_loadtitles) + 1 & ")"
        Debug.Print Application.StatusBar: DoEvents
        
        LoadTable WB
        
continue:

        Set WB = Nothing
        Set OPWB = Nothing
        
    Next
    
    Application.StatusBar = False
    
End Sub

'Загрузка данных из таблицы. Чтение данных листов книги
Private Sub LoadTable(ByRef WB As Excel.Workbook)
    Dim SH As Excel.Worksheet
    
    For Each SH In WB.Worksheets
        If cnt_row(SH) < start_row Then GoTo continue
        
        If list_load_dp <> "" And InStr(1, list_load_dp, SH.Name & vbTab, vbTextCompare) = 0 Then GoTo continue
        
        sSH_Name = SH.Name
        mass_data = GetMassData(SH)
        LoadDataInDict
        
continue:
    Next
End Sub

'Загрузка данных в словарь
Private Sub LoadDataInDict()
    Dim q&, sKey$
        
    For q = LBound(mass_data) To UBound(mass_data)
        
        If Trim(mass_data(q, cNumMark)) = "" Then GoTo continue
        If Trim(mass_data(q, cNumName)) = "" Then GoTo continue
        
        sKey = GetKeyElement(q)
        
        If Not dict_data.exists(sKey) Then dict_data.Add key:=sKey, Item:=tmpdict(q)
        
continue:
    Next
    
End Sub

'Ключ для словаря с элементами титул, dp, марка, наименование, позиция, вес
Private Property Get GetKeyElement$(ByRef q&)
    Dim sKey$
    sKey = GetTitle(sWB_Name) & vbTab
    sKey = sKey & sSH_Name & vbTab
    sKey = sKey & mass_data(q, cNumMark) & vbTab
    sKey = sKey & mass_data(q, cNumName) & vbTab
    sKey = sKey & mass_data(q, cNumPos) & vbTab
    sKey = sKey & mass_data(q, cNumWght_MF)
    GetKeyElement = sKey
End Property

'Данные в словаре
Private Function tmpdict(ByRef q&) As Object
    Set tmpdict = CreateObject("Scripting.Dictionary")
    With tmpdict
        .Add key:=.Count, Item:=mass_data(q, cNumMark)
        .Add key:=.Count, Item:=mass_data(q, cNumName)
        .Add key:=.Count, Item:=mass_data(q, cNumPos)
        
        .Add key:=.Count, Item:=mass_data(q, cNumWght_MF)
        .Add key:=.Count, Item:=mass_data(q, cNumWght_VK)
        
        .Add key:=.Count, Item:=mass_data(q, cNumInsp_Num)
        .Add key:=.Count, Item:=mass_data(q, cNumAorpi_Name)
        .Add key:=.Count, Item:=mass_data(q, cNumAorpi_Date)
        
        .Add key:=.Count, Item:=mass_data(q, cNumSert_Name)
        .Add key:=.Count, Item:=mass_data(q, cNumSert_Date)
    End With
End Function

'Поиск даты в строке материала
Private Property Get GetTitle$(ByRef sTXT)
    Dim objExec As Object, tmpstr$
    Set objExec = RegExp(sTXT, "[A-Z]\d-\d{2}-?[\d\w]*")
    If objExec.Count = 1 Then
        tmpstr = objExec(0).Value
        GetTitle = tmpstr
    End If
    Set objExec = Nothing
End Property

'Свойство создания (настроек) объекта регулярных выражений
Private Property Get RegExp(ByRef sTXT, ByRef sMask$) As Object
    With CreateObject("VBScript.RegExp")
        .Global = True 'Поиск по всему тексту
        .Pattern = sMask 'Шаблон поиска
        .IgnoreCase = True
        Set RegExp = .Execute(sTXT)
    End With
End Property
