Attribute VB_Name = "UpdateTable"
Option Explicit

'======================================================
'Все настройки и логика работы таблицы База данных
'по элементам, а также сводной таблице по данному файлу
'находятся в классе CL_THISWORKBOOK
'======================================================

'======================================================
'Все настройки и логика обработки данных таблицы
'мастер файл находятся в классе CL_MF_MONT_TSB
'======================================================

'======================================================
'Все настройки и логика обработки данных таблицы входного
'контроля находятся в классе CL_INPUTCONTROL
'======================================================

Public Const debug_mode As Boolean = False

'Переменные макроса
Private THWB As New CL_THISWORKBOOK     'Экземпляр класса этой книги
Private update_mode As Byte

'Список выбираемых индексов столбцов
Public Enum col_select
    column_aosr
    column_diagram
End Enum


'Установка режима обновления
Public Property Let SetUpdMode(i_mode As Byte)
    update_mode = i_mode
End Property

'Массив уникальных значений в выбранном столбце для выпадающего титула
Public Property Get MassUniqueVal(ByRef column_select As col_select)
    MassUniqueVal = THWB.GetMassUniqueVal(column_select)
End Property

'Поиск инспекций в файле выгрузки из аис-нск
Public Sub SearchInspData(ByRef sKey$, ByRef iRow&)

    If THWB.DictDataInsp Is Nothing Then
        Dim INSP As New CL_INSPECTION
        INSP.LoadData
        Set THWB.DictDataInsp = INSP.DictData
        Set INSP = Nothing
    End If
    THWB.SearchInspData sKey, iRow
End Sub

'Таблица к САПР программе
Public Sub CadSpec(ByRef sNumberDoc$)
    
    Dim CAD As New CL_CAD_SPECIFICATION
    Dim mass()
    
    mass = CAD.GetMassDiagramm(sNumberDoc)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    CAD.CreateDocument sNumberDoc, mass
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Set CAD = Nothing
    
End Sub

'Приложение к реестру №4
Public Sub RegP03(ByRef sNumberDoc$)
    
    Dim REG3 As New CL_REGISTRY_P03
    Dim dict_ic As Object   'Данные из файлов входного контроля
    Dim mass(), tmpdict As Object, sFile$
    
    sFile = FileDialog_SaveAs(sNumberDoc & ".xlsx")
    If sFile = "" Then Exit Sub
    
    Set tmpdict = REG3.GetDownLoadedList(sNumberDoc) 'Список загружаемых параметров
    
    'Данные из файлов входного контроля
    Dim IC As New CL_INPUTCONTROLS
    IC.SetMassLoaded_Title = tmpdict.items()(0) 'Массив загружаемых титулов
    IC.SetListLoaded_DP = tmpdict.items()(1)    'Массив загружаемых dp
    IC.LoadData
    Set dict_ic = IC.DictData
    Set IC = Nothing
    
    'Сопоставление данных ведомости с данными входного контроля
    REG3.SearchInputControl dict_ic
    
    'Массив таблицы
    mass = REG3.GetMassRegistry
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    REG3.CreateDocument sNumberDoc, mass, sFile
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Set REG3 = Nothing
    Set dict_ic = Nothing
    Set tmpdict = Nothing
    
End Sub

'Диалоговое окно для выбора каталога сохранения
Private Function FileDialog_SaveAs$(ByRef sFile$)
    Dim val, sInitFilename$
    sInitFilename = Environ("USERPROFILE") & "\Desktop\" & sFile
    val = Application.GetSaveAsFilename(InitialFileName:=sInitFilename, _
        FileFilter:="Excel Files (*.xlsx), *.xlsx", _
        Title:="Выберите куда сохранить файл: " & sFile, _
        ButtonText:="Сохранить")
    If val = False Then Exit Function
    FileDialog_SaveAs = val
End Function


'Стартовая точка макроса по обновлению таблицы базы данных по элементам металла
Public Sub Prog001()
    
    Dim dict_mf As Object   'Данные из мастер файла
    Dim dict_pv As Object   'Данные из мастер файла для сводной
    Dim dict_ic As Object   'Данные из файлов входного контроля
    Dim dict_in As Object   'Данные из файла инспекции (аис-нск)
    
    If Not Attention Then Exit Sub
    
    Dim t!: t = Timer  'Запуск таймера работы макроса
    
    If Not debug_mode Then On Error GoTo Error
    
    'Данные из мастер файла МК
    Dim MFMK As New CL_MF_MONT
    MFMK.LoadData 'Загрузка данных из мастер файла
    Set dict_mf = MFMK.DictData
    Set dict_pv = MFMK.dictDataDP
    Set MFMK = Nothing
    
    
    'Данные инспекций монтажа аий-нск
    If update_mode = 2 Then GoTo skip_load_inspection
    Dim INSP As New CL_INSPECTION
    INSP.LoadData
    Set dict_in = INSP.DictData
    Set INSP = Nothing
skip_load_inspection:
    
    'Обновление базы данных по элементам
    If update_mode = 2 Then GoTo skip_load_updatetable
    Set THWB.DictData = dict_mf
    Set THWB.DictDataInsp = dict_in
    THWB.UpdateTable
    Set THWB = Nothing
skip_load_updatetable:
    
    'Данные из файлов входного контроля
    If update_mode = 1 Then GoTo skip_load_inputcontrols
    Dim IC As New CL_INPUTCONTROLS
    Set IC.DictData = dict_pv
    IC.LoadData2    'Загрузка данных из файлов входного контроля ADO
    'IC.LoadData    'Загрузка данных из файлов входного контроля WB.Open
    Set dict_ic = IC.DictData
    Set IC = Nothing
skip_load_inputcontrols:
    
    
    'Данные из файла геодезии
    If update_mode = 1 Then GoTo skip_load_geodetic
    Dim GEO As New CL_GEODETIC
    Set GEO.DictData = dict_pv
    GEO.LoadData
    Set GEO = Nothing
skip_load_geodetic:
    
    
    'Создание сводной таблицы
    If update_mode = 1 Then GoTo skip_load_pivottable
    Dim PVT As New CL_PIVOT
    Set PVT.DictDataMF = dict_pv
    Set PVT.DictDataIC = dict_ic
    PVT.MassDataTB = THWB.MassData
    PVT.UpdateTable
    Set PVT = Nothing
skip_load_pivottable:

    'Очистка переменных
    Set dict_mf = Nothing
    Set dict_pv = Nothing
    Set dict_ic = Nothing
    Set dict_in = Nothing
    
    Debug.Print "Работа макроса завершена"
    
    MsgBox "Обновление завершено." & Chr(10) & "Время работы макроса: " & _
        Format((Timer - t) / 86400, "Long Time") & Chr(10), _
        vbInformation, "Информация"
        
Exit Sub
Error: UserExeption Err
End Sub

'Вывод сообщения при запуске макроса обновления
Private Function Attention() As Boolean
    Dim sTXT$, Answer As Byte
    sTXT = "Обновление занимает много времени." & Chr(10)
    sTXT = sTXT & "Вы действительно хотите обновить?" & Chr(10)
    sTXT = sTXT & String(60, "-") & Chr(10)
    If update_mode = 0 Or update_mode = 1 Then sTXT = sTXT & Chr(149) & " " & Лист1.Name & Chr(10)
    If update_mode = 0 Or update_mode = 2 Then sTXT = sTXT & Chr(149) & " " & Лист2.Name & Chr(10)
    
    sTXT = sTXT & String(60, "-") & Chr(10)
    
    sTXT = sTXT & "Рекомендации для обновления:" & Chr(10)
    sTXT = sTXT & Chr(149) & " Сделать копию данного файла на случай, если что-то пойдет не по плану." & Chr(10)
    sTXT = sTXT & Chr(149) & " Сохранить (А лучше сохранить и закрыть) все прочие Ваши таблицы EXCEL." & Chr(10)
    sTXT = sTXT & String(60, "-") & Chr(10)
    
    sTXT = sTXT & ">> этот файл:" & ThisWorkbook.Name & " <<" & Chr(10)
    sTXT = sTXT & ">> " & Application.OperatingSystem
    sTXT = sTXT & " Версия EXCEL: " & Application.Version & " <<" & Chr(10)
    
    sTXT = sTXT & String(60, "-") & Chr(10)
    
    sTXT = sTXT & "Да   - Обновить." & Chr(10)
    sTXT = sTXT & "Нет - Не обновлять." & Chr(10)
    
    Answer = MsgBox(sTXT, vbInformation + vbDefaultButton2 + vbYesNo, "Требуется действие пользователя")
    If Answer = 6 Then Attention = True
End Function

'Процедура обработки debug ошибок макроса
Private Sub UserExeption(ByRef Er As ErrObject)
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    Select Case Er.Number
        Case 510: Er.Description = "Не найден файл: " & Er.Source
        Case 511: Er.Description = "Не найден файл журнала входного контроля: " & Er.Source
        Case 512: Er.Description = "Не найден каталог входного контроля: " & Er.Source
        Case 513: Er.Description = "Не найден файл данных геодезии: " & Er.Source
    End Select
    
    Dim sTXT$
    sTXT = "Макрос завершился с ошибкой." & Chr(10)
    sTXT = sTXT & Er.Description & IIf(Er.Source <> "", Chr(10) & Er.Source, "")
    
    MsgBox sTXT, vbCritical, "ОШИБКА (№" & Er.Number & ")"
    
    End
    
End Sub
