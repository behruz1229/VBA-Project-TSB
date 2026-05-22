VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_THISWORKBOOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'======================================================
'Класс для обработки данных таблицы Ведомость МК
'Все настройки таблицы Ведомость МК находятся здесь
'======================================================

'Столбцы таблицы:
Private Const L_Title$ = "A"    'Титул
Private Const L_Sheet$ = "B"    'Наименование листа с таблицей
Private Const L_DP$ = "C"       'Номер DP элемента
Private Const L_Mark$ = "D"     'Марка элемента
Private Const L_Name$ = "E"     'Наименование элемента
Private Const L_Pos$ = "F"      'Счетчик элементов (Позииция элемента)
Private Const L_Wght$ = "G"     'Вес элемента
Private Const L_Note$ = "H"     'Примечание из мастер файла
Private Const L_Date$ = "I"     'Дата монтажа
Private Const L_DateId$ = "J"   'Дата исполнительной
'------------------------------------------------------
Private Const L_AxesLetter$ = "K"       'Оси Буква
Private Const L_AxesNumber$ = "L"       'Оси цифра
Private Const L_AltitudeMark$ = "M"     'Высотная отметка
'------------------------------------------------------
Private Const L_NumMontDiagram$ = "N"   'Номер монтажной схемы
'------------------------------------------------------
Private Const L_Insp_NumSmu$ = "O"      'Исходящий номер сму инспекций
Private Const L_Insp_NumAis$ = "P"      '№ RFI на монтаж
Private Const L_Insp_StatGSP$ = "Q"     'Статус RFI ГСП
Private Const L_Insp_StatSK$ = "R"      'Статус RFI СК
Private Const L_Insp_ResltGSP$ = "S"    'Заключение RFI ГСП
Private Const L_Insp_ResltSK$ = "T"     'Заключение RFI СК
Private Const L_Mont_Note$ = "U"        'Примечание по монтажу
'------------------------------------------------------
Private Const L_AOSR_Num$ = "V"         'Номер АОСР
Private Const L_AOSR_Stat$ = "W"        'Статус АОСР
Private Const L_Count_RFI$ = "Y"        'Счётчик для RFI
'------------------------------------------------------
Private Const L_DateMastering$ = "Z"    'Столбец дата освоения
'======================================================

'Настройки таблицы
Private Const start_row As Byte = 3

Private Const color_DelElement& = XlRgbColor.rgbLightSalmon    'Цвет удаленного элемента
Private Const color_NewElement& = XlRgbColor.rgbLightGreen     'Цвет добавленного элемента

'Цвет заголовка таблицы
Private Const color_HeadMasterFile& = XlRgbColor.rgbLightSkyBlue    'Данные из мастер файла
Private Const color_HeadAxis& = XlRgbColor.rgbCadetBlue             'Оси и высотные отметки
Private Const color_HeadNumMountScheme& = XlRgbColor.rgbLightBlue   'Номер монтажной схемы
Private Const color_HeadInspAisNSK& = XlRgbColor.rgbBurlyWood       'Ииспекции аис-нск
Private Const color_HeadAosrAookCntRFI& = XlRgbColor.rgbDarkKhaki   'АОСР, АООК, Счетчик
Private Const color_HeadDateComleted& = XlRgbColor.rgbChartreuse    'Дата освоения

'Переменные класса
Private cnt_columns&
Private mass_data()
Private dict_data As Object
Private dict_insp As Object
Private OtherFunction As New CL_OTHERFUN

'Индексы данных словаря из мастер файла
Public Enum data_table
    idx_title = 1
    idx_sheet
    idx_dp
    idx_mark
    idx_name
    idx_pos
    idx_wght
    idx_note
    idx_date
    
    idx_date_id
    idx_axis_letter
    idx_axis_number
    idx_axis_altitude
    
    idx_diagram
    
    idx_insp_smu
    idx_insp_rfi
    idx_insp_stat_gsp
    idx_insp_stat_sk
    idx_insp_reslt_gsp
    idx_insp_reslt_sk
    idx_insp_note
    
    idx_aosr_num
    idx_aosr_stat
    idx_aook_num
    idx_cntrfi
    
    idx_mastering
End Enum


'Конструктор и деструктор класса
Private Sub Class_Initialize()
    cnt_columns = dict_header.Count    'Определение колличества столбцов в таблице
End Sub
Private Sub Class_Terminate()
    Set dict_data = Nothing
    Set dict_insp = Nothing
    Set OtherFunction = Nothing
End Sub


'Объект рабочего листа с таблицей
Public Property Get SH_MAIN() As Excel.Worksheet
    Set SH_MAIN = Лист1
End Property

'Словарь с данными по элементам
Public Property Set DictData(ByRef dict As Object)
    Set dict_data = dict
End Property

'Словарь с данными инспекций
Public Property Set DictDataInsp(ByRef dict As Object)
    Set dict_insp = dict
End Property
Public Property Get DictDataInsp() As Object
    Set DictDataInsp = dict_insp
End Property

'Массив данных таблицы
Public Property Get MassData()
    Application.StatusBar = "Загрузка данных из таблицы: """ & SH_MAIN.Name & """. Подождите..."
    Debug.Print Application.StatusBar: DoEvents
    MassData = GetMassInTable
    Application.StatusBar = False
End Property


'Область табличных данных
Private Property Get GetRangeTable(ByRef L_Column$) As Range
    With SH_MAIN
        Set GetRangeTable = .Range(L_Column & start_row).CurrentRegion. _
            Offset(start_row - 1).Resize(cnt_row - start_row + 1)
    End With
End Property

'Загрузка табличных данных в массив
Private Function GetMassInTable()
    GetMassInTable = GetRangeTable(L_Sheet).Value
End Function

'Последняя строка данных в таблице
Private Property Get cnt_row&()
    cnt_row = SH_MAIN.Range("A1").CurrentRegion.Rows.Count
    'cnt_row = SH_MAIN.Cells(Rows.Count, idx_sheet).End(xlUp).Row
End Property

'Формирование массива уникальных значений по выбранному столбцу
Public Function GetMassUniqueVal(column_select As col_select)
    
    Dim q&, sKey$, i, dict As Object, RG As Range, AR As Range, L_Column$
    
    If column_select = column_diagram Then L_Column = L_NumMontDiagram
    If column_select = column_aosr Then L_Column = L_AOSR_Num
    
    Set RG = GetRangeTable(L_Column).Columns(L_Column).SpecialCells(xlCellTypeVisible)
    If RG.Row < start_row Then Exit Function
    
    Set dict = CreateObject("Scripting.Dictionary")

    'Отбор уникальных значений
    For Each AR In RG.Areas
        For Each i In AR
            If dict.Count > 19 Then GoTo continue
            If IsError(i.Value) Then GoTo continue
            
            sKey = Trim(i.Value)
            If sKey = "" Then GoTo continue
            
            If Not dict.exists(sKey) Then dict.Add key:=sKey, Item:=Empty
continue:
        Next
    Next
    
    If dict.Count = 0 Then Exit Function
    
    Dim SRT As New CL_SORT
    SRT.SetSortOrder = set_ASC
    GetMassUniqueVal = SRT.Sort(dict.keys())
    Set SRT = Nothing

End Function



'Основная процедура обновления таблицы
Public Sub UpdateTable()
    Dim dict As Object
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    With SH_MAIN
        .Columns.Hidden = False             'Показать все столбцы
        If .FilterMode Then .ShowAllData    'Очистить автофильтр
    End With
    
    'Проверка соответствия данных с мастер файлом
    Set dict = GetAddDictData(dict_discrepancy)
    CreateHeader        'Создание заголовков таблицы
    AddDataInTable dict 'Добавление строк данных в таблицу
    SortTable           'Сортировка таблицы
    FormatingTable      'Форматирование таблицы
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Set dict = Nothing
    
End Sub

'Словарь для проверки на несоответствие поиск по ключу (Марка Наименование Позиция Вес)
Private Function dict_discrepancy() As Object
    Dim iData, iCol, cnt_&, sKey$, dict As Object
    
    Application.StatusBar = "Проверка на несоответствие с мастер файлом. Подождите..."
    Debug.Print Application.StatusBar: DoEvents
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each iCol In dict_data.items
        For Each iData In iCol
            cnt_ = cnt_ + 1
            sKey = iData(idx_title) & vbTab
            sKey = sKey & iData(idx_dp) & vbTab
            sKey = sKey & iData(idx_mark) & vbTab
            sKey = sKey & iData(idx_name) & vbTab
            sKey = sKey & iData(idx_pos) & vbTab
            sKey = sKey & iData(idx_wght)
            If Not dict.exists(sKey) Then dict.Add key:=sKey, Item:=iCol(cnt_)
        Next
        cnt_ = 0
    Next
    
    Set dict_discrepancy = dict
    Set dict = Nothing
    
End Function

'Проверка соответствия данных с мастер файлом
Private Function GetAddDictData(ByRef dict As Object) As Object
    Dim q&, sKey, iRow&

    mass_data = GetMassInTable 'Массив данных из ведомости
    
    'Поиск несоответствий в массиве данных таблицы с пометкой
    ' в словаре (Марка Наименование Позиция Вес)
    With SH_MAIN
        For q = LBound(mass_data) To UBound(mass_data)
            
            sKey = GetKey(q) 'Ключ: Титул, DP, Марка, Имя, Позиция, Вес
            iRow = q + start_row - 1 'Строка в таблице
            
            If dict.exists(sKey) Then
                
                'Поиск инспекции монтажа
                SearchInspData Trim(mass_data(q, idx_insp_smu)), iRow
                
                'Проверка даты монтажа
                If Not mass_data(q, idx_date) = dict(sKey)(idx_date) Then
                    With .Range(L_Date & iRow)
                        .NumberFormat = "dd/mm/yyyy"    'Формат даты монтажа
                        .Value = dict(sKey)(idx_date)   'Значение даты монтажа
                    End With
                End If
                
                'Проверка даты исполнительной
                If IsEmpty(mass_data(q, idx_date_id)) And IsDate(dict(sKey)(idx_date)) Then
                    With .Range(L_DateId & iRow)
                        .NumberFormat = "dd/mm/yyyy"    'Формат даты монтажа
                        .Value = dict(sKey)(idx_date)   'Значение даты монтажа
                    End With
                End If
                
                'Проверка примечания
                If IsEmpty(mass_data(q, idx_note)) Then
                    If Not mass_data(q, idx_note) = dict(sKey)(idx_note) Then
                        .Range(L_Note & iRow).Value = dict(sKey)(idx_note)   'Значение даты монтажа
                    End If
                End If
                
                'Удаление цвета строки
                With .Range("A" & iRow, .Cells(iRow, cnt_columns))
                    .Interior.Color = xlNone
                    .Font.Color = 0
                End With
                
                dict.Remove sKey    'Удаление из словаря записи об найденом элементе
                
            Else
                dict.Add key:=iRow & vbTab & "_removed", Item:=Empty 'Пометка на удаление
            End If
            
        Next
    End With
    
    Set GetAddDictData = dict
    Application.EnableEvents = True

End Function

'Поиск инспекции монтажа
Public Sub SearchInspData(ByRef sKey$, ByRef iRow&)
    Dim q&, mass(), dict As Object
    
    If dict_insp Is Nothing Then Exit Sub
    If dict_insp.Count = 0 Then Exit Sub
    If Not dict_insp.exists(sKey) Then Exit Sub
    
    Set dict = dict_insp(sKey)
    With SH_MAIN
        With .Range(L_Insp_NumAis & iRow)
            If .Value <> dict(1) Then .Value = dict(1)
        End With
        With .Range(L_Insp_StatGSP & iRow)
            If .Value <> dict(2) Then .Value = dict(2)
        End With
        With .Range(L_Insp_StatSK & iRow)
            If .Value <> dict(3) Then .Value = dict(3)
        End With
        With .Range(L_Insp_ResltGSP & iRow)
            If .Value <> dict(4) Then .Value = dict(4)
        End With
        With .Range(L_Insp_ResltSK & iRow)
            If .Value <> dict(5) Then .Value = dict(5)
        End With
    End With
    
    Set dict = Nothing
    
End Sub


'Вывод информации в таблицу в случае расхождения
Private Sub AddDataInTable(ByRef dict As Object)
    Dim iKey, iRow&
    
    If dict.Count = 0 Then Exit Sub
    
    With SH_MAIN
        
        For Each iKey In dict
            If InStr(1, iKey, "_removed", vbBinaryCompare) = 0 Then
                'Добавление новых элементов
                iRow = cnt_row + 1
                
                Application.StatusBar = "Добавление новых элементов: " & Replace(iKey, vbTab, "_")
                Debug.Print Application.StatusBar: DoEvents
                
                'Форматирвоание некоторых вставляемых элементов
                .Range(L_Date & iRow).NumberFormat = "dd/mm/yyyy"     'Формат даты монтажа
                .Range(L_DateId & iRow).NumberFormat = "dd/mm/yyyy"     'Формат даты монтажа
                .Range(L_AxesNumber & iRow).NumberFormat = "@"
                .Range(L_AltitudeMark & iRow).NumberFormat = "@"
                
                'Вставка значений
                .Range(L_Title & iRow).Value = dict(iKey)(idx_title)  'Титул
                .Range(L_Sheet & iRow).Value = dict(iKey)(idx_sheet)  'Лист
                .Range(L_DP & iRow).Value = dict(iKey)(idx_dp)        'DP
                .Range(L_Mark & iRow).Value = dict(iKey)(idx_mark)    'Марка
                .Range(L_Name & iRow).Value = dict(iKey)(idx_name)    'Наименование
                .Range(L_Pos & iRow).Value = dict(iKey)(idx_pos)      'Позиция
                .Range(L_Wght & iRow).Value = dict(iKey)(idx_wght)    'Вес
                .Range(L_Note & iRow).Value = dict(iKey)(idx_note)    'Примечание
                .Range(L_Date & iRow).Value = dict(iKey)(idx_date)    'Дата монтажа
                .Range(L_DateId & iRow).Value = dict(iKey)(idx_date)  'Дата монтажа
                
                'Отметка строки цветом для новых элементов
                .Range("A" & iRow, .Cells(iRow, cnt_columns)).Interior.Color = color_NewElement
                
            Else
                'Отметка строки цветом для удаленных элементов
                iRow = Split(iKey, vbTab)(0)
                Application.StatusBar = "Пометка элемента на удаление строка №: " & iRow
                .Range("A" & iRow, .Cells(iRow, cnt_columns)).Interior.Color = color_DelElement
            End If
        Next
        
    End With
    
    Application.StatusBar = False
    
End Sub

'Ключ: Титул, DP, Марка, Имя, Позиция, Вес
Private Property Get GetKey$(ByRef q&)
    Dim sKey$
    sKey = mass_data(q, idx_title) & vbTab
    sKey = sKey & mass_data(q, idx_dp) & vbTab
    sKey = sKey & mass_data(q, idx_mark) & vbTab
    sKey = sKey & mass_data(q, idx_name) & vbTab
    sKey = sKey & mass_data(q, idx_pos) & vbTab
    sKey = sKey & mass_data(q, idx_wght)
    GetKey = sKey
End Property

'Сортировка таблицы по трем полям: Лист, DP, Наименование материала
Private Sub SortTable()
    
    Application.StatusBar = "Сортировка таблицы. Подождите..."
    Debug.Print Application.StatusBar: DoEvents
    
    With SH_MAIN
    
        .Sort.SortFields.Clear  'Очистка полей сортировки
        
        '1. Поле сортировки: Наименование листа таблицы в мастер файле
        .Sort.SortFields.Add key:=Range(L_Sheet & start_row & ":" & L_Sheet & cnt_row), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
        '2. Поле сортировки: DP
        .Sort.SortFields.Add key:=Range(L_DP & start_row & ":" & L_DP & cnt_row), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
        '3. Поле сортировки: Марка
        .Sort.SortFields.Add key:=Range(L_Mark & start_row & ":" & L_Mark & cnt_row), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            
        '4. Поле сортировки: Наименование
        .Sort.SortFields.Add key:=Range(L_Name & start_row & ":" & L_Name & cnt_row), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            
        '5. Поле сортировки: Позиция
        .Sort.SortFields.Add key:=Range(L_Pos & start_row & ":" & L_Pos & cnt_row), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
        With .Sort 'Применение сортировки
            .SetRange GetRangeTable(L_Sheet) 'Область сортировки
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    
    Application.StatusBar = False
    
End Sub

'Формирование названий столбцов
Private Function dict_header() As Object
    Set dict_header = CreateObject("Scripting.Dictionary")
    With dict_header
        .Add key:=.Count + 1, Item:="Титул"
        .Add key:=.Count + 1, Item:="Наименование листа таблицы в мастер файле"
        .Add key:=.Count + 1, Item:="DP"
        .Add key:=.Count + 1, Item:="Марка"
        .Add key:=.Count + 1, Item:="Наименование"
        .Add key:=.Count + 1, Item:="Пози-ция"
        
        .Add key:=.Count + 1, Item:="Масса" & Chr(10) & "[кг]"
        .Add key:=.Count + 1, Item:="Примечание"
        
        .Add key:=.Count + 1, Item:="Дата" & Chr(10) & "монтажа"
        .Add key:=.Count + 1, Item:="Дата" & Chr(10) & "испол-ной"
        
        .Add key:=.Count + 1, Item:="Ось буква"
        .Add key:=.Count + 1, Item:="Ось цифра"
        .Add key:=.Count + 1, Item:="Высотная отметка"
        
        .Add key:=.Count + 1, Item:="Номер монтажной схемы"
        
        .Add key:=.Count + 1, Item:="Исх. Номер на RFI"
        .Add key:=.Count + 1, Item:="№ RFI на монтаж"
        .Add key:=.Count + 1, Item:="Статус RFI ГСП"
        .Add key:=.Count + 1, Item:="Статус RFI СК"
        .Add key:=.Count + 1, Item:="Заключение ГСП"
        .Add key:=.Count + 1, Item:="Заключение СК"
        
        .Add key:=.Count + 1, Item:="Примечание по монтажу"
        
        .Add key:=.Count + 1, Item:="№ АОСР"
        .Add key:=.Count + 1, Item:="Статус АОСР"
        .Add key:=.Count + 1, Item:="№ АООК"
        .Add key:=.Count + 1, Item:="Счётчик для RFI"
        .Add key:=.Count + 1, Item:="Дата освоения"
    End With
    
End Function

'Создание  заголовков таблицы
Private Sub CreateHeader()
    Dim dict As Object
    
    Application.StatusBar = "Обновление заголовков таблицы. Подождите..."
    Debug.Print Application.StatusBar: DoEvents
    
    'Ширина столбцов
    OtherFunction.ColumnsWidth SH_MAIN, 1, 9, 31, 10, 13, 25, 6, 10, 22, 12, 12, 8, 8, 16, 40, 23, 23, 23, 23, 23, 23, 40, 23, 23, 32, 23, 12
    
    Set dict = dict_header
    With SH_MAIN
        
        .Rows(1).RowHeight = 25.5
        .Rows(2).RowHeight = 15
            
         'Форматирование заголовков
        OtherFunction.FormatCells .Range("A1", .Cells(2, cnt_columns)), xlHAlignCenter, xlVAlignCenter, Bold:=True, Wrap:=True
        
        'Раскраска столбцов заголовка
        .Range("A1:I2").Interior.Color = color_HeadMasterFile
        .Range("J1:M2").Interior.Color = color_HeadAxis
        .Range("N1:N2").Interior.Color = color_HeadNumMountScheme
        .Range("O1:U2").Interior.Color = color_HeadInspAisNSK
        .Range("V1:Y2").Interior.Color = color_HeadAosrAookCntRFI
        .Range("Z1:Z2").Interior.Color = color_HeadDateComleted
        
        'Вставка наименований заголовков
        .Range("A1", .Cells(1, cnt_columns)).Value = dict.items
        
        'Вставка нумерации столбцов
        .Range("A2", .Cells(2, cnt_columns)).Value = dict.keys
        
        'Установка автофильтра
        If Not .AutoFilterMode Then .Range("A2", .Cells(2, cnt_columns)).AutoFilter
        
        'Установка флажка Собятия книги
        SetCheckBoxesEvents
        
    End With
    
    Application.StatusBar = False

End Sub

'Форматирвоание таблицы
Private Sub FormatingTable()
    
    Application.StatusBar = "Форматирование таблицы. Подождите..."
    Debug.Print Application.StatusBar: DoEvents
    
    With GetRangeTable(L_Sheet)
        
        'Формат таблицы по умолчанию
        OtherFunction.FormatCells .Offset, xlHAlignCenter, xlVAlignCenter
        
        'Форматирование отдельных столбцов по форматам данных
        .Columns(L_AxesNumber).NumberFormat = "@"
        .Columns(L_AltitudeMark).NumberFormat = "@"
        .Columns(L_Date).NumberFormat = "dd/mm/yyyy"
        .Columns(L_DateId).NumberFormat = "dd/mm/yyyy"
        
        .Columns(L_Title).Font.Italic = True
        .Columns(L_Sheet).Font.Italic = True
        .Columns(L_DP).Font.Italic = True
        .Columns(L_Mark).Font.Italic = True
        .Columns(L_Name).Font.Italic = True
        .Columns(L_Pos).Font.Italic = True
        .Columns(L_Wght).Font.Italic = True
        .Columns(L_Note).Font.Italic = True
        .Columns(L_Date).Font.Italic = True
        
        'Перенос текста в столбце примечание из мастер файла
        .Columns(L_Mark).WrapText = True
        '.Columns(L_Note).WrapText = True
        .Columns(L_Note).HorizontalAlignment = xlHAlignLeft
        .Columns(L_DateMastering).WrapText = True
        
    End With
    
    OtherFunction.BoldersCells SH_MAIN.Range(L_Title & start_row).CurrentRegion, xlThin
    
    Application.StatusBar = False
    
End Sub

'Установка флажка управления событиями книги
Private Sub SetCheckBoxesEvents()
    Dim i, X!, Y!
    
    With SH_MAIN
        
        'Удаление всех чек боксов
        For Each i In .CheckBoxes
            i.Delete
        Next
        
        'Координаты вставки флажка
        With .Range(L_DateMastering & 1).Offset(columnoffset:=1)
            X = .Left
            Y = .Top
        End With
        
        'Вставка флажка
        With .CheckBoxes.Add(X, Y, 80, 25)
            .Name = "set_events"
            .Caption = "События книги"
            .OnAction = "'SetEnableEvents'"
            .Value = True
        End With
    End With
End Sub
