Attribute VB_Name = "Update_SMU_MK_RFI"
Option Explicit

Sub Update_SMU_MK_RFI()

    ' --- ОБЪЯВЛЕНИЯ ---
    Dim wbDest As Workbook
    Dim wsDest As Worksheet
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim srcFilePath As Variant
    Dim dictO As Object, dictP As Object
    Dim key As String
    Dim lastRowDest As Long, lastColDest As Long
    Dim lastRowSrc As Long, lastColSrc As Long
    Dim arrDest As Variant, arrSrc As Variant
    Dim i As Long
    Dim colA As Long, colC As Long, colD As Long, colF As Long, colG As Long, colO As Long, colP As Long
    Dim srcOValue As Variant, srcPValue As Variant
    Dim cellValue As String
    Dim startTime As Double
    Dim totalRows As Long, processed As Long
    Dim destSheetName As String, srcSheetName As String

    ' Настройки
    destSheetName = "База данных по элементам"
    srcSheetName = "База данных по элементам"
    Const PROGRESS_STEP As Long = 1000 ' Обновлять статус каждые N строк

    ' Инициализация
    startTime = Timer
    Set wbDest = ActiveWorkbook
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler

    ' --- 1. ПРОВЕРКА ЛИСТА В ПРИЁМНИКЕ ---
    On Error Resume Next
    Set wsDest = wbDest.Worksheets(destSheetName)
    On Error GoTo ErrorHandler
    If wsDest Is Nothing Then
        MsgBox "В активной книге (приёмник) отсутствует лист '" & destSheetName & "'.", vbCritical, "Ошибка"
        GoTo Cleanup
    End If
    Debug.Print "Приёмник: " & wbDest.Name & " / " & wsDest.Name

    ' --- 2. ВЫБОР ФАЙЛА ИСТОЧНИКА ---
    srcFilePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xls; *.xlsx; *.xlsm; *.xlsb), *.xls; *.xlsx; *.xlsm; *.xlsb", _
        Title:="Выберите файл Ведомость элементов МК ТСБ, который был скопирован от СМУ")
    If srcFilePath = False Then
        MsgBox "Файл не выбран. Макрос остановлен.", vbExclamation, "Отмена"
        GoTo Cleanup
    End If
    Debug.Print "Выбран источник: " & srcFilePath

    ' --- 3. ОТКРЫТИЕ ИСТОЧНИКА ---
    ' Открываем в режиме ReadOnly, не обновляя связи
    Set wbSrc = Workbooks.Open(Filename:=srcFilePath, UpdateLinks:=False, ReadOnly:=True)
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(srcSheetName)
    On Error GoTo ErrorHandler
    If wsSrc Is Nothing Then
        MsgBox "В файле-источнике отсутствует лист '" & srcSheetName & "'.", vbCritical, "Ошибка"
        wbSrc.Close SaveChanges:=False
        GoTo Cleanup
    End If
    Debug.Print "Источник: " & wbSrc.Name & " / " & wsSrc.Name

    ' --- 4. ОПРЕДЕЛЕНИЕ ГРАНИЦ ДАННЫХ В ПРИЁМНИКЕ ---
    With wsDest
        ' Последняя строка данных по столбцу A
        lastRowDest = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRowDest < 3 Then
            MsgBox "На листе приёмника нет данных (начиная с 3 строки).", vbExclamation, "Нет данных"
            wbSrc.Close SaveChanges:=False
            GoTo Cleanup
        End If
        ' Последний столбец по 3-й строке (первая строка данных)
        lastColDest = .Cells(3, .Columns.Count).End(xlToLeft).Column
        ' Гарантируем, что захватим столбцы до P
        If lastColDest < 16 Then lastColDest = 16 ' P = 16-й столбец
        ' Читаем массив данных с 3 строки
        arrDest = .Range(.Cells(3, 1), .Cells(lastRowDest, lastColDest)).Value
    End With
    totalRows = UBound(arrDest, 1)
    Debug.Print "Приёмник: строк данных = " & totalRows & ", столбцов = " & lastColDest

    ' Определяем индексы столбцов в массиве (относительные, т.к. массив начинается с 1)
    colA = 1                ' A
    colC = 3                ' C
    colD = 4                ' D
    colF = 6                ' F
    colG = 7                ' G
    colO = 15               ' O (15-й столбец)
    colP = 16               ' P (16-й столбец)

    ' --- 5. ПОСТРОЕНИЕ СЛОВАРЯ ИСТОЧНИКА ДЛЯ СТОЛБЦА O ---
    Set dictO = CreateObject("Scripting.Dictionary")
    dictO.CompareMode = vbTextCompare ' Игнорировать регистр? По желанию. Если регистр важен, уберите строку.
    ' Получаем данные источника
    With wsSrc
        lastRowSrc = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRowSrc < 3 Then
            MsgBox "На листе источника нет данных (начиная с 3 строки).", vbExclamation, "Нет данных"
            wbSrc.Close SaveChanges:=False
            GoTo Cleanup
        End If
        lastColSrc = .Cells(3, .Columns.Count).End(xlToLeft).Column
        If lastColSrc < 16 Then lastColSrc = 16
        arrSrc = .Range(.Cells(3, 1), .Cells(lastRowSrc, lastColSrc)).Value
    End With
    Debug.Print "Источник: строк данных = " & UBound(arrSrc, 1) & ", столбцов = " & lastColSrc

    ' Заполняем словарь для O
    Application.StatusBar = "Построение словаря источника для столбца O..."
    For i = 1 To UBound(arrSrc, 1)
        ' Формируем ключ: A|C|D|F|G, удаляем пробелы
        key = Join(Array( _
            CStr(arrSrc(i, colA)), _
            CStr(arrSrc(i, colC)), _
            CStr(arrSrc(i, colD)), _
            CStr(arrSrc(i, colF)), _
            CStr(arrSrc(i, colG))), "|")
        key = Replace(key, " ", "") ' удаляем все пробелы

        srcOValue = arrSrc(i, colO)
        ' Проверяем, что O не пусто и не состоит только из пробелов
        If Not IsEmpty(srcOValue) Then
            If Len(Trim(CStr(srcOValue))) > 0 Then
                dictO(key) = srcOValue ' последнее значение перезаписывает предыдущее
            End If
        End If

        If i Mod PROGRESS_STEP = 0 Then
            Application.StatusBar = "Обработка источника O: " & i & " из " & UBound(arrSrc, 1)
            DoEvents
        End If
    Next i
    Debug.Print "Словарь O построен. Уникальных ключей: " & dictO.Count

    ' --- 6. ОБРАБОТКА ПРИЁМНИКА: СТОЛБЕЦ O ---
    Application.StatusBar = "Заполнение столбца O в приёмнике..."
    processed = 0
    For i = 1 To totalRows
        cellValue = CStr(arrDest(i, colO))
        ' Если ячейка пустая или содержит только пробелы
        If Len(Trim(cellValue)) = 0 Then
            ' Формируем ключ из приёмника
            key = Join(Array( _
                CStr(arrDest(i, colA)), _
                CStr(arrDest(i, colC)), _
                CStr(arrDest(i, colD)), _
                CStr(arrDest(i, colF)), _
                CStr(arrDest(i, colG))), "|")
            key = Replace(key, " ", "")
            If dictO.exists(key) Then
                arrDest(i, colO) = dictO(key)
                processed = processed + 1
            Else
                ' Оставляем пустым, но на всякий случай очищаем пробелы
                arrDest(i, colO) = ""
            End If
        End If

        If i Mod PROGRESS_STEP = 0 Then
            Application.StatusBar = "Приёмник O: обработано " & i & " из " & totalRows
            DoEvents
        End If
    Next i
    Debug.Print "Столбец O: обновлено ячеек = " & processed

    ' --- 7. ПОСТРОЕНИЕ СЛОВАРЯ ИСТОЧНИКА ДЛЯ СТОЛБЦА P ---
    Set dictP = CreateObject("Scripting.Dictionary")
    dictP.CompareMode = vbTextCompare
    Application.StatusBar = "Построение словаря источника для столбца P..."
    For i = 1 To UBound(arrSrc, 1)
        key = Join(Array( _
            CStr(arrSrc(i, colA)), _
            CStr(arrSrc(i, colC)), _
            CStr(arrSrc(i, colD)), _
            CStr(arrSrc(i, colF)), _
            CStr(arrSrc(i, colG))), "|")
        key = Replace(key, " ", "")

        srcPValue = arrSrc(i, colP)
        If Not IsEmpty(srcPValue) Then
            If Len(Trim(CStr(srcPValue))) > 0 Then
                dictP(key) = srcPValue
            End If
        End If

        If i Mod PROGRESS_STEP = 0 Then
            Application.StatusBar = "Обработка источника P: " & i & " из " & UBound(arrSrc, 1)
            DoEvents
        End If
    Next i
    Debug.Print "Словарь P построен. Уникальных ключей: " & dictP.Count

    ' --- 8. ОБРАБОТКА ПРИЁМНИКА: СТОЛБЕЦ P ---
    Application.StatusBar = "Заполнение столбца P в приёмнике..."
    processed = 0
    For i = 1 To totalRows
        cellValue = CStr(arrDest(i, colP))
        If Len(Trim(cellValue)) = 0 Then
            key = Join(Array( _
                CStr(arrDest(i, colA)), _
                CStr(arrDest(i, colC)), _
                CStr(arrDest(i, colD)), _
                CStr(arrDest(i, colF)), _
                CStr(arrDest(i, colG))), "|")
            key = Replace(key, " ", "")
            If dictP.exists(key) Then
                arrDest(i, colP) = dictP(key)
                processed = processed + 1
            Else
                arrDest(i, colP) = ""
            End If
        End If

        If i Mod PROGRESS_STEP = 0 Then
            Application.StatusBar = "Приёмник P: обработано " & i & " из " & totalRows
            DoEvents
        End If
    Next i
    Debug.Print "Столбец P: обновлено ячеек = " & processed

    ' --- 9. ВЫГРУЗКА МАССИВА ОБРАТНО НА ЛИСТ ПРИЁМНИКА ---
    Application.StatusBar = "Запись данных обратно на лист приёмника..."
    wsDest.Range("A3").Resize(totalRows, lastColDest).Value = arrDest
    Debug.Print "Данные записаны."

    ' --- 10. ЗАВЕРШЕНИЕ ---
    Dim elapsed As Double
    elapsed = Round(Timer - startTime, 2)
    MsgBox "Обработка завершена!" & vbCrLf & _
           "Время выполнения: " & elapsed & " сек.", vbInformation, "Готово"
    Debug.Print "Время выполнения: " & elapsed & " сек."

Cleanup:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Set dictO = Nothing
    Set dictP = Nothing
    Set wsDest = Nothing
    Set wsSrc = Nothing
    Set wbSrc = Nothing
    Set wbDest = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Произошла ошибка: " & Err.Description & vbCrLf & _
           "Номер ошибки: " & Err.Number, vbCritical, "Ошибка выполнения"
    Debug.Print "ОШИБКА: " & Err.Description & " (" & Err.Number & ")"
    Resume Cleanup
End Sub
