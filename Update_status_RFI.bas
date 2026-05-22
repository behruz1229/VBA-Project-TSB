Attribute VB_Name = "Update_status_RFI"
Option Explicit

Sub Update_status_RFI()
    '-----------------------------------------------------------------------
    ' ОПТИМИЗИРОВАННАЯ ВЕРСИЯ (ИСПРАВЛЕНО ЧТЕНИЕ ИСТОЧНИКА)
    ' Макрос обновляет столбцы Q и R в активной книге (лист "База данных по элементам")
    ' на основе данных из выбранного файла (лист "Инспекции").
    ' Сопоставление: столбец P (приёмник) <-> столбец C (источник)
    ' Обновление: Q <- AG, R <- AH
    '-----------------------------------------------------------------------
    
    Dim wbReceiver As Workbook
    Dim wsReceiver As Worksheet
    Dim lastRowReceiver As Long
    Dim recArr As Variant
    Dim i As Long
    Dim dict As Object
    Dim sourceFilePath As Variant
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim lastRowSource As Long
    Dim srcArr As Variant
    Dim wasAlreadyOpen As Boolean
    Dim updateCountQ As Long
    Dim updateCountR As Long
    Dim processedCount As Long
    Dim recalcState As XlCalculation
    Dim screenUpdState As Boolean
    Dim eventsState As Boolean
    Dim startTime As Double
    Dim elapsed As Double
    Dim statusInterval As Long
    
    startTime = Timer
    
    ' --- 1. Проверка активной книги ---
    Set wbReceiver = ActiveWorkbook
    If wbReceiver.Name = "PERSONAL.XLSB" Then
        MsgBox "Пожалуйста, активируйте нужную книгу (приёмник) перед запуском макроса.", vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    ' --- 2. Настройка листа-приёмника ---
    On Error Resume Next
    Set wsReceiver = wbReceiver.Worksheets("База данных по элементам")
    On Error GoTo 0
    If wsReceiver Is Nothing Then
        MsgBox "В активной книге не найден лист 'База данных по элементам'.", vbCritical, "Ошибка"
        Exit Sub
    End If
    
    lastRowReceiver = wsReceiver.Cells(wsReceiver.Rows.Count, 1).End(xlUp).Row
    If lastRowReceiver < 3 Then
        MsgBox "В листе 'База данных по элементам' нет данных (строки начинаются с 3).", vbInformation, "Выход"
        Exit Sub
    End If
    
    ' --- 3. Чтение приёмника в массив (P, Q, R) ---
    recArr = wsReceiver.Range("P3:R" & lastRowReceiver).Value2
    processedCount = UBound(recArr, 1)
    
    If processedCount = 0 Then
        MsgBox "Нет данных для обработки.", vbInformation, "Выход"
        Exit Sub
    End If
    
    ' --- 4. Выбор файла-источника ---
    sourceFilePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files,*.xls*;*.xlsx;*.xlsm;*.xlsb", _
        Title:="Выберите инспекции, скачанный из АИС НСК", _
        MultiSelect:=False)
    If sourceFilePath = False Then
        MsgBox "Операция отменена пользователем.", vbInformation, "Отмена"
        Exit Sub
    End If
    
    ' --- 5. Открытие источника ---
    wasAlreadyOpen = False
    On Error Resume Next
    Set wbSource = Workbooks(Dir(sourceFilePath))
    If wbSource Is Nothing Then
        Set wbSource = Workbooks.Open(sourceFilePath, ReadOnly:=True, UpdateLinks:=False)
        If wbSource Is Nothing Then
            MsgBox "Не удалось открыть файл-источник: " & sourceFilePath, vbCritical, "Ошибка"
            Exit Sub
        End If
    Else
        If wbSource.FullName <> sourceFilePath Then
            Set wbSource = Workbooks.Open(sourceFilePath, ReadOnly:=True, UpdateLinks:=False)
            If wbSource Is Nothing Then
                MsgBox "Не удалось открыть файл-источник: " & sourceFilePath, vbCritical, "Ошибка"
                Exit Sub
            End If
        Else
            wasAlreadyOpen = True
        End If
    End If
    On Error GoTo 0
    
    ' --- 6. Проверка листа "Инспекции" ---
    Set wsSource = Nothing
    On Error Resume Next
    Set wsSource = wbSource.Worksheets("Инспекции")
    On Error GoTo 0
    If wsSource Is Nothing Then
        MsgBox "В файле-источнике не найден лист 'Инспекции'.", vbCritical, "Ошибка"
        If Not wasAlreadyOpen Then wbSource.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' --- 7. Чтение источника (ИСПРАВЛЕННЫЙ БЛОК) ---
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row  ' определяем по столбцу C
    If lastRowSource < 2 Then
        MsgBox "В листе 'Инспекции' нет данных (строки начинаются с 2).", vbInformation, "Выход"
        If Not wasAlreadyOpen Then wbSource.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Читаем диапазон C2:AH & lastRowSource в массив
    ' Столбец C = индекс 1, AG = индекс 31, AH = индекс 32
    srcArr = wsSource.Range("C2:AH" & lastRowSource).Value2
    
    ' --- 8. Закрытие источника ---
    If Not wasAlreadyOpen Then
        wbSource.Close SaveChanges:=False
    End If
    Set wbSource = Nothing
    Set wsSource = Nothing
    
    ' --- 9. Построение словаря ---
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim j As Long
    Dim srcKey As String
    For j = 1 To UBound(srcArr, 1)
        srcKey = Trim(srcArr(j, 1) & "")
        If srcKey <> "" Then
            dict(srcKey) = Array(srcArr(j, 31), srcArr(j, 32))  ' AG и AH
        End If
    Next j
    
    Erase srcArr  ' освобождаем память
    
    ' --- 10. Подготовка к обновлению ---
    screenUpdState = Application.ScreenUpdating
    recalcState = Application.Calculation
    eventsState = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    updateCountQ = 0
    updateCountR = 0
    statusInterval = 5000
    
    Debug.Print String(80, "=")
    Debug.Print "Запуск обновления (исправленная версия): " & Now
    Debug.Print "Всего строк в приёмнике: " & processedCount
    Debug.Print "Размер словаря источника: " & dict.Count
    
    ' --- 11. Обработка в памяти ---
    Dim recKey As String
    Dim srcVals As Variant
    Dim curQ As Variant, curR As Variant
    Dim wasUpdatedQ As Boolean, wasUpdatedR As Boolean
    
    For i = 1 To processedCount
        recKey = Trim(recArr(i, 1) & "")
        If recKey <> "" Then
            If dict.exists(recKey) Then
                srcVals = dict(recKey)
                
                wasUpdatedQ = False
                wasUpdatedR = False
                
                curQ = recArr(i, 2)
                If IsEmpty(curQ) Or Trim(curQ & "") = "" Then
                    recArr(i, 2) = srcVals(0)
                    updateCountQ = updateCountQ + 1
                    wasUpdatedQ = True
                End If
                
                curR = recArr(i, 3)
                If IsEmpty(curR) Or Trim(curR & "") = "" Then
                    recArr(i, 3) = srcVals(1)
                    updateCountR = updateCountR + 1
                    wasUpdatedR = True
                End If
                
                If wasUpdatedQ Or wasUpdatedR Then
                    Debug.Print "Строка " & (i + 2) & " (ключ: " & recKey & "): " & _
                                IIf(wasUpdatedQ, "Q->" & srcVals(0), "") & _
                                IIf(wasUpdatedQ And wasUpdatedR, "; ", "") & _
                                IIf(wasUpdatedR, "R->" & srcVals(1), "")
                End If
            End If
        End If
        
        If i Mod statusInterval = 0 Then
            Application.StatusBar = "Обработка строки " & i & " из " & processedCount & _
                                    " (обновлено Q: " & updateCountQ & ", R: " & updateCountR & ")"
            DoEvents
        End If
    Next i
    
    ' --- 12. Запись изменений на лист ---
    wsReceiver.Range("Q3:Q" & lastRowReceiver).Value2 = Application.Index(recArr, 0, 2)
    wsReceiver.Range("R3:R" & lastRowReceiver).Value2 = Application.Index(recArr, 0, 3)
    
    ' --- 13. Восстановление настроек ---
    Application.StatusBar = False
    Application.ScreenUpdating = screenUpdState
    Application.Calculation = recalcState
    Application.EnableEvents = eventsState
    
    elapsed = Timer - startTime
    
    Debug.Print "Обработка завершена за " & Format(elapsed, "0.0") & " сек."
    Debug.Print "Обновлено ячеек Q: " & updateCountQ
    Debug.Print "Обновлено ячеек R: " & updateCountR
    Debug.Print String(80, "=")
    
    MsgBox "Обновление завершено." & vbCrLf & _
           "Обработано строк: " & processedCount & vbCrLf & _
           "Обновлено ячеек Q: " & updateCountQ & vbCrLf & _
           "Обновлено ячеек R: " & updateCountR & vbCrLf & _
           "Время выполнения: " & Format(elapsed, "0.0") & " сек.", _
           vbInformation, "Результат"
    
    Set dict = Nothing
    Erase recArr
End Sub
