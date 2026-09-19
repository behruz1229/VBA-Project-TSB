Option Explicit

Sub Update_SMU_MK_RFI()
    ' --- ОБЪЯВЛЕНИЯ И КОНСТАНТЫ ---
    Const DEFAULT_SRC_PATH As String = "~\Ведомость элементов МК ТСБ-СМУ.xlsb" ' <--- путь к файлу
    Const DEST_SHEET_NAME As String = "База данных по элементам"
    Const SRC_SHEET_NAME As String = "База данных по элементам"
    
    Dim wbDest As Workbook, wsDest As Worksheet
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim srcFilePath As Variant
    Dim dict As Object
    Dim key As String
    Dim lastRowDest As Long, lastRowSrc As Long
    Dim arrDest As Variant, arrSrc As Variant
    Dim i As Long, srcRows As Long, destRows As Long
    Dim startTime As Double, phaseStart As Double
    Dim elapsed As Double, remaining As Double
    Dim lastPct As Long, currentPct As Long, phaseNum As Long
    Dim phaseDesc As String
    Dim vA As Variant, vC As Variant, vD As Variant, vF As Variant, vO As Variant
    Dim existingVal As String
    
    ' --- ИНИЦИАЛИЗАЦИЯ ---
    startTime = Timer
    Set wbDest = ActiveWorkbook
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = False
    
    On Error GoTo ErrorHandler
    
    Debug.Print String(80, "=")
    Debug.Print "Запуск Update_SMU_MK_RFI: " & Now
    
    ' --- 1. ПРОВЕРКА ЛИСТА В ПРИЁМНИКЕ ---
    phaseNum = 0
    phaseDesc = "Проверка приёмника"
    UpdateStatusBar phaseNum, phaseDesc, 0, 0, 0, -1
    
    On Error Resume Next
    Set wsDest = wbDest.Worksheets(DEST_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If wsDest Is Nothing Then
        MsgBox "В активной книге отсутствует лист '" & DEST_SHEET_NAME & "'.", vbCritical, "Ошибка"
        GoTo Cleanup
    End If
    
    ' --- 2. ПОДГОТОВКА ДАННЫХ ПРИЁМНИКА ---
    phaseNum = 1
    phaseDesc = "Чтение приёмника"
    phaseStart = Timer
    
    With wsDest
        lastRowDest = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRowDest < 3 Then
            MsgBox "На листе приёмника нет данных.", vbExclamation, "Нет данных"
            GoTo Cleanup
        End If
        
        ' Очищаем ТОЛЬКО столбец O (столбец P не трогаем)
        .Range(.Cells(3, 15), .Cells(lastRowDest, 15)).ClearContents
        
        ' Читаем массив строго от A (1) до O (15)
        arrDest = .Range(.Cells(3, 1), .Cells(lastRowDest, 15)).Value2
    End With
    
    destRows = UBound(arrDest, 1)
    Debug.Print "Приёмник прочитан. Строк: " & destRows & ". Время: " & Format(Timer - phaseStart, "0.00") & " сек."

    ' --- 3. ПОИСК И ОТКРЫТИЕ ФАЙЛА-ИСТОЧНИКА ---
    phaseDesc = "Открытие источника"
    
    If Dir(DEFAULT_SRC_PATH) <> "" Then
        srcFilePath = DEFAULT_SRC_PATH
    Else
        If MsgBox("Файл не найден по пути:" & vbCrLf & DEFAULT_SRC_PATH & vbCrLf & vbCrLf & "Выбрать файл вручную?", _
                  vbYesNo + vbExclamation, "Файл не найден") = vbYes Then
            srcFilePath = Application.GetOpenFilename("Excel Files (*.xlsb; *.xlsx; *.xlsm; *.xls), *.xlsb; *.xlsx; *.xlsm; *.xls")
            If VarType(srcFilePath) = vbBoolean Then GoTo Cleanup
        Else
            GoTo Cleanup
        End If
    End If

    Do
        Set wbSrc = Workbooks.Open(fileName:=srcFilePath, UpdateLinks:=False, ReadOnly:=True)
        On Error Resume Next
        Set wsSrc = wbSrc.Worksheets(SRC_SHEET_NAME)
        On Error GoTo ErrorHandler
        
        If wsSrc Is Nothing Then
            wbSrc.Close SaveChanges:=False
            Set wbSrc = Nothing
            If MsgBox("В файле отсутствует лист '" & SRC_SHEET_NAME & "'. Выбрать другой?", _
                      vbYesNo + vbCritical, "Лист не найден") = vbYes Then
                srcFilePath = Application.GetOpenFilename("Excel Files (*.xlsb; *.xlsx; *.xlsm; *.xls), *.xlsb; *.xlsx; *.xlsm; *.xls")
                If VarType(srcFilePath) = vbBoolean Then GoTo Cleanup
            Else
                GoTo Cleanup
            End If
        Else
            Exit Do
        End If
    Loop

    ' --- 4. ЧТЕНИЕ ИСТОЧНИКА И ПОСТРОЕНИЕ СЛОВАРЯ (ВАРИАНТ 1) ---
    phaseNum = 1
    phaseDesc = "Построение словаря"
    phaseStart = Timer
    lastPct = 0
    
    With wsSrc
        lastRowSrc = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRowSrc < 3 Then
            MsgBox "На листе источника нет данных.", vbExclamation, "Нет данных"
            GoTo Cleanup
        End If
        arrSrc = .Range(.Cells(3, 1), .Cells(lastRowSrc, 15)).Value2
    End With
    
    srcRows = UBound(arrSrc, 1)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' ВАЖНО: vbBinaryCompare для стабильности при больших объёмах
    dict.CompareMode = vbBinaryCompare
    
    For i = 1 To srcRows
        vA = arrSrc(i, 1): If IsError(vA) Then vA = ""
        vC = arrSrc(i, 3): If IsError(vC) Then vC = ""
        vD = arrSrc(i, 4): If IsError(vD) Then vD = ""
        vF = arrSrc(i, 6): If IsError(vF) Then vF = ""
        vO = arrSrc(i, 15): If IsError(vO) Then vO = ""
        
        ' Сборка ключа (A, C, D, F) - БЕЗ G
        key = CStr(vA & "") & "|" & CStr(vC & "") & "|" & CStr(vD & "") & "|" & CStr(vF & "")
        
        ' Очистка от невидимых символов
        key = Replace(key, " ", "")
        key = Replace(key, ChrW(160), "")
        key = Replace(key, vbCr, "")
        key = Replace(key, vbLf, "")
        key = Replace(key, vbTab, "")
        
        ' Добавляем в словарь ТОЛЬКО если в столбце O есть данные
        If Len(Trim(CStr(vO & ""))) > 0 Then
            ' === ВАРИАНТ 1: НЕ ПЕРЕЗАПИСЫВАТЬ НЕПУСТОЕ ЗНАЧЕНИЕ ===
            If dict.Exists(key) Then
                ' Ключ уже есть - проверяем, не пустое ли там значение
                existingVal = CStr(dict(key))
                If Len(Trim(existingVal)) = 0 Then
                    ' Было пустое - заменяем на новое непустое
                    dict(key) = vO
                End If
                ' Если уже есть непустое значение - НЕ перезаписываем (оставляем первое)
            Else
                ' Ключа ещё нет - добавляем
                dict(key) = vO
            End If
            ' =======================================================
        End If
        
        ' Обновление прогресс-бара
        currentPct = Int(i / srcRows * 100)
        If currentPct > lastPct Then
            lastPct = currentPct
            elapsed = Timer - phaseStart
            If currentPct > 0 Then
                remaining = (elapsed / currentPct) * (100 - currentPct)
            Else
                remaining = 0
            End If
            UpdateStatusBar phaseNum, phaseDesc, currentPct, i, srcRows, remaining
        End If
    Next i
    
    Debug.Print "Словарь построен. Уникальных ключей: " & dict.Count & ". Время: " & Format(Timer - phaseStart, "0.00") & " сек."

    ' --- 5. ОБРАБОТКА МАССИВА ПРИЁМНИКА ---
    phaseNum = 2
    phaseDesc = "Заполнение столбца O"
    phaseStart = Timer
    lastPct = 0
    
    For i = 1 To destRows
        vA = arrDest(i, 1): If IsError(vA) Then vA = ""
        vC = arrDest(i, 3): If IsError(vC) Then vC = ""
        vD = arrDest(i, 4): If IsError(vD) Then vD = ""
        vF = arrDest(i, 6): If IsError(vF) Then vF = ""
        
        key = CStr(vA & "") & "|" & CStr(vC & "") & "|" & CStr(vD & "") & "|" & CStr(vF & "")
        key = Replace(key, " ", "")
        key = Replace(key, ChrW(160), "")
        key = Replace(key, vbCr, "")
        key = Replace(key, vbLf, "")
        key = Replace(key, vbTab, "")
        
        ' Если ключ найден, записываем значение в столбец O (индекс 15)
        If dict.Exists(key) Then
            arrDest(i, 15) = dict(key)
        End If
        
        ' Обновление прогресс-бара
        currentPct = Int(i / destRows * 100)
        If currentPct > lastPct Then
            lastPct = currentPct
            elapsed = Timer - phaseStart
            If currentPct > 0 Then
                remaining = (elapsed / currentPct) * (100 - currentPct)
            Else
                remaining = 0
            End If
            UpdateStatusBar phaseNum, phaseDesc, currentPct, i, destRows, remaining
        End If
    Next i
    
    Debug.Print "Приёмник обработан. Время: " & Format(Timer - phaseStart, "0.00") & " сек."

    ' --- 6. ВЫГРУЗКА ДАННЫХ И ЗАВЕРШЕНИЕ ---
    phaseDesc = "Запись на лист"
    UpdateStatusBar 2, phaseDesc, 100, destRows, destRows, 0
    
    wsDest.Range("A3").Resize(destRows, 15).Value2 = arrDest
    
    Dim totalTime As Double
    totalTime = Round(Timer - startTime, 2)
    
    Debug.Print "Обработка завершена. Общее время: " & totalTime & " сек."
    Debug.Print String(80, "=")
    
    MsgBox "Обработка завершена!" & vbCrLf & _
           "Уникальных ключей в источнике: " & dict.Count & vbCrLf & _
           "Общее время выполнения: " & totalTime & " сек.", vbInformation, "Готово"

Cleanup:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Set dict = Nothing
    Set wsDest = Nothing: Set wsSrc = Nothing
    Set wbDest = Nothing: Set wbSrc = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Произошла критическая ошибка:" & vbCrLf & Err.Description & vbCrLf & _
           "Код ошибки: " & Err.Number, vbCritical, "Ошибка выполнения"
    Resume Cleanup
End Sub

' =====================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ UI
' =====================================================================

Private Function FormatBar(ByVal pct As Long) As String
    Const W As Long = 20
    Dim filled As Long
    filled = Int(pct * W / 100)
    If filled > W Then filled = W
    FormatBar = "[" & String(filled, ChrW(&H2588)) & _
                String(W - filled, ChrW(&H2591)) & "]"
End Function

Private Function FormatTime(ByVal secs As Double) As String
    Dim m As Long, s As Long
    m = Int(secs / 60)
    s = Int(secs Mod 60)
    FormatTime = m & "м " & s & "с"
End Function

Private Sub UpdateStatusBar(ByVal phase As Long, ByVal desc As String, ByVal pct As Long, _
                            ByVal current As Long, ByVal total As Long, Optional ByVal remTime As Double = -1)
    Dim timeStr As String
    If remTime >= 0 Then
        timeStr = " | ~" & FormatTime(remTime)
    Else
        timeStr = ""
    End If
    
    Application.StatusBar = "Фаза " & phase & "/2 " & FormatBar(pct) & " " & pct & "% — " & desc & " | " & current & " из " & total & timeStr
    DoEvents
End Sub
