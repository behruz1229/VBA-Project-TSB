Option Explicit

' =====================================================================
' МАКРОС: Update_status_RFI
' Назначение: Обновление столбцов P, Q, R в приёмнике на основе источника.
' Ключ: Столбец O (приёмник) -> Столбец D (источник).
' Маппинг: Источник C -> Приёмник P, AG -> Q, AH -> R.
' =====================================================================

Sub Update_status_RFI()
    ' --- ОБЪЯВЛЕНИЯ И КОНСТАНТЫ ---
    ' <--- ИЗМЕНИТЕ ПУТЬ НА АКТУАЛЬНЫЙ (оставьте слэш в конце)
    Const DEFAULT_SRC_FOLDER As String = "\\vls.lan\ULVZG-DFS\ПТС\1.14. ТСБ и МОТ_Исполнительная документация КМ\Выгрузки\RFI\"
    Const FILE_MASK As String = "Инспекции на*.xlsx"
    Const DEST_SHEET_NAME As String = "База данных по элементам"
    Const SRC_SHEET_NAME As String = "Инспекции"
    
    Dim wbDest As Workbook, wsDest As Worksheet
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim srcFilePath As Variant
    Dim dictP As Object, dictQ As Object, dictR As Object
    Dim key As String
    Dim lastRowDest As Long, lastRowSrc As Long
    Dim arrDest As Variant, arrSrc As Variant
    Dim i As Long, srcRows As Long, destRows As Long
    Dim startTime As Double, phaseStart As Double
    Dim elapsed As Double, remaining As Double
    Dim lastPct As Long, currentPct As Long, phaseNum As Long
    Dim phaseDesc As String
    
    ' Переменные для поиска файла
    Dim fileName As String, newestFile As String
    Dim newestDate As Date
    
    ' Переменные для значений
    Dim vKey As Variant, vC As Variant, vAG As Variant, vAH As Variant
    Dim valC As String, valAG As String, valAH As String
    Dim existingP As String, existingQ As String, existingR As String
    
    ' --- ИНИЦИАЛИЗАЦИЯ ---
    startTime = Timer
    Set wbDest = ActiveWorkbook
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = False
    
    On Error GoTo ErrorHandler
    
    Debug.Print String(80, "=")
    Debug.Print "Запуск Update_status_RFI: " & Now
    
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
        
        ' Очищаем столбцы P (16), Q (17), R (18)
        .Range(.Cells(3, 16), .Cells(lastRowDest, 18)).ClearContents
        
        ' Читаем массив строго от A (1) до R (18)
        arrDest = .Range(.Cells(3, 1), .Cells(lastRowDest, 18)).Value2
    End With
    
    destRows = UBound(arrDest, 1)
    Debug.Print "Приёмник прочитан. Строк: " & destRows & ". Время: " & Format(Timer - phaseStart, "0.00") & " сек."

    ' --- 3. ПОИСК ФАЙЛА-ИСТОЧНИКА ПО МАСКЕ ---
    phaseDesc = "Поиск источника"
    UpdateStatusBar 0, phaseDesc, 0, 0, 0, -1
    
    ' Ищем самый свежий файл в папке
    fileName = Dir(DEFAULT_SRC_FOLDER & FILE_MASK)
    newestDate = #1/1/1900#
    
    Do While fileName <> ""
        If FileDateTime(DEFAULT_SRC_FOLDER & fileName) > newestDate Then
            newestDate = FileDateTime(DEFAULT_SRC_FOLDER & fileName)
            newestFile = DEFAULT_SRC_FOLDER & fileName
        End If
        fileName = Dir
    Loop
    
    If newestFile <> "" Then
        srcFilePath = newestFile
    Else
        ' Fallback: если в папке ничего нет, спрашиваем пользователя
        If MsgBox("Файлы по маске '" & FILE_MASK & "' не найдены в папке:" & vbCrLf & DEFAULT_SRC_FOLDER & vbCrLf & vbCrLf & "Выбрать файл вручную?", _
                  vbYesNo + vbExclamation, "Файл не найден") = vbYes Then
            srcFilePath = Application.GetOpenFilename("Excel Files (*.xlsx; *.xlsb; *.xlsm), *.xlsx; *.xlsb; *.xlsm")
            If VarType(srcFilePath) = vbBoolean Then GoTo Cleanup
        Else
            GoTo Cleanup
        End If
    End If

    ' --- 4. ОТКРЫТИЕ ИСТОЧНИКА ---
    Set wbSrc = Workbooks.Open(fileName:=srcFilePath, UpdateLinks:=False, ReadOnly:=True)
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(SRC_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If wsSrc Is Nothing Then
        MsgBox "В файле '" & wbSrc.Name & "' отсутствует лист '" & SRC_SHEET_NAME & "'.", vbCritical, "Лист не найден"
        GoTo Cleanup
    End If

    ' --- 5. ЧТЕНИЕ ИСТОЧНИКА И ПОСТРОЕНИЕ СЛОВАРЕЙ ---
    phaseNum = 1
    phaseDesc = "Построение словарей"
    phaseStart = Timer
    lastPct = 0
    
    With wsSrc
        lastRowSrc = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRowSrc < 2 Then
            MsgBox "На листе источника нет данных.", vbExclamation, "Нет данных"
            GoTo Cleanup
        End If
        ' Читаем от A до AH (34 столбца), чтобы индексы совпадали с буквами
        arrSrc = .Range(.Cells(2, 1), .Cells(lastRowSrc, 34)).Value2
    End With
    
    ' Закрываем источник сразу после чтения
    wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing: Set wsSrc = Nothing
    
    srcRows = UBound(arrSrc, 1)
    
    ' Создаем 3 словаря для P, Q, R
    Set dictP = CreateObject("Scripting.Dictionary"): dictP.CompareMode = vbBinaryCompare
    Set dictQ = CreateObject("Scripting.Dictionary"): dictQ.CompareMode = vbBinaryCompare
    Set dictR = CreateObject("Scripting.Dictionary"): dictR.CompareMode = vbBinaryCompare
    
    For i = 1 To srcRows
        ' Ключ из столбца D (индекс 4)
        vKey = arrSrc(i, 4): If IsError(vKey) Then vKey = ""
        
        ' Значения: C (3), AG (33), AH (34)
        vC = arrSrc(i, 3): If IsError(vC) Then vC = ""
        vAG = arrSrc(i, 33): If IsError(vAG) Then vAG = ""
        vAH = arrSrc(i, 34): If IsError(vAH) Then vAH = ""
        
        ' Сборка и жесткая очистка ключа
        key = CStr(vKey & "")
        key = Replace(key, " ", "")
        key = Replace(key, ChrW(160), "")
        key = Replace(key, vbCr, "")
        key = Replace(key, vbLf, "")
        key = Replace(key, vbTab, "")
        
        If Len(key) > 0 Then
            If Not dictP.Exists(key) Then
                ' Ключа нет - добавляем все три значения
                dictP(key) = vC
                dictQ(key) = vAG
                dictR(key) = vAH
            Else
                ' === ВАРИАНТ Б: Не перезаписывать непустое значение ===
                existingP = CStr(dictP(key) & "")
                existingQ = CStr(dictQ(key) & "")
                existingR = CStr(dictR(key) & "")
                
                If Len(Trim(existingP)) = 0 And Len(Trim(CStr(vC & ""))) > 0 Then dictP(key) = vC
                If Len(Trim(existingQ)) = 0 And Len(Trim(CStr(vAG & ""))) > 0 Then dictQ(key) = vAG
                If Len(Trim(existingR)) = 0 And Len(Trim(CStr(vAH & ""))) > 0 Then dictR(key) = vAH
            End If
        End If
        
        ' Обновление прогресс-бара
        currentPct = Int(i / srcRows * 100)
        If currentPct > lastPct Then
            lastPct = currentPct
            elapsed = Timer - phaseStart
            If currentPct > 0 Then remaining = (elapsed / currentPct) * (100 - currentPct) Else remaining = 0
            UpdateStatusBar phaseNum, phaseDesc, currentPct, i, srcRows, remaining
        End If
    Next i
    
    Debug.Print "Словари построены. Уникальных ключей: " & dictP.Count & ". Время: " & Format(Timer - phaseStart, "0.00") & " сек."

    ' --- 6. ОБРАБОТКА МАССИВА ПРИЁМНИКА ---
    phaseNum = 2
    phaseDesc = "Заполнение P, Q, R"
    phaseStart = Timer
    lastPct = 0
    
    For i = 1 To destRows
        ' Ключ из столбца O (индекс 15)
        vKey = arrDest(i, 15): If IsError(vKey) Then vKey = ""
        
        key = CStr(vKey & "")
        key = Replace(key, " ", "")
        key = Replace(key, ChrW(160), "")
        key = Replace(key, vbCr, "")
        key = Replace(key, vbLf, "")
        key = Replace(key, vbTab, "")
        
        ' Если ключ найден, записываем значения в P (16), Q (17), R (18)
        If dictP.Exists(key) Then
            arrDest(i, 16) = dictP(key)
            arrDest(i, 17) = dictQ(key)
            arrDest(i, 18) = dictR(key)
        End If
        
        ' Обновление прогресс-бара
        currentPct = Int(i / destRows * 100)
        If currentPct > lastPct Then
            lastPct = currentPct
            elapsed = Timer - phaseStart
            If currentPct > 0 Then remaining = (elapsed / currentPct) * (100 - currentPct) Else remaining = 0
            UpdateStatusBar phaseNum, phaseDesc, currentPct, i, destRows, remaining
        End If
    Next i
    
    Debug.Print "Приёмник обработан. Время: " & Format(Timer - phaseStart, "0.00") & " сек."

    ' --- 7. ВЫГРУЗКА ДАННЫХ И ЗАВЕРШЕНИЕ ---
    phaseDesc = "Запись на лист"
    UpdateStatusBar 2, phaseDesc, 100, destRows, destRows, 0
    
    wsDest.Range("A3").Resize(destRows, 18).Value2 = arrDest
    
    Dim totalTime As Double
    totalTime = Round(Timer - startTime, 2)
    
    Debug.Print "Обработка завершена. Общее время: " & totalTime & " сек."
    Debug.Print String(80, "=")
    
    MsgBox "Обработка завершена!" & vbCrLf & _
           "Уникальных ключей в источнике: " & dictP.Count & vbCrLf & _
           "Общее время выполнения: " & totalTime & " сек.", vbInformation, "Готово"

Cleanup:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Set dictP = Nothing: Set dictQ = Nothing: Set dictR = Nothing
    Set wsDest = Nothing: Set wsSrc = Nothing
    Set wbDest = Nothing: Set wbSrc = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Произошла критическая ошибка:" & vbCrLf & Err.Description & vbCrLf & _
           "Код ошибки: " & Err.Number, vbCritical, "Ошибка выполнения"
    Resume Cleanup
End Sub

' =====================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ UI (Идентичны первому макросу)
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



