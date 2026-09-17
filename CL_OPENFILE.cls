VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_OPENFILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'======================================================
'Класс поиска и открытия файлов с таблицами Excel
'======================================================

'Переменные класса:
Private objFSO As Object
Private isTableOpen As Boolean
Private WB As Excel.Workbook

Private sFile$
Private mask_file$
Private coll_folder As New Collection
Private isShowFileDialog As Boolean

'Конструктор и деструктор класса
Private Sub Class_Initialize()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    coll_folder.Add ThisWorkbook.Path   'Добавление рабочего каталога с макросом для поиска вайла
End Sub
Private Sub Class_Terminate()
    Set objFSO = Nothing
    If Not isTableOpen And Not WB Is Nothing Then WB.Close False
    Set WB = Nothing
End Sub

'Пинг только сетевой ресурс на доступность
Private Property Get Ping(ByVal addr$) As Boolean
    Dim SQL$
    If InStr(1, addr, "\\", vbBinaryCompare) = 0 Then Ping = True: Exit Property
    addr = GetIpAddressFromLine(addr) 'IP- Адрес или адрес домена из адресной строки
    On Error Resume Next
    SQL$ = "SELECT * FROM Win32_PingStatus WHERE Address = '" & addr & "'"
    Ping = GetObject("winmgmts:").ExecQuery(SQL).ItemIndex(0).StatusCode = 0
End Property
Private Property Get GetIpAddressFromLine$(ByRef sTXT$)
    Dim i_StartLet%, i_EndLet%
    i_StartLet = InStr(1, sTXT, "\\", vbBinaryCompare) + 2
    i_EndLet = InStr(i_StartLet, sTXT, "\", vbBinaryCompare)
    GetIpAddressFromLine = Mid(sTXT, i_StartLet, Abs(i_EndLet - i_StartLet))
End Property

'Настройка: Установка маски файла для поиска
Public Property Let SetMaskFile(ByRef strmask$)
    mask_file = GetFileNameSTR(strmask)
End Property

'Настройка: диалоговое окно выбора файла, если он не найден
Public Property Let ShowFileDialog(ByRef isShow As Boolean)
    isShowFileDialog = isShow
End Property

'Доступ к коллекции каталогов поиска файла
Public Property Get ListFolder() As Collection
    Set ListFolder = coll_folder
End Property

'Вывод имени загружаемого файла
Public Property Get GetFileName$()
    If Not isExistsFile Then Exit Property
    GetFileName = objFSO.GetFileName(sFile)
End Property

'Вывод имени загружаемого файла
Public Property Get GetFullFileName$()
    If Not isExistsFile Then Exit Property
    GetFullFileName = sFile
End Property

'Вывод даты и времени последнего изменения загружаемого файла
Public Property Get GetDateLastModified() As Date
    If Not isExistsFile Then Exit Property
    GetDateLastModified = objFSO.GetFile(sFile).DateLastModified
End Property

'Передача рабочей книги с данными из класса
Public Function GetWorkBook() As Excel.Workbook
    If Not isExistsFile Then Exit Function
    OpenFile
    Set GetWorkBook = WB
End Function


'Проверка и поиск\выбор файла книги
Private Property Get isExistsFile() As Boolean
    If Not sFile = "" Then isExistsFile = True: Exit Property
    If isFoundFile Then isExistsFile = True: Exit Property Else sFile = FileDialog(mask_file)
    If objFSO.FileExists(sFile) Then isExistsFile = True
End Property


'Возвращает имя файла из текста полного пути или sPathFile$
Private Property Get GetFileNameSTR$(ByVal sPathFile$)
    Dim cntletter%
    cntletter = InStr(1, StrReverse(sPathFile), Application.PathSeparator, vbBinaryCompare)
    If cntletter > 0 Then
        GetFileNameSTR = Right(sPathFile, cntletter - 1)
        sPathFile = Replace(sPathFile, Right(sPathFile, cntletter), "")
        coll_folder.Add sPathFile
        Exit Property
    End If
    GetFileNameSTR = sPathFile
End Property

'Поиск файлов в каталогах из коллекции файлов
Private Property Get isFoundFile() As Boolean
    Dim i, sFolder$
    For Each i In coll_folder
        sFolder = i
        If Ping(sFolder) Then
            sFile = GetSearchFile(sFolder)
            If objFSO.FileExists(sFile) Then isFoundFile = True: Exit Property
        Else
            Err.Raise 75, sFolder, "Ошибка доступа к пути/файлу"
        End If
    Next
End Property

'Функция поиска полного пути файла
Private Property Get GetSearchFile$(sFolder$)
    Dim iFile As Object
    If Not objFSO.FolderExists(sFolder) Then Exit Property
    If objFSO.GetFolder(sFolder).Files.Count = 0 Then Exit Property
    For Each iFile In objFSO.GetFolder(sFolder).Files
        If InStr(1, iFile.Name, "~$", vbBinaryCompare) Then GoTo continue
        If InStr(1, mask_file, "*", vbBinaryCompare) > 0 Then
            If Not LCase(iFile.Name) Like "*" & LCase(mask_file) & "*" Then GoTo continue
        Else
            If InStr(1, iFile.Name, mask_file, vbTextCompare) = 0 Then GoTo continue
        End If
        GetSearchFile = iFile.Path
        Exit Property
continue:
    Next
End Property

'Диалоговое окно для выбора файла
Private Function FileDialog$(ByRef sFile$)
    If Not isShowFileDialog Then Exit Function
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*;*.xla*"
        .ButtonName = "Выбрать"
        .Title = "Выберите файл " & sFile
        .InitialFileName = ThisWorkbook.Path & Application.PathSeparator
        If .Show = 0 Then isShowFileDialog = False: Exit Function
        FileDialog = .SelectedItems(1)
    End With
End Function

'Открытие таблицы
Private Sub OpenFile()
    Dim sShortFileName$, ScreenUpdating As Boolean, EnableEvents As Boolean
    
    With Application
    
        ScreenUpdating = .ScreenUpdating    'Сохранение состояния флагов макроса
        EnableEvents = .EnableEvents        'Сохранение состояния флагов макроса
        sShortFileName = objFSO.GetFileName(sFile) 'Имя файла с расширением
        
        If isBookOpen(sShortFileName) Then
            Set WB = Workbooks(sShortFileName)
            
            isTableOpen = True
        Else
        
            If EnableEvents Then .EnableEvents = False
            If ScreenUpdating Then .ScreenUpdating = False
            
            Set WB = Workbooks.Open(sFile, False, True)
            
            WB.Windows(1).Visible = False
            If ScreenUpdating Then .ScreenUpdating = True
            If EnableEvents Then .EnableEvents = True
            isTableOpen = False
        End If
        
    End With
End Sub

'Функция для проверки открыт ли фаил с данными для дальнейшей работы
Private Property Get isBookOpen(sWbName$) As Boolean
    On Error Resume Next
    isBookOpen = Not Workbooks(sWbName) Is Nothing
End Property


