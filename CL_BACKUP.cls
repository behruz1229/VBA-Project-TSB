VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_BACKUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Настройки архивирования файла
Private Const sPath7z$ = "C:\Program Files\7-Zip\7z.exe"
Private Const Arhiv As String = " a -t7z -ssw -mmt4 -mx9 -sdel "

'Имя каталога для резервного копирования файлов
Private Const sFolderNameBackUpFiles$ = "Резервные копии файлов"

'Число файлов в резервном каталоге
Private Const CntBackUpFiles As Byte = 40

'Переменные класса
Private ps$
Private sFolder$
Private objFSO As Object
Private WshShell As Object

'Констрктор и деструктор класса
Private Sub Class_Initialize()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set WshShell = CreateObject("WScript.Shell")
    ps = Application.PathSeparator
    sFolder = ThisWorkbook.Path & ps & sFolderNameBackUpFiles & ps & objFSO.GetBaseName(ThisWorkbook.Name)
End Sub
Private Sub Class_Terminate()
    Set objFSO = Nothing
    Set WshShell = Nothing
End Sub

'Резервное копирование таблицы
Public Sub GetBackUpTable()
    Dim q&, sFile$
    
    If ThisWorkbook.ReadOnly Then Exit Sub
    
    Application.StatusBar = "Резервное копирование: " & ThisWorkbook.Name & "..."
    Debug.Print Application.StatusBar: DoEvents
    
    'Создание каталога для резервного копирования если его нет
    If Not objFSO.FolderExists(sFolder) Then CreateDir sFolder
    
    'Удаление старых файлов
    KeepTheLastTenFiles sFolder
    
    'Формирование имени файла
    sFile = sFolder & ps & Format(Date, "yyyy_mm_dd")
    sFile = sFile & "_" & Format(Time, "hh-mm-ss")
    sFile = sFile & "_" & Environ("USERNAME")
    sFile = sFile & "_" & ThisWorkbook.Name
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Резервное копирование файла
    ThisWorkbook.SaveCopyAs Filename:=sFile
    
    'Изменение формата на архивный .xlsb (не использовать)
    'ChangeXlFormat sFile
    
    'Архивирование файла
    GetCompressionFile sFile
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
    Debug.Print "Резервное копирование: OK!"

End Sub

'Удаление старых файлов
Private Sub KeepTheLastTenFiles(ByRef sFoler$)
    Dim iFile As Object, dict_files As Object, dtKey
    
    If objFSO.GetFolder(sFoler).Files.Count < CntBackUpFiles Then Exit Sub
    
    Set dict_files = CreateObject("Scripting.Dictionary")
    
    'Чтение каталога с файлами
    For Each iFile In objFSO.GetFolder(sFoler).Files
        dtKey = iFile.DateLastModified
        If Not dict_files.exists(dtKey) Then dict_files.Add key:=dtKey, Item:=iFile.Path
    Next
    
    'Сортировка файлов по дате создания
    Dim SRT As New CL_SORT
    SRT.SetSortOrder = set_ASC
    Set dict_files = SRT.Sort(dict_files)
    Set SRT = Nothing
    
    'Удаление самых старых файлов
    For Each dtKey In dict_files.items
        If objFSO.GetFolder(sFoler).Files.Count < CntBackUpFiles Then Exit For
        objFSO.DeleteFile dtKey, True
    Next
    
    Set dict_files = Nothing
    
End Sub

'Смена формата на xlsb (архивный формат)
Private Sub ChangeXlFormat(ByRef sFile$)
    Dim WB As Excel.Workbook
    Set WB = Workbooks.Open(sFile, False, True)
    WB.SaveAs Left(sFile, InStrRev(sFile, ".") - 1), XlFileFormat.xlExcel12
    If objFSO.FileExists(sFile) Then objFSO.DeleteFile sFile, True
    sFile = WB.FullName
    WB.Close
    Set WB = Nothing
End Sub

'Архивирование файла
Private Sub GetCompressionFile(ByRef sFile$)
    Dim sFileArj$, sShell$, RetCode&
    
    If Not objFSO.FileExists(sPath7z) Then Exit Sub
    
    'Формирование коммандной строки
    sFileArj = Left(sFile, InStrRev(sFile, ".") - 1) '& ".zip"
    sShell = Chr(34) & sPath7z & Chr(34) & Arhiv
    sShell = sShell & Chr(34) & sFileArj & Chr(34) & " "
    sShell = sShell & Chr(34) & sFile & Chr(34)

    RetCode = WshShell.Run(sShell, 0, True)
    If RetCode <> 0 Then If objFSO.FileExists(sFileArj) Then objFSO.DeleteFile sFileArj, True

End Sub
