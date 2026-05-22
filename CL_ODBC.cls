VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_ODBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'======================================================
'Класс открытия и чтения файлов с таблицами Excel средствами ADO
'======================================================

'Настройки подключения
Private Const sDrv$ = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};ReadOnly=1;"


'Переменные класса
Private ps$, sTmpFolder$, sSrcFile$, sLdtFile$
Private str_from_$
Private str_filter_$
Private str_order_$
Private str_group_$

Private CN As Object    'Connection (DAO)
Private RS As Object    'Recordsets (DAO)
Private CT As Object    'ADOX.Catalog

Private objFSO As Object

Private coll_fields_name As Collection
Private coll_select_flds As Collection

'Тип блокировки для Recordset, который будет открыт
Private Enum LockType
    adLockReadOnly = 1
    adLockPessimistic
    adLockOptimistic
    adLockBatchOptimistic
    adLockUnspecified = -1
End Enum

'Specifies whether an object is open or closed, _
 connecting to a data source, executing a command, or retrieving data.
Private Enum State
    adStateClosed = 0       'The object is closed
    adStateOpen = 1         'The object is open
    adStateConnecting = 2   'The object is connecting
    adStateExecuting = 4    'The object is executing a command
    adStateFetching = 8     'The rows of the object are being retrieved
End Enum

'Определяет режим доступа для изменения данных в сеансе.
Private Enum Mode
    adModeUnknown = 0           'режим доступа не установлен или не может быть определён (по умолчанию).
    adModeRead = 1              'режим только для чтения.
    adModeWrite = 2             'режим только для записи.
    adModeReadWrite = 3         'режим для чтения и записи.
    adModeShareDenyRead = 4     'не разрешает открывать соединение на чтение другим пользователям.
    adModeShareDenyWrite = 8    'не разрешает открывать соединение на запись другим пользователям.
    adModeShareExclusive = 12   'не разрешает открывать соединение другим пользователям.
    adModeShareDenyNone = 16    'разрешает открывать соединение с любым видом доступа другим пользователям.
End Enum


'Конструктор и деструктор класса
Private Sub Class_Initialize()
    
    Set CN = CreateObject("ADODB.Connection")
    Set RS = CreateObject("ADODB.Recordset")
    Set CT = CreateObject("ADOX.Catalog")
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Создание временного каталога
    ps = Application.PathSeparator
    sTmpFolder = ThisWorkbook.Path & ps & "tmp_"
    If Not objFSO.FolderExists(sTmpFolder) Then CreateDir sTmpFolder
    
    CN.Mode = Mode.adModeRead
    RS.LockType = LockType.adLockReadOnly
End Sub
Private Sub Class_Terminate()
    
    If Not RS.State = adStateClosed Then RS.Close
    Set RS = Nothing
    
    If Not CN.State = adStateClosed Then CN.Close
    Set CN = Nothing
    
    Set CT = Nothing
    
    Set coll_fields_name = Nothing
    Set coll_select_flds = Nothing
    
    'Удаление временного каталога
    If objFSO.FolderExists(sTmpFolder) Then objFSO.DeleteFolder sTmpFolder, True
    Set objFSO = Nothing
    
    If debug_mode Then Application.StatusBar = False
    
End Sub

'Вывод файла из которого загружаются данные
Public Property Get GetLoadedFileName$()
    GetLoadedFileName = sLdtFile
End Property

'Исполнение запроса к базе данных
Private Property Let Set_Recordset(ByRef sSql$)
    If debug_mode Then Debug.Print "sql: " & sSql
    If Not RS.State = adStateClosed Then RS.Close
    RS.Open sSql & ";", CN
End Property


'Установка имени файла excel
Public Function SetXlFile(ByVal sFileName$) As Boolean
    
    'Проверка наличия файла
    If objFSO.FileExists(ThisWorkbook.Path & ps & objFSO.GetFileName(sFileName)) Then
        sFileName = ThisWorkbook.Path & ps & objFSO.GetFileName(sFileName)
    End If
    If Not objFSO.FileExists(sFileName) Then Exit Function
    
    'Копирование файла источника данных
    sLdtFile = sFileName
    sSrcFile = sTmpFolder & ps & objFSO.GetFileName(sFileName)
    objFSO.CopyFile sFileName, sSrcFile, True
    
    'Загрузка файла для чтения
    If Not CN.State = adStateClosed Then CN.Close
    CN.Open sDrv & "DBQ=" & sSrcFile & ";", -1
    
    If debug_mode Then Application.StatusBar = "Load: " & sFileName
    If debug_mode Then Debug.Print Application.StatusBar: DoEvents
    
    CT.ActiveConnection = CN
    DeleteParam
    
    SetXlFile = True
    
End Function

'Возвращает колекцию листов из подключенной книги
Public Function GetCollSheets() As Collection
    Dim SH, tmpstr$
    
    Set GetCollSheets = New Collection
    For Each SH In CT.Tables
        tmpstr = SH.Name
        
        If InStr(1, tmpstr, "$", vbBinaryCompare) = 0 Then GoTo continue
        If InStr(1, tmpstr, "FilterDatabase", vbBinaryCompare) > 0 Then GoTo continue
        If InStr(1, tmpstr, "Print_Titles", vbBinaryCompare) > 0 Then GoTo continue
        
        tmpstr = Left(tmpstr, InStr(2, tmpstr, "$", vbBinaryCompare) - 1)
        tmpstr = Replace(tmpstr, "'", "")
        tmpstr = Replace(tmpstr, "#", ".")
        
        GetCollSheets.Add tmpstr
continue:
    Next
    
End Function


'Возвращает имя столбца по его индексу
Public Property Get GetFieldName$(ByRef idx_fld%)
    GetFieldName$ = GetCollFields(idx_fld)
End Property

'Возвращает коллекцию имен столбцов в таблице
Public Property Get GetCollFields() As Collection
    If coll_fields_name.Count = 0 Then LoadFieldCollection
    Set GetCollFields = coll_fields_name
End Property
Private Sub LoadFieldCollection()
    Dim i, tmpstr$
    Set_Recordset = str_fields_ & str_from_
    For Each i In RS.Fields
        coll_fields_name.Add GetSqlString(i.Name)
    Next
End Sub


'Установка коллекции возвращаемых столбцов в запросе SQL (SELECT ...)
Public Property Set SELECT_(ByRef coll As Collection)
    Set coll_select_flds = coll
End Property
Public Property Get SELECT_() As Collection
    If coll_select_flds.Count = 0 Then str_filter_ = "": str_order_ = "": str_group_ = ""
    Set SELECT_ = coll_select_flds
End Property
Private Property Get str_fields_$()
    Dim i, tmpstr$
    If coll_select_flds.Count = 0 Then str_fields_ = "SELECT *": Exit Property
    For Each i In coll_select_flds: tmpstr = tmpstr & i & ", ":  Next
    str_fields_ = "SELECT " & Left(tmpstr, Len(tmpstr) - 2)
End Property

'Установка имени листа с таблицей
Public Property Let FROM(ByRef sFrom$)
    str_from_ = " FROM [" & sFrom & "$]"
    DeleteParam
End Property

'Фильтрация строк
Public Property Let WHERE(ByRef sFltr$)
    str_filter_ = " WHERE " & sFltr
End Property

'Сортировка строк
Public Property Let ORDER_BY(ByRef sOrder$)
    str_order_ = " ORDER BY " & sOrder
End Property

'Группировка строк
Public Property Let GROUP_BY(ByRef sGroup$)
    str_group_ = " GROUP BY " & sGroup
End Property

'Загрузка данных из таблицы
Public Function Execute(Optional ByVal sSql$ = "")
    Dim mass
    
    If sSql = "" Then 'Пример: "SELECT * FROM [WB$]"
        sSql = str_fields_ & str_from_
        sSql = sSql & IIf(str_filter_ = "", "", str_filter_)
        sSql = sSql & IIf(str_order_ = "", "", str_order_)
        sSql = sSql & IIf(str_group_ = "", "", str_group_)
    End If
    
    Set_Recordset = sSql
    If Not RS.EOF Then mass = RS.GetRows(-1)
    
    If debug_mode Then If Not IsEmpty(mass) Then _
        Debug.Print "res: OK!", UBound(mass, 2) + 1 Else Debug.Print "res: ERR"
    
    If IsEmpty(mass) Then Exit Function
    
    Execute = Transpose_(mass)
    
End Function

'Переворачивание массива данных строки столбцы
Private Function Transpose_(ByRef mass)
    Dim q&, w&, tmpmass()
    
    Select Case UBound(mass)
        Case Is = 0 'Преобразование одномерного массива
            ReDim tmpmass(1 To UBound(mass, 2) + 1)
            For q = LBound(tmpmass) To UBound(tmpmass)
                tmpmass(q) = cNullToEmpty(mass(0, q - 1))
            Next
            
        Case Else 'Преобразование двумерного массива
            ReDim tmpmass(1 To UBound(mass, 2) + 1, 1 To UBound(mass, 1) + 1)
            For q = LBound(tmpmass) To UBound(tmpmass)
                For w = LBound(tmpmass, 2) To UBound(tmpmass, 2)
                    tmpmass(q, w) = cNullToEmpty(mass(w - 1, q - 1))
                Next
            Next
            
    End Select
    
    Transpose_ = tmpmass
    
End Function

'Обвертование строки в одинарные кавычки если есть пробелы
Private Property Get GetSqlString$(ByRef sTXT$)
    'GetSqlString = "`" & sTXT & "`"
    GetSqlString = IIf(InStr(1, sTXT, Chr(32), vbBinaryCompare) > 0, "`" & sTXT & "`", sTXT)
End Property

'Преобразование значения Null в Empty
Private Property Get cNullToEmpty$(ByRef val)
    If IsNull(val) Then cNullToEmpty = Empty Else cNullToEmpty = val
End Property

'Сброс значений переменных
Private Sub DeleteParam()
    Set coll_select_flds = New Collection
    Set coll_fields_name = New Collection
    str_filter_ = ""
    str_order_ = ""
    str_group_ = ""
End Sub

