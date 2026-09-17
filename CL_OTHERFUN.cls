VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_OTHERFUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Установка ширины столбцов
Public Sub ColumnsWidth(ByRef SH As Excel.Worksheet, ByRef start_col%, ParamArray arr())
    Dim q&, w&: For q = start_col To start_col + UBound(arr)
        SH.Columns(q).ColumnWidth = arr(w): w = w + 1
    Next
End Sub

'Функция форматирования ячеек
Public Sub FormatCells(ByRef RG As Range, _
    Optional ByRef HA As XlHAlign = xlCenter, _
    Optional ByRef VA As XlVAlign = xlCenter, _
    Optional ByRef FontName$ = "Arial", _
    Optional ByRef FontSize% = 10, _
    Optional ByRef Merge As Boolean = False, _
    Optional ByRef Bold As Boolean = False, _
    Optional ByRef Wrap As Boolean = False, _
    Optional ByRef Italic As Boolean = False)
 
    If Merge Then RG.Merge
    RG.HorizontalAlignment = HA
    RG.VerticalAlignment = VA
    RG.Font.Name = FontName
    RG.Font.Size = FontSize
    RG.Font.Bold = Bold
    RG.Font.Italic = Italic
    RG.WrapText = Wrap
    
End Sub

'Разлиновка таблицы (Создание границ ячеек)
Public Sub BoldersCells(ByRef RG As Range, Optional ByRef xLine As XlBorderWeight = xlThin)
    RG.Borders(xlEdgeTop).Weight = xLine
    RG.Borders(xlEdgeBottom).Weight = xLine
    RG.Borders(xlEdgeLeft).Weight = xLine
    RG.Borders(xlEdgeRight).Weight = xLine
    RG.Borders(xlInsideVertical).Weight = xLine
    RG.Borders(xlInsideHorizontal).Weight = xLine
End Sub

'Проверка наличия листа в книге
Public Function shExists(ByRef WB As Excel.Workbook, ByRef shName$) As Boolean
    On Error Resume Next
    shExists = (WB.Worksheets(shName).Name = shName)
End Function

'Функция добавления листа с именем sTXT
Public Function CreateSheet(ByVal shName$) As Excel.Worksheet
    
    shName = IIf(Len(shName) > 29, ".." & Right(shName, 29), shName)
    
    'Добавление нового листа в книгу если его нет
    With ThisWorkbook.Worksheets
        If Not shExists(ThisWorkbook, shName) Then .Add(after:=.Item(.Count)).Name = shName
        Set CreateSheet = .Item(shName)
    End With
    
    CreateSheet.Cells.Delete
        
End Function
