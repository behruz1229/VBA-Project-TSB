VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_CAD_SPECIFICATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================
'    Класс для создание таблицы со спецификацией для систем
'      автоматизированного проэктирования работ программ
'=============================================================


'Столбцы
Private Const L_PP$ = "A"
Private Const L_Mat$ = "B"
Private Const L_Cnt$ = "C"
Private Const L_Wght_eq$ = "D"
Private Const L_Wght_sum$ = "E"

'Настройка таблицы
Private Const start_row As Byte = 3

'Переменные класса
Private sum_wght#
Private mass_data()
Private dict_data As Object
Private SH As Excel.Worksheet
Private OtherFunction As New CL_OTHERFUN
Private THWB As New CL_THISWORKBOOK


'Конструктор и деструктор класса
Private Sub Class_Initialize()
    Set dict_data = CreateObject("Scripting.Dictionary")
    mass_data = THWB.MassData
End Sub
Private Sub Class_Terminate()
    Set dict_data = Nothing
    Set THWB = Nothing
    Set OtherFunction = Nothing
    Set SH = Nothing
End Sub

'Номера столбцов
Private Property Get cNumPP%()
    cNumPP = Columns(L_PP).Column
End Property

'Последняя строка данных в таблице
Private Property Get cnt_row&()
    cnt_row = SH.Cells(Rows.Count, cNumPP).End(xlUp).Row
End Property


'Создание документа
Public Sub CreateDocument(ByRef sFilter$, ByRef mass())
    
    'Создание листа если его нет
    Set SH = OtherFunction.CreateSheet(sFilter)
    CreateHeader sFilter    'Создание шапки таблицы
    
    With SH.Range(L_PP & start_row).Resize(UBound(mass), UBound(mass, 2))
        OtherFunction.FormatCells .Offset, xlHAlignCenter, xlVAlignCenter, "Times New Roman", 11, Wrap:=True
        .Value = mass
    End With
    
    CreateFooter    'Создание подвала итого
    
    OtherFunction.BoldersCells SH.Range(L_PP & start_row).CurrentRegion, xlThin
    PrinterConfiguration
    SH.Activate
    
End Sub

'Вывод данных в массив для вставки
Private Function GetMassData()
    Dim q&, w&, i, mass()
    ReDim mass(1 To dict_data.Count, 1 To 5)
    For Each i In dict_data.items
        q = q + 1
        mass(q, 1) = q
        mass(q, 2) = i(1)(0) & " " & i(1)(1)
        mass(q, 3) = i.Count
        mass(q, 4) = i(1)(2)
        mass(q, 5) = i(1)(2) * i.Count
        sum_wght = sum_wght + mass(q, 5)
    Next
    GetMassData = mass
End Function

'Массив вставляемой таблицы
Public Function GetMassDiagramm(ByRef sFilter$)
    LoadDataInDict sFilter
    GetMassDiagramm = GetMassData
End Function

'Формирование массива данных для вставки таблицы Реестр к п.3
Private Sub LoadDataInDict(ByRef sFilter$)
    Dim q&, sKey$
    
    For q = LBound(mass_data) To UBound(mass_data)
        
        If IsError(mass_data(q, idx_diagram)) Then GoTo continue
        
        sKey = Trim(mass_data(q, idx_diagram))
        If sKey <> sFilter Then GoTo continue
        
        sKey = GetKeyElement(q)
        If Not dict_data.exists(sKey) Then dict_data.Add key:=sKey, Item:=New Collection
        dict_data(sKey).Add tmpdict(q)
        
continue:
    Next
    
End Sub
Private Function tmpdict(q) As Object
    Set tmpdict = CreateObject("Scripting.Dictionary")
    With tmpdict
        .Add key:=.Count, Item:=mass_data(q, idx_name)    'Наименование элемента
        .Add key:=.Count, Item:=mass_data(q, idx_mark)    'Марка элемента
        .Add key:=.Count, Item:=mass_data(q, idx_wght)    'Вес элемента
    End With
End Function

'Ключ для словаря с элементами титул, dp, марка, наименование, позиция, вес
Private Property Get GetKeyElement$(ByRef q&)
    Dim sKey$
    sKey = mass_data(q, idx_title) & vbTab
    sKey = sKey & mass_data(q, idx_dp) & vbTab
    sKey = sKey & mass_data(q, idx_mark) & vbTab
    sKey = sKey & mass_data(q, idx_name) & vbTab
    sKey = sKey & mass_data(q, idx_wght)
    GetKeyElement = sKey
End Property

'Формирование названий столбцов
Private Function dict_header() As Object
    Set dict_header = CreateObject("Scripting.Dictionary")
    With dict_header
        .Add key:=.Count + 1, Item:="Поз."
        .Add key:=.Count + 1, Item:="Наименование"
        .Add key:=.Count + 1, Item:="Кол-во"
        .Add key:=.Count + 1, Item:="Масса ед., кг"
        .Add key:=.Count + 1, Item:="Масса общая, кг"
    End With
End Function

'Шапка таблицы Приложение к акту освидетельствования скрытых работ
Private Sub CreateHeader(ByRef sFilter$)
    Dim addr$, tmpstr$
    
    OtherFunction.ColumnsWidth SH, 1, 9, 40, 12, 12, 15 'Ширина столбцов
    
    With SH
        .Rows("1:1").RowHeight = 20 'Высоты строк
        With .Range("A1:E1")    'Заголовоки документа
            OtherFunction.FormatCells .Offset, xlHAlignCenter, xlVAlignBottom, "Calibri", 14, Merge:=True, Bold:=True
            .Value = "Спецификация"
        End With
        With .Range("A2:E2")
            OtherFunction.FormatCells .Offset, xlHAlignCenter, xlVAlignBottom, "Times New Roman", 11
            .Value = dict_header.items
        End With
        If Not .AutoFilterMode Then .Range("A2:E2").AutoFilter
    End With
    
End Sub

'Подвал ИТОГО
Private Sub CreateFooter()
    Dim iRow&
    
    iRow = cnt_row + 1
    
    With SH.Range(L_Wght_sum & iRow)
        OtherFunction.FormatCells .Offset, xlHAlignCenter, xlVAlignCenter, "Times New Roman", 11, Bold:=True
        .Value = sum_wght
    End With
    
    With SH.Range(L_PP & iRow & ":" & L_Wght_eq & iRow)
        OtherFunction.FormatCells .Offset, xlHAlignRight, xlVAlignCenter, "Times New Roman", 11, Bold:=True, Merge:=True
        .Value = "Масса общая:"
    End With
    
End Sub
    

'настройки принтинга листа
Private Sub PrinterConfiguration()

    Application.PrintCommunication = False
    
    With SH.PageSetup
        
        .PrintTitleRows = "4:4"
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "Страница &P из &N"
        .LeftMargin = Application.InchesToPoints(1)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.8)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    
    Application.PrintCommunication = True
    
End Sub

