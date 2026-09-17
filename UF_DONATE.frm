VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_DONATE 
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3675
   OleObjectBlob   =   "UF_DONATE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_DONATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit

'константы для функций API
Private Const GWL_STYLE As Long = -16& 'для установки нового вида окна
Private Const GWL_EXSTYLE = -20& 'для расширенного стиля окна
Private Const WS_CAPTION As Long = &HC00000 'определяет заголовок
Private Const WS_BORDER As Long = &H800000 'определяет рамку формы

'Функции API, применяемые для поиска окна и изменения его стиля
#If VBA7 Then
    Private Declare PtrSafe Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "User32" (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function DrawMenuBar Lib "user32.dll" (ByVal hwnd As Long) As Long
#End If

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Source$
    Source = "https://www.tinkoff.ru/cf/7xgUs1rFSvi"
    MsgBox Source, vbInformation, "Сбор на отдых"
End Sub

Private Sub UserForm_Initialize()
    Dim ihWnd As LongPtr, hStyle As LongPtr
    
    Width = 174     'Размер окна ширина
    Height = 212    'Размер окна высота
    
    'ищем окно формы среди всех открытых окон
    If val(Application.Version) < 9 Then
        ihWnd = FindWindow("ThunderXFrame", Me.Caption) 'для Excel 97
    Else
        ihWnd = FindWindow("ThunderDFrame", Me.Caption) 'для Excel 2000 и выше
    End If
    'получаем информацию о найденном окне(стили и т.д.)
    hStyle = GetWindowLong(ihWnd, GWL_STYLE)
    'назначаем переменной новый стиль для окна формы
    hStyle = hStyle And Not WS_CAPTION And Not WS_BORDER
    'изменяем вид окна: убираем меню(заголовок) и рамку
    SetWindowLong ihWnd, GWL_STYLE, hStyle
    SetWindowLong ihWnd, GWL_EXSTYLE, 0
    'перерисовываем форму, точнее строку меню(заголовка)
    DrawMenuBar ihWnd
    'меняем размер формы, т.к. сделали смещение элементов формы вверх на высоту заголовка
    Me.Height = Me.Height + (GWL_EXSTYLE - 9)
    
    'Настройка формы при запуске
    StartUpPosition = 0
    Top = Application.Top + (Application.Height / 2) - (Height / 2) + 10
    Left = Application.Left + (Application.Width / 2) - Height
    
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
