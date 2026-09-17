VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL_SORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'======================================================
'Класс для сортировки массивов или словарей
'======================================================

'Методы сортировки:
Public Enum SortOrder
    set_ASC
    set_DESC
End Enum
Private set_sort As SortOrder

'Мутод уствновки и чтения метода сортировки
Public Property Let SetSortOrder(ByRef Order As SortOrder)
    set_sort = Order
End Property
Public Property Get SetSortOrder() As SortOrder
    SetSortOrder = set_sort
End Property

'Сортировка словаря(по ключам)\массива(одномерный)
Public Function Sort(ByRef val, Optional ByRef field%)
    Dim mass()
    Select Case TypeName(val)
        Case Is = "Dictionary"
            If val.Count <= 1 Then Set Sort = val: Exit Function
            mass = val.keys
            SelectionSortSingle mass
            Set Sort = CreateNewDictionary(mass, val)
        
        Case Is = "Collection"
            If val.Count <= 1 Then Set Sort = val: Exit Function
            mass = GetIndexMass(val, field)
            SelectionSortMulti mass, 2
            Set Sort = CreateNewCollection(mass, val)
            
        Case Else
            mass = val
            If isSingledImensMass(mass) Then SelectionSortSingle mass Else SelectionSortMulti mass, field
            Sort = mass
        End Select
End Function

'Проверка размерности массива
Private Property Get isSingledImensMass(ByRef val) As Boolean
    Dim C&: On Error Resume Next
    C = UBound(val, 1)
    isSingledImensMass = IIf(Err.Number = 0, True, False)
End Property

'Создание подменного массива (индекс по порядку, сортируемое поле)
Private Function GetIndexMass(ByRef val, ByRef field%)
    Dim q&, mass()
    ReDim mass(1 To val.Count, 1 To 2)
    For q = 1 To val.Count
        mass(q, 1) = q
        mass(q, 2) = val(q)(field)
    Next
    GetIndexMass = mass
End Function

'Создание коллекции из индексов сортированного массива
Private Function CreateNewCollection(ByRef mass, ByRef val) As Collection
    Dim q&, w&, tmpcoll As New Collection
    For q = LBound(mass) To UBound(mass)
        w = mass(q, 1)
        tmpcoll.Add val(w)
    Next
    Set CreateNewCollection = tmpcoll
End Function

'Сортировка одномерного массива методом выборки
Private Sub SelectionSortSingle(ByRef mass)
    Dim i&, j&, iMin&, tmpval
    For i = LBound(mass) To UBound(mass) - 1
        iMin = i + 1
        For j = i To UBound(mass)
            If set_sort = set_ASC Then If mass(j) < mass(iMin) Then iMin = j
            If set_sort = set_DESC Then If mass(j) > mass(iMin) Then iMin = j
        Next
        If iMin <> i And mass(i) <> mass(iMin) Then
            tmpval = mass(i)
            mass(i) = mass(iMin)
            mass(iMin) = tmpval
        End If
    Next
End Sub

'Сортировка двухмерного массива методом выборки
Private Sub SelectionSortMulti(ByRef mass, ByRef field%)
    Dim i&, j&, iMin&
    For i = LBound(mass) To UBound(mass) - 1
        iMin = i + 1
        For j = i To UBound(mass)
            If set_sort = set_ASC Then If mass(j, field) < mass(iMin, field) Then iMin = j
            If set_sort = set_DESC Then If mass(j, field) > mass(iMin, field) Then iMin = j
        Next
        If iMin <> i And mass(i, field) <> mass(iMin, field) Then
            ReplacingFields mass, iMin, i
        End If
    Next
End Sub
Private Sub ReplacingFields(ByRef mass, ByRef iMin&, ByRef i&)
    Dim q&, tmpval
    For q = LBound(mass, 2) To UBound(mass, 2)
        tmpval = mass(i, q)
        mass(i, q) = mass(iMin, q)
        mass(iMin, q) = tmpval
    Next
End Sub

'Создание нового словаря по отсортированному списку
Private Function CreateNewDictionary(ByRef massKey(), ByRef val) As Object
    Dim i&, obj As Object: Set obj = CreateObject("Scripting.Dictionary")
    For i = LBound(massKey) To UBound(massKey)
        obj.Add key:=massKey(i), Item:=val(massKey(i))
    Next
    Set CreateNewDictionary = obj
End Function
