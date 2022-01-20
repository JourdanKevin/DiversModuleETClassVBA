Attribute VB_Name = "ArrayMod"
Function Count(arr) As Integer 'retourne le nombre d'element d'une liste
    Count = LastI(arr) + 1
End Function
Function LastI(arr As Variant) As Integer 'revoie le dernier index d'une liste
    LastI = UBound(arr) - LBound(arr)
End Function
Function ArrayAdd(arr As Variant, value As Variant)
    ReDim Preserve arr(UBound(arr) + 1) ' Redimension:
    arr(UBound(arr)) = value ' Fill last element
    ArrayAdd = arr
End Function
Function ArrayAddIfNotExist(arr As Variant, value As Variant)
    If Not IsInArray(arr, value) Then
        arr = ArrayAdd(arr, value)
    End If
    ArrayAddIfNotExist = arr
End Function
Function IsInArray(arr As Variant, value As Variant)
    IsInArray = Not IsError(Application.Match(value, arr, 0))
End Function
Function ArrayRemoveDuplicates(arr As Variant)
    Set SheetTmp = ActiveWorkbook.Sheets.add
    SheetTmp.Range(Cells(1, 1), Cells(Count(arr), 1)) = ArrayToRange(arr)
    LastR = lastRow(SheetTmp)
    SheetTmp.Range("A1:A" & LastR).RemoveDuplicates Columns:=1, Header:=xlGuess
    ArrayRemoveDuplicates = RangeToArray(SheetTmp.Range("A1:A" & lastRow(SheetTmp)))
    DeleteSheets SheetTmp
End Function
Function ArrayToRange(arr As Variant)
    ArrayToRange = Application.WorksheetFunction.Transpose(arr)
End Function
Public Function RangeToArray(rng As Range) As Variant()
    RangeToArray = Array()
    For Each element In rng
        RangeToArray = ArrayAd(RangeToArray, element)
    Next element
End Function
Sub testConvert()
    arr = ArrayAddIfNotExist(Array("1", "2", "3"), "4")
End Sub
