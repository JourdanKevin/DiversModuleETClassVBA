VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private privList As Variant
Private privCurrentValue As Variant
Private length As Integer
Private compteur As Variant
Private Sub class_Initialize()
    Set compteur = New CompteurClass
End Sub
Property Get GetList() As Variant
    GetList = privList
End Property
Property Get Count() As Integer
    Count = length
End Property
Property Get CurrentValue()
    CurrentValue = privCurrentValue
End Property
Function List(ParamArray ListGive() As Variant)
    If TypeName(ListGive(0)) = "Variant()" Then
        privList = ListGive(0)
    Else
        privList = ListGive()
    End If
    length = CountLength
End Function
Private Function CountLength()
    CountLength = UBound(privList) - LBound(privList) + 1
End Function
Function NextVal(Optional pas As Integer = 1) As Variant
    With compteur
        If .compteur < length Then
            privCurrentValue = privList(.UpGet(pas))
        Else
            privCurrentValue = False
            .reset
        End If
        NextVal = privCurrentValue
    End With
End Function
Function ValNext(Optional pas As Integer = 1) As Variant
    With compteur
        If .compteur < length Then
            privCurrentValue = privList(.GetUp(pas))
        Else
            privCurrentValue = False
            .reset
        End If
        NextVal = privCurrentValue
        privCurrentValue = privList(.compteur)
    End With
    
End Function
Function LastVal(Optional pas As Integer = 1) As Variant
    If compteur > 0 Then
        privCurrentValue = privList(compteur)
        incremente (-pas)
    Else
        privCurrentValue = False
        ResetCompteur (lenght - 1)
    LastVal = privCurrentValue
End Function
Function ValLast(Optional pas As Integer = 1) As Variant
    With compteur
        If .compteur > 0 Then
            privCurrentValue = privList(compteur)
            incremente (-pas)
        Else
            privCurrentValue = False
            ResetCompteur (lenght - 1)
        LastVal = privCurrentValue
    End With
End Function
Function reset(Optional index As Integer = 0)
    compteur.reset (index)
End Function
Function incremente(Optional pas As Integer = 1)
    compteur.GetUp (pas)
End Function
Function AddElement(value)
    ReDim Preserve privList(UBound(privList) + 1) ' Redimension:
    privList(UBound(privList)) = value ' Fill last element
End Function
'End Function
'Function DelElement(ParamArray Text())
'    printList Text()
'End Function
'Function DelLastElement(ParamArray Text())
'    printList Text()
'End Function
'Function DelFirstElement(ParamArray Text())
'    printList Text()
'End Function
'Function DelIdElement(ParamArray Text())
'    printList Text()
'End Function
