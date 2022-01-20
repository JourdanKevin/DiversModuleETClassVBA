VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GlobClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private privVal As New Collection
Property Get Val() As Variant
    Val = privVal
End Property
Function CreateVar(item As Variant, Optional key As Variant)
    If TypeName(var) = "Empty" Then
        privVal.add item
    Else
        privVal.add item, key
    End If
End Function
Function GetVal(key As Variant) As Variant
    GetVal = privVal(key)
End Function
Function delete(key As Variant)
    privVal.Remove (key)
End Function
