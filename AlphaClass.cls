VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AlphaClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PrivMajAlpha As Variant
Private PrivMinAlpha As Variant
Private PrivNum As Variant
Private PrivCurrentList As String
Private PrivCurrentVal As String
Private Collec As New Collection
Private Sub class_Initialize()
    Set PrivMajAlpha = New ListClass
    Set PrivMinAlpha = New ListClass
    Set PrivNum = New ListClass
    PrivMajAlpha.List "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
    PrivMinAlpha.List "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
    PrivNum.List "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
    Collec.add PrivMajAlpha, "maj"
    Collec.add PrivMinAlpha, "min"
    Collec.add PrivNum, "num"
End Sub
Property Get List(nameList As String)
    List = Collec(nameList).GetList
End Property
Property Let CurrentList(nameList As String)
    PrivCurrentList = nameList
End Property
Property Get CurrentVal()
    CurrentVal = PrivCurrentVal
End Property
Function NextVal(Optional pas As Integer = 1) As Variant
    PrivCurrentVal = Collec(PrivCurrentList).NextVal(pas)
    NextVal = PrivCurrentVal
End Function
Function LastVal(Optional pas As Integer = 1) As Variant
    PrivCurrentVal = Collec(PrivCurrentList).LastVal(pas)
    NextVal = PrivCurrentVal
End Function
Function ResetCompteur(Optional index As Integer = 0)
    Collec(PrivCurrentList).reset (index)
End Function

