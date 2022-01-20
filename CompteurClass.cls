VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CompteurClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private privCompteur As Integer
Private Sub class_Initialize()
    privCompteur = 0
End Sub
Property Get compteur() As Integer
    compteur = privCompteur
End Property
Property Let compteur(GetCompt As Integer)
    privCompteur = GetCompt
End Property
Function reset(Optional Val As Integer = 0)
    privCompteur = Val
End Function
Function UpGet(Optional pas As Integer = 1)
    privCompteur = privCompteur + pas
    UpGet = compteur
End Function
Function GetUp(Optional pas As Integer = 1)
    GetUp = compteur
    privCompteur = privCompteur + pas
End Function
Function DownGet(Optional pas As Integer = 1)
    privCompteur = privCompteur - pas
    DownGet = compteur
End Function
Function GetDown(Optional pas As Integer = 1)
    GetDown = compteur
    privCompteur = privCompteur - pas
End Function
