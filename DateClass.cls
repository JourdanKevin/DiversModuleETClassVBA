VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PrivWorkDate As Variant
Private Sub class_Initialize()
    PrivWorkDate = CDate(GetDate)
End Sub
Property Get WorkDate() As String
    WorkDate = PrivWorkDate
End Property
Property Let WorkDate(WorkDateGive As String)
    If WorkDateGive <> "" Then
        PrivWorkDate = CDate(WorkDateGive)
    End If
End Property
Function ReplaceForm(ByVal form As Variant) As String
     form = Replace(form, "A", "Y") 'Les dates sont en anglais, pour la conversion, on renplace les A par des Y et J par des D
     ReplaceForm = Replace(form, "J", "D")
End Function
Function replaceDateString(chaine As String, Optional formGive As String = "AAAAMMJJ", Optional JourEnMoin As Integer = 0) As Variant 'renvoie la chaine de caractere avec la date souhaiter (date du jour - jour en moin), attention préciser un format si ce n'est pas AAAAMMJJ (par défaut)
    form = ReplaceForm(formGive)
    Select Case InStr(formGive, "au") 's'il y a un au dans le format alors on veut 2 date (de temps a temps)
        Case 0 'pas de "au"
            replaceDateString = Replace(chaine, formGive, GetDate(JourEnMoin, form)) 'remplace simplement la chaine avec la date sous le format souhaiter
        Case Else 'il y a un "au"
            tempValue = Split(form, "au") 'on separe la date avant le au de la date apéres le au
            dateStart = Replace(tempValue(0), tempValue(0), GetDate(JourEnMoin, tempValue(0))) 'on remplace la date du début avant le au
            dateEnd = Replace(tempValue(1), tempValue(1), GetDate(form:=tempValue(1))) ' puis la date de fin aprés le au
            replaceDateString = Replace(chaine, formGive, dateStart & "au" & dateEnd) ' on renvoie la nouvelle chaine en réassamblant avec les dates
    End Select
End Function
Function replaceArrayDateStringDefault(ParamArray Name() As Variant) 'remplace une list de chaine de caractere a remplacer avec la date du jour sour le format par défaut
    For Each element In Name
        element.Name = replaceDateString(element.Name)
    Next element
End Function
Function GetDate(Optional JourEnMoin As Integer = 0, Optional form As Variant = "YYYY/MM/DD") As Variant
    form = ReplaceForm(form)
    Select Case PrivWorkDate
        Case Is <> ""
            GetDate = Format(PrivWorkDate - JourEnMoin, form)
        Case Else
            GetDate = Format(Date - JourEnMoin, form)
        End Select
End Function
Function GetDateArray(Optional delimiter As String = "/", Optional form As Variant = "YYYY/MM/DD", Optional JourEnMoin As Integer = 0) As Variant
    GetDateArray = Split(GetDate(JourEnMoin, form), delimiter)
End Function
