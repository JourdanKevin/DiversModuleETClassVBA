Attribute VB_Name = "DebugMod"
Function clog(var As Variant) 'afficher une valeur de variable dans la console, et va choisir comment l'afficher en fonction du type
    Select Case TypeName(var)
        Case "String"
            clogVar (var)
        Case "String()"
            printList (var)
    End Select
End Function
Function clogVar(chaine As String) 'afficher une chaine
    Debug.Print (chaine)
End Function
Function printList(privList As Variant) 'afficher un array
    Dim stringList As String
    stringList = "["
    For Each element In privList
        stringList = stringList & "'" & element & "'" & ","
    Next element
    stringList = Left(stringList, Len(stringList) - 2) & "']"
    Debug.Print stringList
End Function
