Attribute VB_Name = "DateMod"
Function GetDate(Optional JourEnMoin As Integer = 0, Optional form As Variant = "YYYY/MM/DD") As String 'renvoie une chaine de caract�re de la date du jour - nombre de jour en moin sour le format soouhaiter (par d�faut "YYYY/MM/DD")
    Select Case (form)
        Case "YYYY/MM/DD" 'format par defaut
            GetDate = Format(Date - JourEnMoin, "YYYY/MM/DD") 'r�cupere la date du jour moin joue en moin
            Exit Function
        Case Else 'sinon
            dateTempArray = GetDateArray(JourEnMoin) '0 = Y, 1 = M, 2 = D 'on r�cupere la date sous une liste o� le premier element est l'ann�, le deuxi�me le mois et enfin le troisi�me le jour
            dateTempString = Replace(form, "YYYY", dateTempArray(0)) 'remplace les ann�e du format souhait� (les YYYY) par l'ann� r�cuperer
            dateTempString = Replace(dateTempString, "MM", dateTempArray(1)) 'ici le mois MM
            GetDate = Replace(dateTempString, "DD", dateTempArray(2)) 'enfin le jour ici JJ
    End Select
End Function
Function replaceDateString(chaine As String, Optional Format As String = "AAAAMMJJ", Optional JourEnMoin As Integer = 0) As Variant 'renvoie la chaine de caractere avec la date souhaiter (date du jour - jour en moin), attention pr�ciser un format si ce n'est pas AAAAMMJJ (par d�faut)
    format2 = Replace(Format, "A", "Y") 'Les dates sont en anglais, pour la conversion, on renplace les A par des Y et J par des D
    format2 = Replace(format2, "J", "D")
    Select Case InStr(Format, "au") 's'il y a un au dans le format alors on veut 2 date (de temps a temps)
        Case 0 'pas de "au"
            replaceDateString = Replace(chaine, Format, GetDate(form:=format2)) 'remplace simplement la chaine avec la date sous le format souhaiter
        Case Else 'il y a un "au"
            tempValue = Split(format2, "au") 'on separe la date avant le au de la date ap�res le au
            dateStart = Replace(tempValue(0), tempValue(0), GetDate(JourEnMoin, tempValue(0))) 'on remplace la date du d�but avant le au
            dateEnd = Replace(tempValue(1), tempValue(1), GetDate(form:=tempValue(1))) ' puis la date de fin apr�s le au
            replaceDateString = Replace(chaine, Format, dateStart & "au" & dateEnd) ' on renvoie la nouvelle chaine en r�assamblant avec les dates
    End Select
End Function
Function DateReplacePointBySlash(value As String) As String 'remplace les "." par "/" utile pour des dates, mais fonctionne avec n'importe quel chaine de caract�re
    DateReplacePointBySlash = Replace(value, ".", "/")
End Function
Function GetDateArray(Optional JourEnMoin As Integer = 0) As Variant
    GetDateArray = Split(GetDate(JourEnMoin), "/") 'genere le tableaux avec la date du jour, le premier element est l'ann�, le deuxi�me le mois et enfin le troisi�me le jour
End Function
Function replaceArrayDateStringDefault(ParamArray Name() As Variant) 'remplace une list de chaine de caractere a remplacer avec la date du jour sour le format par d�faut
    For Each element In Name
        element.Name = replaceDateString(element.Name)
    Next element
End Function

