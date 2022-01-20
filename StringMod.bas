Attribute VB_Name = "StringMod"
Function DelCharStartAndLastDelCharEnd(chaine, Optional numberFirst As Integer = 1, Optional numberLast As Integer = 1) As String ''supprime un nombre de caractére au début, puis un autre nombre a la fin d'une chaine de caractère
    DelCharStartAndLastDelCharEnd = DelCharEnd(DelCharStart(chaine, numberFirst), numberLast)
End Function
Function DelCharStart(chaine, Optional number As Integer = 1) As String 'supprime nb caractère au début d'une chaine d'une chaine
     DelCharStart = Right(chaine, Len(chaine) - number)
End Function
Function DelCharEnd(chaine, Optional number As Integer = 1) As String 'supprime nb caractère a la fin d'une chaine
    DelCharEnd = Left(chaine, Len(chaine) - number)
End Function
Function ReplaceCharEnd(chaine As String, replaceBy As String, Optional number As Integer = 1) '' remplace un certain nombre de caractére a la fin de la chaine par d'autre
    ReplaceCharEnd = DelCharEnd(chaine, number) & replaceBy
End Function
Function SplitCharAtNumberChar(chaine, nb) As Variant '' sépare la chaine en 2 au numéro de caractère donné
    SplitCharAtNumberChar = Array(GetNbCharStart(chaine, nb), DelCharStart(chaine, nb))
End Function
Function GetNbCharStart(chaine, nb) As String '' récupére les n premiers caractére d'une chaine
    GetNbCharStart = Left(chaine, nb)
End Function
Function GetNbCharEnd(chaine, nb) As String '' récupére les n derniers caractére d'une chaine donné
    GetNbCharEnd = Right(chaine, nb)
End Function
Function InsertAtNbChar(chaine, add, nb) As String '' insert une chaine a l'interieur d'une autre chaine au numéro de char donné
    tChaineTemp = SplitCharAtNumberChar(chaine, nb)
    GetNbCharEnd = tChaineTemp(0) & add & tChaineTemp(1)
End Function
Function pPath(c As String) As String
    pPath = Mid(c, 1, InStrRev(c, "\") - 1)
End Function


