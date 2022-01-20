Attribute VB_Name = "StringMod"
Function DelCharStartAndLastDelCharEnd(chaine, Optional numberFirst As Integer = 1, Optional numberLast As Integer = 1) As String ''supprime un nombre de caract�re au d�but, puis un autre nombre a la fin d'une chaine de caract�re
    DelCharStartAndLastDelCharEnd = DelCharEnd(DelCharStart(chaine, numberFirst), numberLast)
End Function
Function DelCharStart(chaine, Optional number As Integer = 1) As String 'supprime nb caract�re au d�but d'une chaine d'une chaine
     DelCharStart = Right(chaine, Len(chaine) - number)
End Function
Function DelCharEnd(chaine, Optional number As Integer = 1) As String 'supprime nb caract�re a la fin d'une chaine
    DelCharEnd = Left(chaine, Len(chaine) - number)
End Function
Function ReplaceCharEnd(chaine As String, replaceBy As String, Optional number As Integer = 1) '' remplace un certain nombre de caract�re a la fin de la chaine par d'autre
    ReplaceCharEnd = DelCharEnd(chaine, number) & replaceBy
End Function
Function SplitCharAtNumberChar(chaine, nb) As Variant '' s�pare la chaine en 2 au num�ro de caract�re donn�
    SplitCharAtNumberChar = Array(GetNbCharStart(chaine, nb), DelCharStart(chaine, nb))
End Function
Function GetNbCharStart(chaine, nb) As String '' r�cup�re les n premiers caract�re d'une chaine
    GetNbCharStart = Left(chaine, nb)
End Function
Function GetNbCharEnd(chaine, nb) As String '' r�cup�re les n derniers caract�re d'une chaine donn�
    GetNbCharEnd = Right(chaine, nb)
End Function
Function InsertAtNbChar(chaine, add, nb) As String '' insert une chaine a l'interieur d'une autre chaine au num�ro de char donn�
    tChaineTemp = SplitCharAtNumberChar(chaine, nb)
    GetNbCharEnd = tChaineTemp(0) & add & tChaineTemp(1)
End Function
Function pPath(c As String) As String
    pPath = Mid(c, 1, InStrRev(c, "\") - 1)
End Function


