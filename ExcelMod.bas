Attribute VB_Name = "ExcelMod"
Function lastRow(fichier As Variant, Optional Col As Variant = 1) As String 'donne la dernière ligne d'une feuille sur une colonne (si pas de colonne préciser, prend la première colone en valeur)
    lastRow = fichier.Cells(Rows.Count, Col).End(xlUp).Row
End Function
Function lastCol(fichier As Variant, ligne As Long) As String 'meme chose que lastRow mais pour les colonne
    lastCol = Split(Columns(fichier.Cells(ligne, Columns.Count).End(xlToLeft).Column).Address(ColumnAbsolute:=False), ":")(1)
End Function
Function FirstNoBlank(shRange As Variant) As Variant 'retourne la premiere valeur et ligne non blank (vide) dans une plage de cellule
    Dim Found As Boolean
    Found = False
    For Each Cell In shRange
        If Not IsEmpty(Cell) Then
            FirstNoBlank = Array(Cell.Row, Cell)
            Found = True
            Exit For
        End If
    Next Cell
    If Not Found Then
        FirstNoBlank = Array(False)
    End If
End Function
Function CopySheets(wkToCopy As Variant, chemin As Variant, Optional sheetsTarget As Variant = 1) As String 'copie une feuille ("par defaut la première, mais on peut spécifier son nom ou son numeroe) d'un classeur sur un autre classeur (met la feuille en derniére position) et retourne le nom de cette feuille (pour la rentrer dans une variable par exemple)
    Application.ScreenUpdating = False 'évite certains effet visuel lors de l'ouverture d'un workbook par exemple pour ameliorer les performance
    Set wkAcopier = Workbooks.Open(chemin) 'ouverture du classeur
    With wkAcopier
        .Sheets(sheetsTarget).Copy After:=wkToCopy.Sheets(wkToCopy.Sheets.Count) 'on copie la feuille en derniére postion
        CopySheets = .Sheets(sheetsTarget).Name 'on renvoie le nom de la feuille copier
        .Close SaveChanges:=False 'on ferme le classeur sans sauvergarder
    End With
End Function
Function DeleteSheets(ParamArray tSheets() As Variant) 'supprime une liste de feuille a supprimer passer les variable en parametre
    Application.DisplayAlerts = False 'empeche les alerte demandant demandant confirmation de suppression
    For Each Sheet In tSheets
        Sheet.delete 'supprime chaque feuille une a une jusqu'a la fin de notre liste
    Next Sheet
    Application.DisplayAlerts = True 'on remet les alertes
End Function
Function CopyRangeOfCell(rangeToCopy As Variant, rangeCopyTo As Variant, Optional speCop As String) 'copier une plage de cellule d'une feuille vers une autre feuille : rangeToCopy = plage copier ; rangeCopyTo = plage coller
    Select Case speCop
        Case "values" 'si values est passer en parametre seul les valeurs seront copier et non les formules
            rangeToCopy.Copy 'equivalent a crtl+C de la plage
            rangeCopyTo.PasteSpecial Paste:=xlPasteValues 'equivalent a coller uniquement les valeur
            Application.CutCopyMode = False 'deselection l'effet de copy (comme celui du ctrl+V)
        Case Else
            rangeToCopy.Copy Destination:=rangeCopyTo 'copy brut de la feuille que l'on desire copier a la feuille que l'on desire coller
    End Select
End Function
Function FileExist(File As String) As Boolean 'Verifie si u fichier existe en lui donnant en parametre le chemin avec nom du fichier
    If File <> "" And Len(Dir(File)) > 0 Then 'la condition qui permet de vérifier
        FileExist = True ' return True, il existe
    Else
        FileExist = False ' return False, N'existe PAS
    End If
End Function
Public Function VerifFileAndExitWithMsgBox(File As Variant) As Boolean 'Liste des fichier a vérifier l'existance, et va retourner True si tous existe ou False avec un message box de tous les fichier manquant
    Existe = True 'on initialise l'existance a vrai
    textError = "Fichier Manquant : " 'on iniatilise le message listant les fichier manquant
    For Each fichier In File ' on boucle sur notre liste contenant tout les fichier a verifier
        If Not FileExist(CStr(fichier)) Then 'utilise la fonction de verification et si return False alors
            textError = textError & fichier & "    " 'on rajoute le nom du fichier manquant au message
            Existe = False 'on passe en l'existance de tout les fichier a faux
        End If
    Next fichier
    If Not Existe Then 'Si l'existance de tout les fichier est a faux alors
        MsgBox textError 'afficher le msg des fichier manquant dans un msgBox (popup)
    End If
    VerifFileAndExitWithMsgBox = Existe 'return False si il manque un fichier ou vrai si tous les fichiers sont présent
End Function
Function GetLastRowAndReturnMajor(sh, ParamArray Col() As Variant) As Integer 'Récupere n colonne et renvoie la derniere ligne avec la valeur la plus élever
    last1 = 0
    For Each value In Col
        last2 = lastRow(sh, value)
        If last1 < last2 Then
            last1 = last2
        End If
    Next value
    GetLastRowAndReturnMajor = last1
End Function
Function CopyRangeOfCellArray(ParamArray toCopy() As Variant) 'recupere une liste comportant des liste avec les donnée necessaire a CopyRangeOfCell Array(rangeToCopy As Variant, rangeCopyTo As Variant), element(0) la variable de la feuille où copier, element(1) la colonne, element(2) la feuille où coller, element(3) la colonne où coller ; Il va prendre jusqu'as la dernièreligne dans les 2 cas
    For Each element In toCopy
        CopyRangeOfCell element(0).Range(element(1) & lastRow(element(0))), element(2).Range(element(3) & lastRow(element(2)) + 1)
    Next element
End Function

