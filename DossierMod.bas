Attribute VB_Name = "Dossier"
Function CreerDossier(chemin As String)
'par: Excel-Malin.com ( https://excel-malin.com )
    On Error GoTo CreerDossierErreur
    
    Dim PremierDossier As String
    Dim CheminReseau As Boolean
    Dim CheminPartielOK As String
    Dim CheminPartiel, PartieDeChemin As Integer
    Dim PartiesDeChemin As Variant
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Len(Dir(chemin, vbDirectory)) > 0 Then
    CreerDossier = True
    Exit Function
    Else
            'suppression du dernier backslash si présent
            If Right(chemin, 1) = Application.PathSeparator Then chemin = Left(chemin, Len(chemin) - 1)
            
            'vérificacion si chemin local ou réseau
            If Left(chemin, 2) = "\\" Then
                CheminReseau = True
            Else
                CheminReseau = False
            End If
            
            'décomposition du chemin
            If CheminReseau = False Then
                PartiesDeChemin = Split(chemin, Application.PathSeparator)
                CheminPartielOK = ""
                PremierDossier = LBound(PartiesDeChemin)
            Else
                PartiesDeChemin = Split(Replace(chemin, "\\", ""), Application.PathSeparator)
                CheminPartielOK = ""
                PremierDossier = LBound(PartiesDeChemin) + 1
            End If
        
        'tests et créations de (sous)dossiers
            For PartieDeChemin = PremierDossier To UBound(PartiesDeChemin)
    
                For CheminPartiel = LBound(PartiesDeChemin) To PartieDeChemin
                
                            If CheminReseau = False Then
                                CheminPartielOK = CheminPartielOK & PartiesDeChemin(CheminPartiel) & Application.PathSeparator
                            Else
                                CheminPartielOK = CheminPartielOK & PartiesDeChemin(CheminPartiel) & Application.PathSeparator
                            End If
    
                    If CheminPartiel = PartieDeChemin Then
                            If CheminReseau = False Then
                                        If FSO.FolderExists(CheminPartielOK) = False Then
                                                MkDir CheminPartielOK
                                        End If
                            Else
                                        If Right(CheminPartielOK, 1) = Application.PathSeparator Then _
                                        CheminPartielOK = Left(CheminPartielOK, Len(CheminPartielOK) - 1)
                                        
                                        If Left(CheminPartielOK, 2) <> "\\" Then _
                                        CheminPartielOK = "\\" & CheminPartielOK
                                        
                                        If FSO.FolderExists(CheminPartielOK) = False Then
                                                MkDir CheminPartielOK
                                        End If
                            End If
                    End If
                Next CheminPartiel
                CheminPartielOK = ""
            Next PartieDeChemin
    End If
    
    CreerDossier = True
    Exit Function
CreerDossierErreur:
    CreerDossier = False
End Function


