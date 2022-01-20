Attribute VB_Name = "WordTools"
'Pour utiliser du word penser a activer dans réference : outils>reference>Microsoft Word {version} object library

Function OpenWordDocument(Path As String) As Object

    Set OpenWordDocument = New Word.Application 'créer un objet word
    OpenWordDocument.Documents.Open (Path) 'ouvrir le document word
    
End Function

Function ReplaceWordDocument(OldString As String, NewString As String, Wordoc) 'Wordoc doit être le fichier word

    With Wordoc.Selection.Find 'wordoc.Selection pour selectionner tout le document
        .Text = OldString  'texte a chercher et remplacer
        .Replacement.Text = NewString 'texte a remplacer
        .Forward = True 'mettre le reste des parametre
        .ClearFormatting
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll 'executer le remplacement
    End With

End Function

Function OpenAndReplaceWordDocument(Path As String, OldString As String, NewString As String)

    Set Wordoc = OpenWordDocument(Path)
    Call ReplaceWordDocument(OldString, NewString, Wordoc)
    Set OpenAndReplaceWordDocument = Wordoc
    
End Function


''''''''''''''''''''''''''''''''''''''''''Exemple utilisation remplacer dans word'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Separement()
    Set Wordoc = OpenWordDocument("C:\applocal\exceltoword\Wordtoexcel.docx") 'donner le chemin d'accés au fichier pour l'ouvrir, il vous renverras le fichier word dans une variable, attention mettre Set sur la variable de récupération
    Call ReplaceWordDocument("{remplacer}", "remplacerPar", Wordoc) 'donner quoi remplacer, puis par quoi et enfin la variable retourner par OpenWordDocument
    Call ReplaceWordDocument("{autre}", "batman", Wordoc) 'donner quoi remplacer, puis par quoi et enfin la variable retourner par OpenWordDocument
    
End Sub

Sub Direct()

    Set Wordoc = OpenAndReplaceWordDocument("C:\applocal\exceltoword\Wordtoexcel.docx", "{remplacer}", "remplacerPar") 'chemin du document, quoi remplacer, par quoi

End Sub

