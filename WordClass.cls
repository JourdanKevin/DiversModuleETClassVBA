VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''ATTRIBUS PIVATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private PrivPathOpen, PrivPathSave As String
Private Privdoc, PrivFormat As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''CONSTANTE DE L'OBJET''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const MsgNouvelObj As String = "Si vous voulez travaillez sur un autre document veuillez : 1. supprimer l'objet ou SaveAndClose (Attention certe votre fichier sera sauvegarder, mais pas encore supprimer ; 2. creer un nouvel objet : Dim Words As new WordClass"
Private Const MsgRemplacerPathManquant As String = "Veuillez definir un pathOpen afin de permettre de savoir sur quel fichier faut-il remplacer"
Private Const MsgAucunCheminSave As String = "Attention aucun chemin de sauvegarde n'a était donnée, NameObject.PathSave = pathsave"
Private Const MsgInsert As String = "Aucun document word n'a était ouvert ou créer, soit remplisser le path, soit créer un document word"
Private Const MsgDontCreate As String = "Vous ne pouvez pas créer de document quand vous initialiser un path pour ouvrir un word existant (on ne créer un document word si on en ouvre un autre), si vous voulez créer un nouveau document word veuillez créer un nouvelle objet"
Private Const FormatUnknown As String = "Un des formats n'est pas reconue, seul docx et pdf le son pour le moment"
Private Const MsgAucunFormat As String = "Aucun Format parametrer, veuillez renseigner un format soit a la fin du pathSave avec .docx par exemple ; soit renseigner sur le parametre format = 'docx,pdf'"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''VARIABLE INITIALISER LORS DE LA CREATION''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private PrivWordoc As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''INITIALISATION''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub class_Initialize()
    Set PrivWordoc = New Word.Application 'créer un objet word
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''PROPERTY''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Let PathOpen(pathGive As String)
    If PrivPathOpen = "" Then 'Seulement le premier path que l'on donnera seras pris en compte et ne pourras plus être modifier (on travaille sur un seul word par objet), pour en ouvrir un autre recréer un nouvelle objet
        PrivPathOpen = pathGive
        OpenWordDocument
    Else
        MsgBox MsgNouvelObj
    End If
End Property
Property Get PathOpen() As String
    Path = PrivPathOpen
End Property
Property Let PathSave(pathGive As String)
    If PrivPathSave <> "" And pathGive <> PrivPathSave And PrivPathOpen <> "" Then
        If Privdoc <> "" Then
            SaveAndClose
        End If
        OpenWordDocument
    End If
    PrivPathSave = pathGive
End Property
Property Get PathSave() As String
    PathSave = PrivPathSave
End Property
Property Let Format(FormatGive As String)
    PrivFormat = FormatGive
End Property
Property Get Format() As String
    Format = PrivFormat
End Property
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''PRIVATE FUNCTION''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OpenWordDocument()
    Set Privdoc = PrivWordoc.Documents.Open(PrivPathOpen)  'ouvrir le document word
End Function
Private Function CloseWordDocument()
    If Privdoc <> "" Then
        Privdoc.Close wdDoNotSaveChanges
        Privdoc = ""
    End If
End Function
Private Function MsgInsertError()
    MsgBox MsgInsert
End Function
Private Function SelectCaseSave(PathCase As String)
    Select Case Split(PathCase, ".")(1)
        Case "pdf"
            Privdoc.ExportAsFixedFormat OutputFileName:=PrivPathSave, ExportFormat:=wdExportFormatPDF
        Case "docx"
            Privdoc.SaveAs (PrivPathSave)
        Case Else
            MsgBox FormatUnknown
    End Select
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''PUBLIC FUNCTION''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateNewWord()
    If PathOpen = "" Then
        Set Privdoc = PrivWordoc.Documents.add  'ouvrir le document word
    Else
        MsgBox MsgDontCreate
    End If
End Function
Public Function ReplaceWordDocument(OldString As String, NewString As String)
    If Privdoc <> "" Then
        With PrivWordoc.Selection.Find 'wordoc.Selection pour selectionner tout le document
            .Text = OldString  'texte a chercher et remplacer
            .Replacement.Text = NewString 'texte a remplacer
            .Execute Replace:=wdReplaceAll 'executer le remplacement sur toutes les occurences
        End With
    Else
        MsgBox MsgRemplacerPathManquant
    End If
End Function
Public Function InsertEnd(Text As String)
    If Privdoc <> "" Then
        Privdoc.Content.InsertAfter Text
    Else
        MsgInsertError
    End If
    
End Function
Public Function InsertStart(Text As String)
    If PrivWordoc <> "" Then
        PrivWordoc.Selection.InsertBefore Text
    Else
        MsgInsertError
    End If
End Function
Public Function Save()
    If PrivPathSave <> "" Then
        If UBound(Split(PrivPathSave, ".")) - LBound(Split(PrivPathSave, ".")) + 1 > 1 Then
            SelectCaseSave (PrivPathSave)
        ElseIf PrivFormat <> "" Then
            For Each FormatElement In Split(PrivFormat, ",")
                SelectCaseSave (PrivPathSave & "." & Trim(FormatElement))
            Next FormatElement
        Else
            MsgBox MsgAucunFormat
        End If
    Else
        MsgBox MsgAucunCheminSave
    End If
End Function
Public Function SaveAndClose()
    Save
    CloseWordDocument
End Function

