VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegexClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private regex As Object
Private PrivPattern As String
Private Sub class_Initialize()
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
End Sub
Property Let regPattern(PatternGive As String)
    PrivPattern = PatternGive
End Property
Property Get regPattern() As String
    regPattern = PrivPattern
End Property
Private Function setPattern(Optional regPat As String)
    If regPat <> "" Then
        regex.Pattern = regPat
    Else
        regex.Pattern = PrivPattern
    End If
End Function
Private Function setGlobal(Optional regGlob As Boolean = True)
    regex.Global = regGlob
End Function
Public Function Match(chaine As String, Optional regPat As String, Optional regGlob As Boolean = True) As Variant
    setPattern regPat
    setGlobal regGlob
    Set Match = regex.Execute(chaine)
End Function
Public Function Sim(chaine As String, Optional regPat As String, Optional regGlob As Boolean = True) As Variant
    setPattern regPat
    setGlobal regGlob
    Sim = regex.Test(chaine)
End Function


