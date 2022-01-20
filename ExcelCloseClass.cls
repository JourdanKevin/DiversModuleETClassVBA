VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelCloseClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PrivCn As Variant
Private PrivRst As Variant
Private PrivHDR As Boolean
Private PrivoRs As Variant
Private privSheetName As String
Private Sub class_Initialize()
    Set PrivCn = CreateObject("ADODB.Connection")
    Set PrivRst = CreateObject("ADODB.Recordset")
    PrivHDR = False
End Sub
Property Get Rst() As Variant
    Set Rst = PrivRst
End Property
Property Let HDR(HDRGive As Boolean)
    PrivHDR = HDRGive
End Property
Function Connection(fichier As Variant)
    temp = Replace(";Extended Properties=""Excel 12.0;HDR={value};""", "{value}", HDRValue)
    With PrivCn
       .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" _
            & fichier & temp
        .Open
    End With
End Function
Function SheetName()
    Set PrivoRs = PrivCn.OpenSchema(20)
    privSheetName = PrivoRs.Fields("table_name").value
    SheetName = privSheetName
End Function
Function CloseConnection()
    PrivCn.Close
End Function
Function Execute(Texte_SQL As String)
    Set PrivRst = PrivCn.Execute(Texte_SQL)
End Function
Private Function HDRValue()
    If PrivHDR Then
        HDRValue = "NO"
    Else: HDRValue = "YES"
    End If
End Function
