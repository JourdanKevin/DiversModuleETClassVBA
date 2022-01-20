Attribute VB_Name = "TCD"
Function TCDchangeRowPlage(nameTCD As String, shTCD As Variant, shSource As Variant)
    shTCD.PivotTables(nameTCD).ChangePivotCache shTCD.Parent.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=shSource.Range("A1:X" & lastRow(shSource)).Address(ReferenceStyle:=xlR1C1, External:=True), Version:=6)
End Function


