VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        shp.Locked = False
        shp.Visible = False
    
    ActiveWorkbook.Save
    
    Next
End Sub
Private Sub Workbook_Open()
    Application.Wait Now + TimeValue("00:00:13")
    UserForm1.Show vbModeless
End Sub
