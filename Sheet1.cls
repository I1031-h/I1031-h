VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        shp.Locked = False
        shp.Visible = False
    Next
        
        UserForm1.Show vbModeless
End Sub
