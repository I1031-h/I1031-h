VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "商品情報登録"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5355
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim myConn As New ADODB.Connection
    Dim myCmd As New ADODB.Command
    Dim myRS As New ADODB.Recordset
    Dim strSQL As String
    
    myConn.Open ConnectionString:="Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Users\igasa\Desktop\VBAマクロ\商品登録_操作\商品情報.accdb"
      
    strSQL = "SELECT 商品ID FROM 商品情報 WHERE 商品ID='" _
        & TextBox2.Value & "';"

    With myCmd
        .ActiveConnection = myConn
        .CommandText = strSQL
        Set myRS = .Execute
    End With
    
    If myRS.EOF = False Then
        TextBox1 = ""
        TextBox2 = ""
        TextBox3 = ""
        TextBox4 = ""
        TextBox5 = ""
        TextBox6 = ""
        MsgBox "登録済みの商品IDです"
        Exit Sub
      
    Else
        strSQL = "INSERT INTO 商品情報(商品名, 商品ID, 容量, 値段, 分類, 備考) VALUES('" & TextBox1.Value & "', '" & TextBox2.Value & "', '" _
        & TextBox3.Value & "', '" & TextBox4.Value & "', '" & TextBox5.Value & "', '" & TextBox6.Value & "');"
        With myCmd
            .ActiveConnection = myConn
            .CommandText = strSQL
            .Execute
        End With
    End If
    MsgBox "商品情報を登録しました"
    
    TextBox1 = ""
    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
    TextBox5 = ""
    TextBox6 = ""
    
    Set myCmd = Nothing
    myConn.Close: Set myConn = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim shp As Shape

For Each shp In ActiveSheet.Shapes
    shp.Locked = False
    shp.Visible = True
    Next
End Sub
