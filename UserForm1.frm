VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "���i���o�^"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5355
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
        "Data Source=C:\Users\igasa\Desktop\VBA�}�N��\���i�o�^_����\���i���.accdb"
      
    strSQL = "SELECT ���iID FROM ���i��� WHERE ���iID='" _
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
        MsgBox "�o�^�ς݂̏��iID�ł�"
        Exit Sub
      
    Else
        strSQL = "INSERT INTO ���i���(���i��, ���iID, �e��, �l�i, ����, ���l) VALUES('" & TextBox1.Value & "', '" & TextBox2.Value & "', '" _
        & TextBox3.Value & "', '" & TextBox4.Value & "', '" & TextBox5.Value & "', '" & TextBox6.Value & "');"
        With myCmd
            .ActiveConnection = myConn
            .CommandText = strSQL
            .Execute
        End With
    End If
    MsgBox "���i����o�^���܂���"
    
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
