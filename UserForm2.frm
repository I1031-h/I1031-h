VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   1110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3495
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim myConn As New ADODB.Connection
    Dim myCmd As New ADODB.Command
    Dim myRS As New ADODB.Recordset
    Dim strSQL As String
    Dim ws As Worksheet
    Dim a As Long
    Dim textb As String
    Dim textb1 As String
    Dim ProductName As String
    
    Set ws = Worksheets("データ操作")
    textb = TextBox1.Value
    textb1 = Range("A3", Cells(Rows.Count, 1).End(xlUp)).Find(textb).Row
    
    ProductName = Range("C" & textb1).Value
    
    a = MsgBox("商品名:" & ProductName & "のデータを削除しますか？", vbYesNo)
    
    If a = vbYes Then
    
    Sheets("データ操作").Unprotect
    
    myConn.Open ConnectionString:="Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Users\igasa\Desktop\VBAマクロ\商品登録_操作\商品情報.accdb"
        
    strSQL = "DELETE FROM 商品情報 WHERE ID=" & TextBox1.Value
    myConn.Execute strSQL
    
    ws.Range("A" & textb1).Select
    Selection.EntireRow.Delete Shift:=xlUp
    
    Unload UserForm2
    MsgBox "商品名:" & ProductName & "のデータを削除しました"
    
    Set myCmd = Nothing
    myConn.Close: Set myConn = Nothing
    
    Else
    Unload UserForm2
    
    End If
    
    ActiveSheet.Protect
    
End Sub

Private Sub UserForm_Click()

End Sub
