Attribute VB_Name = "Module1"
Sub 再表示()
    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        shp.Locked = False
        shp.Visible = False
    Next
    
    UserForm1.Show vbModeless
End Sub
Sub データ抽出()
    Dim myConn As New ADODB.Connection
    Dim myCmd As New ADODB.Command
    Dim myRS As New ADODB.Recordset
    Dim strSQL As String
    
    ActiveSheet.Unprotect
    
    myConn.Open ConnectionString:="Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Users\igasa\Desktop\VBAマクロ\商品登録_操作\商品情報.accdb"
    
    strSQL = "SELECT ID,商品ID,商品名,容量,値段,分類,備考 FROM 商品情報"
    
    myRS.Open strSQL, myConn
    
    Worksheets("データ操作").Range("A3").CopyFromRecordset myRS
    myRS.Close
    myConn.Close
    
    'Range("A1:C7").Locked = True
    Range("B3", Cells(Rows.Count, 2).End(xlUp)).Locked = False
    Range("C3", Cells(Rows.Count, 3).End(xlUp)).Locked = False
    Range("D3", Cells(Rows.Count, 4).End(xlUp)).Locked = False
    Range("E3", Cells(Rows.Count, 5).End(xlUp)).Locked = False
    Range("F3", Cells(Rows.Count, 6).End(xlUp)).Locked = False
    Range("G3", Cells(Rows.Count, 7).End(xlUp)).Locked = False
    
    
    Range("A2", Cells(Rows.Count, 1).End(xlUp)).Locked = True
    ActiveSheet.Protect
    
    MsgBox "データを抽出しました"
    
    Set myCmd = Nothing
    Set myConn = Nothing
    
End Sub
Sub データベース更新()
    Dim myConn As New ADODB.Connection
    Dim myCmd As New ADODB.Command
    Dim myRS As New ADODB.Recordset
    Dim strSQL As String
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = Worksheets("データ操作")
    
    ActiveSheet.Unprotect
    
    myConn.Open ConnectionString:="Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Users\igasa\Desktop\VBAマクロ\商品登録_操作\商品情報.accdb"
    
    i = 3
    Do While ws.Cells(i, 1) <> ""
           
       With myCmd
       
        strSQL = "UPDATE 商品情報 SET 商品ID= '" & ws.Cells(i, 2) & "', 商品名= '" & ws.Cells(i, 3) & "', 容量= '" _
         & ws.Cells(i, 4) & "', 値段= '" & ws.Cells(i, 5) & "', 分類= '" & ws.Cells(i, 6) & "', 備考= '" & ws.Cells(i, 7) & "' WHERE ID= " & ws.Cells(i, 1) & ""
        '数値の場合は「'」でくくる必要がない（文字列は必要）
        
        .ActiveConnection = myConn
        .CommandText = strSQL
        .Execute
        
        End With
        
    i = i + 1
    
    Loop
    
    MsgBox "データベースを更新しました"
    
    Range("B3", Cells(Rows.Count, 2).End(xlUp)).Locked = False
    Range("C3", Cells(Rows.Count, 3).End(xlUp)).Locked = False
    Range("D3", Cells(Rows.Count, 4).End(xlUp)).Locked = False
    Range("E3", Cells(Rows.Count, 5).End(xlUp)).Locked = False
    Range("F3", Cells(Rows.Count, 6).End(xlUp)).Locked = False
    Range("G3", Cells(Rows.Count, 7).End(xlUp)).Locked = False
    
    Range("A2", Cells(Rows.Count, 1).End(xlUp)).Locked = True
    ActiveSheet.Protect
    
    Set myCmd = Nothing
    myConn.Close: Set myConn = Nothing
        
End Sub
Sub 保護解除()
    ActiveSheet.Unprotect
End Sub
Sub データ削除()
    UserForm2.Show vbModeless
End Sub

