Attribute VB_Name = "Module1"
Sub �ĕ\��()
    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        shp.Locked = False
        shp.Visible = False
    Next
    
    UserForm1.Show vbModeless
End Sub
Sub �f�[�^���o()
    Dim myConn As New ADODB.Connection
    Dim myCmd As New ADODB.Command
    Dim myRS As New ADODB.Recordset
    Dim strSQL As String
    
    ActiveSheet.Unprotect
    
    myConn.Open ConnectionString:="Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Users\igasa\Desktop\VBA�}�N��\���i�o�^_����\���i���.accdb"
    
    strSQL = "SELECT ID,���iID,���i��,�e��,�l�i,����,���l FROM ���i���"
    
    myRS.Open strSQL, myConn
    
    Worksheets("�f�[�^����").Range("A3").CopyFromRecordset myRS
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
    
    MsgBox "�f�[�^�𒊏o���܂���"
    
    Set myCmd = Nothing
    Set myConn = Nothing
    
End Sub
Sub �f�[�^�x�[�X�X�V()
    Dim myConn As New ADODB.Connection
    Dim myCmd As New ADODB.Command
    Dim myRS As New ADODB.Recordset
    Dim strSQL As String
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = Worksheets("�f�[�^����")
    
    ActiveSheet.Unprotect
    
    myConn.Open ConnectionString:="Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Users\igasa\Desktop\VBA�}�N��\���i�o�^_����\���i���.accdb"
    
    i = 3
    Do While ws.Cells(i, 1) <> ""
           
       With myCmd
       
        strSQL = "UPDATE ���i��� SET ���iID= '" & ws.Cells(i, 2) & "', ���i��= '" & ws.Cells(i, 3) & "', �e��= '" _
         & ws.Cells(i, 4) & "', �l�i= '" & ws.Cells(i, 5) & "', ����= '" & ws.Cells(i, 6) & "', ���l= '" & ws.Cells(i, 7) & "' WHERE ID= " & ws.Cells(i, 1) & ""
        '���l�̏ꍇ�́u'�v�ł�����K�v���Ȃ��i������͕K�v�j
        
        .ActiveConnection = myConn
        .CommandText = strSQL
        .Execute
        
        End With
        
    i = i + 1
    
    Loop
    
    MsgBox "�f�[�^�x�[�X���X�V���܂���"
    
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
Sub �ی����()
    ActiveSheet.Unprotect
End Sub
Sub �f�[�^�폜()
    UserForm2.Show vbModeless
End Sub
