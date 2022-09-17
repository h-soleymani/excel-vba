Function SQL_Select(SQL As String) As Variant
Dim Conn As New ADODB.connection
Dim Rst As ADODB.Recordset
FilePath = ThisWorkbook.FullName
connstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FilePath & ";Extended Properties=""Excel 12.0 Macro;HDR=YES"";"
Conn.Open connstr
Dim result As Variant
Set Rst = New ADODB.Recordset
Rst.Open SQL, Conn
result = Rst.GetRows
Rst.Close
Set Rst = Nothing
On Error GoTo 0
SQL_Select = result
End Function
