Attribute VB_Name = "connection"

Public conn As ADODB.connection
Public Sub connection()
Set conn = New ADODB.connection
conn.CursorLocation = adUseClient
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\P4\librarymgmtsystem.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
conn.Open

End Sub
