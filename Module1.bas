Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public dsn As New ADODB.Recordset
Public mk As ADODB.Recordset
Public n As ADODB.Recordset

Public Sub db()
Set con = New ADODB.Connection
Set dsn = New ADODB.Recordset
Set mk = New ADODB.Recordset
Set n = New ADODB.Recordset
con.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbakademik.mdb"
End Sub
Public Sub a()
On Error Resume Next
End Sub
