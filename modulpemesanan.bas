Attribute VB_Name = "modulpemesanan"
Public KON As New ADODB.Connection
Public rspesan As New ADODB.Recordset
Public rspemesan As New ADODB.Recordset
Public rsadmin As New ADODB.Recordset
Public rskereta As New ADODB.Recordset
Public rstiket As New ADODB.Recordset
Public rsmember As New ADODB.Recordset
Public rstransaksi As New ADODB.Recordset
Sub koneksi()
Set KON = New ADODB.Connection
Set rspesan = New ADODB.Recordset
Set rspemesan = New ADODB.Recordset
Set rsadmin = New ADODB.Recordset
Set rskereta = New ADODB.Recordset
Set rstiket = New ADODB.Recordset
Set rsmember = New ADODB.Recordset
Set rstransaksi = New ADODB.Recordset
KON.ConnectionString = "driver=mysql odbc 3.51 driver;server=localhost;uid=root;db=db_tiket;"
KON.Open
End Sub
