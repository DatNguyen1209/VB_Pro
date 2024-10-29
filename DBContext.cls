VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private conn As ADODB.Connection

Public Sub OpenConnection()
    Set conn = New ADODB.Connection

    Dim connectionString As String
    connectionString = "Provider=SQLOLEDB;Data Source=9801PCS0304\MSSQLSERVER01;Initial Catalog=DB_LEARN;Integrated Security=SSPI;"

    conn.Open connectionString

'    If conn.State = adStateOpen Then
'        MsgBox "Connected to DB successfully!"
'    Else
'        MsgBox "Connected to DB failded"
'    End If
End Sub

Public Property Get Connection() As ADODB.Connection
    Set Connection = conn
End Property

Public Sub CloseConnection()
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Sub
Public Function ExecuteQuery(query As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    rs.Open query, conn, adOpenStatic, adLockReadOnly
    Set ExecuteQuery = rs
End Function
