VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6390
   LinkTopic       =   "Form2"
   ScaleHeight     =   1920
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSize 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Height          =   405
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Size:"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name:"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private f2dataSelect As PositionDTO
Private Sub btnCancel_Click(Index As Integer)
    Form2.Hide
    Form2.Refresh
End Sub



Public Sub SaveData(data As PositionDTO, db As DBContext)
     Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
   
    Set cmd.ActiveConnection = db.Connection
  
    cmd.CommandText = "INSERT INTO POSITION(ID, CODE, NAME, SIZE) VALUES(?, ?, ?, ?)"
    
    cmd.Parameters.Append cmd.CreateParameter(, adInteger, adParamInput, , data.Id)
    cmd.Parameters.Append cmd.CreateParameter(, adVarChar, adParamInput, 50, data.Code)
    cmd.Parameters.Append cmd.CreateParameter(, adVarChar, adParamInput, 50, data.Name)
    cmd.Parameters.Append cmd.CreateParameter(, adInteger, adParamInput, , data.Size)
    cmd.Execute

    MsgBox "Insert data successfully!"
 
    Set cmd = Nothing
End Sub

Private Sub btnSave_Click(Index As Integer)
    Set obj = New PositionDTO

    obj.Size = CInt(txtSize.Text)
    obj.Name = txtName.Text
    Set db = New DBContext
    db.OpenConnection

    Dim rs As ADODB.Recordset
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command

    Set cmd.ActiveConnection = db.Connection
    Dim insertQuery As String
    insertQuery = "INSERT INTO POSITION(CODE, NAME, SIZE, ID) VALUES(?, ?, ?, ?)"
    Dim updateQuery As String
    updateQuery = "UPDATE POSITION SET CODE = ?, NAME = ?, SIZE = ? WHERE ID = ?"

    cmd.CommandText = IIf(f2dataSelect Is Nothing, insertQuery, updateQuery)

    If Not f2dataSelect Is Nothing Then
        Set obj = f2dataSelect
        cmd.Parameters.Append cmd.CreateParameter(, adInteger, adParamInput, , obj.Id)
    Else
        Dim query As String
        query = "SELECT TOP 1 * FROM POSITION P ORDER BY P.ID DESC"
        Set rs = db.ExecuteQuery(query)

        If Not rs.EOF Then
            obj.Id = rs.Fields("ID").Value + 1
        Else
            obj.Id = 1
        End If
        obj.Code = CStr(obj.Id)
    End If

    cmd.Parameters.Append cmd.CreateParameter(, adVarChar, adParamInput, 50, obj.Code)
    cmd.Parameters.Append cmd.CreateParameter(, adVarChar, adParamInput, 50, obj.Name)
    cmd.Parameters.Append cmd.CreateParameter(, adInteger, adParamInput, , obj.Size)
    cmd.Parameters.Append cmd.CreateParameter(, adInteger, adParamInput, , obj.Id)

    cmd.Execute

    Dim result As Integer
    result = MsgBox("Action successfully!", vbOKOnly + vbInformation)

    If result = vbOK Then
        Form2.Hide
        Form1.Form_Load
        f2dataSelect = Nothing
    End If

End Sub

Private data As PositionDTO

Public Sub SetPositionData(ByRef position As PositionDTO)
    Set data = position
    txtName.Text = data.Name
    txtSize.Text = data.Size
    Set f2dataSelect = data
End Sub

