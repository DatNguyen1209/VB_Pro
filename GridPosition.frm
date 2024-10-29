VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnDelete 
      Caption         =   "Xoa"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton btnEdit 
      BackColor       =   &H000040C0&
      Caption         =   "Sua"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton btnAdd 
      BackColor       =   &H0000FFFF&
      Caption         =   "Them moi"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DBContext
Private SelectData As PositionDTO
Public Sub HoldingData(ByRef position As PositionDTO)
    Set SelectData = position
End Sub
Private Sub btnAdd_Click()
    Form2.Show
End Sub
Public Sub DeleteRecord(data As PositionDTO)
    Set db = New DBContext
    
    db.OpenConnection
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.CommandText = "DELETE FROM POSITION WHERE ID = ?"
    Set cmd.ActiveConnection = db.Connection
    cmd.Parameters.Append cmd.CreateParameter(, adInteger, adParamInput, , data.Id)
    cmd.Execute
    Dim result As Integer
    result = MsgBox("Delete successfully!", vbOKOnly + vbInformation)

    If result = vbOK Then
        Form1.Form_Load
        Set SelectData = Nothing
    End If
End Sub
Private Sub btnDelete_Click()
    If Not SelectData Is Nothing Then
        Dim result As Integer

        result = MsgBox("You want to delete record?", vbYesNo + vbQuestion, "Ok")
        
        If result = vbYes Then
            DeleteRecord SelectData
            
        End If
    Else
        MsgBox "No position selected."
    End If
End Sub

Private Sub btnEdit_Click()
    If Not SelectData Is Nothing Then
        Form2.SetPositionData SelectData
        Form2.Show
    Else
        MsgBox "No position selected."
    End If
End Sub

Public Sub Form_Load()
    Set db = New DBContext
    
    db.OpenConnection

    Dim rs As ADODB.Recordset
    Dim query As String
    query = "SELECT * FROM POSITION"
    Set result = db.ExecuteQuery(query)
    Set DataGrid1.DataSource = result
End Sub
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim selectedRow As Integer
    Dim obj As PositionDTO
 
    Set obj = New PositionDTO
    selectedRow = DataGrid1.Row
    
    obj.Id = CInt(DataGrid1.Columns(0).Text)
    obj.Code = DataGrid1.Columns(1).Text
    obj.Name = DataGrid1.Columns(2).Text
    obj.Size = CInt(DataGrid1.Columns(3).Text)

    HoldingData obj
End Sub

