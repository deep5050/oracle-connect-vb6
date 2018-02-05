VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   9000
      TabIndex        =   3
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1440
      Top             =   3720
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;Password=12345;User ID=system;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=12345;User ID=system;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "system"
      Password        =   "12345"
      RecordSource    =   "select * from login"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_login 
      Caption         =   "LOG IN"
      Height          =   855
      Left            =   5040
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txt_pass 
      Height          =   975
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txt_uid 
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "password (number)"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "user id (character)"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_login_Click()
If rs.State = 1 Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.Open "select * from login where userid='" & txt_uid.Text & "' and password=" & txt_pass.Text & " ", con
Set Me.DataGrid1.DataSource = rs
If rs.RecordCount > 0 Then
MsgBox "welcome"
Unload Me
Form3.Show
Else
MsgBox "error"
End If



End Sub



Private Sub Form_Load()
con.Open "Provider=MSDAORA.1;Password=12345;User ID=system;Persist Security Info=True"
End Sub

Private Sub txt_pass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmd_login_Click
End If


End Sub

