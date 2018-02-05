VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10890
   LinkTopic       =   "Form3"
   ScaleHeight     =   5640
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox combo_city 
      Height          =   315
      Left            =   2880
      TabIndex        =   9
      Text            =   "choose one city"
      Top             =   840
      Width           =   4695
   End
   Begin VB.CommandButton cmd_delete 
      Caption         =   "delete"
      Height          =   735
      Left            =   6480
      TabIndex        =   5
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmd_insert 
      Caption         =   "insert"
      Height          =   855
      Left            =   6480
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton cmd_view 
      Caption         =   "view"
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3615
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6376
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
   Begin VB.TextBox txt_phone 
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txt_name 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "phone no"
      Height          =   495
      Left            =   8400
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "city"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_insert_Click()
con.Execute "insert into com values('" & Me.txt_name.Text & "','" & Me.combo_city.Text & "'," & Me.txt_phone.Text & " )"
con.Execute "commit"
Call cmd_view_Click
MsgBox "inserted"
End Sub

Private Sub cmd_view_Click()
If rs.State = 1 Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.Open "select * from com ", con
Set Me.DataGrid1.DataSource = rs
End Sub

Private Sub Form_Load()
Me.combo_city.AddItem "kolkata"
Me.combo_city.AddItem "mumbai"
Me.combo_city.AddItem "delhi"
Me.combo_city.AddItem "chennai"
Me.combo_city.AddItem "pune"

End Sub
