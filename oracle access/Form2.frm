VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6360
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   10890
   LinkTopic       =   "Form2"
   ScaleHeight     =   6360
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   480
      TabIndex        =   16
      Top             =   3960
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4048
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
      Height          =   3015
      Left            =   12000
      Top             =   2880
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   5318
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
      RecordSource    =   "select * from std"
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
   Begin VB.CommandButton cmd_exit 
      Caption         =   "form3"
      Height          =   855
      Left            =   9360
      TabIndex        =   12
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmd_delete 
      Caption         =   "delete"
      Height          =   735
      Left            =   7680
      TabIndex        =   11
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton cmd_update 
      Caption         =   "update"
      Height          =   735
      Left            =   5160
      TabIndex        =   10
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton cmd_insert 
      Caption         =   "insert"
      Height          =   735
      Left            =   2640
      TabIndex        =   9
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmd_view 
      Caption         =   "view"
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmd_last 
      Caption         =   "last"
      Height          =   975
      Left            =   7440
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmd_next 
      Caption         =   "next"
      Height          =   975
      Left            =   5160
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmd_prev 
      Caption         =   "previous"
      Height          =   975
      Left            =   2640
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmd_first 
      Caption         =   "first"
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin MSACAL.Calendar cal_dob 
      Height          =   2295
      Left            =   9240
      TabIndex        =   3
      Top             =   240
      Width           =   3855
      _Version        =   524288
      _ExtentX        =   6800
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2018
      Month           =   1
      Day             =   14
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txt_roll 
      Height          =   615
      Left            =   6120
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txt_name 
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txt_stdid 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "roll (number)"
      Height          =   495
      Left            =   6240
      TabIndex        =   15
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "name ( char)"
      Height          =   495
      Left            =   2880
      TabIndex        =   14
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "id (text)"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cal_dob_Click()
txt_name.Text = CDate(cal_dob.Value)
End Sub

Private Sub cmd_delete_Click()
con.Execute "delete from std where stdid='" & txt_stdid.Text & "' "
Call cmd_view_Click
MsgBox "deleted"
End Sub

Private Sub cmd_exit_Click()
form3.Show
Unload Me
End Sub

Private Sub cmd_first_Click()
Me.cmd_next.Enabled = True
Me.cmd_prev.Enabled = True

Call cmd_view_Click

End Sub

Private Sub cmd_insert_Click()
con.Execute "insert into std  values('" & Me.txt_stdid.Text & "','" & Me.txt_name.Text & "'," & Me.txt_roll & ",to_date('" & cal_dob.Value & "','mm/dd/yyyy' )) "

Call cmd_view_Click
MsgBox "insertion done"
End Sub

Private Sub cmd_last_Click()
rs.MoveLast
Me.cmd_next.Enabled = False
Set Me.DataGrid1.DataSource = rs
Me.txt_name.Text = rs.Fields("name").Value
Me.txt_roll.Text = rs.Fields("roll").Value
Me.txt_stdid.Text = rs.Fields("stdid").Value
Me.cal_dob.Value = rs.Fields("dob").Value
End Sub

Private Sub cmd_next_Click()

rs.MoveNext
Me.cmd_prev.Enabled = True
Me.cmd_last.Enabled = True
Me.cmd_first.Enabled = True

If rs.EOF = True Then
Me.cmd_next.Enabled = False
Me.cmd_last.Enabled = False
MsgBox "this is last record"

Else



Me.txt_name.Text = rs.Fields("name").Value
Me.txt_roll.Text = rs.Fields("roll").Value
Me.txt_stdid.Text = rs.Fields("stdid").Value
Me.cal_dob.Value = rs.Fields("dob").Value

End If


End Sub

Private Sub cmd_prev_Click()
rs.MovePrevious
Me.cmd_last.Enabled = True
Me.cmd_next.Enabled = True

If rs.BOF = True Then
Me.cmd_prev.Enabled = False
MsgBox " this is first record"

Else
Me.txt_name.Text = rs.Fields("name").Value
Me.txt_roll.Text = rs.Fields("roll").Value
Me.txt_stdid.Text = rs.Fields("stdid").Value
Me.cal_dob.Value = rs.Fields("dob").Value
End If

End Sub

Private Sub cmd_update_Click()
con.Execute "update std set name='" & txt_name.Text & "',roll=" & txt_roll.Text & ",dob=to_date('" & cal_dob.Value & "','mm/dd/yyyy') where stdid='" & txt_stdid.Text & "'"

MsgBox "done"
End Sub

Private Sub cmd_view_Click()
If rs.State = 1 Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.Open "select * from std", con
Set Me.DataGrid1.DataSource = rs
rs.MoveFirst

Me.cmd_prev.Enabled = False

Me.txt_name.Text = rs.Fields("name").Value
Me.txt_roll.Text = rs.Fields("roll").Value
Me.txt_stdid.Text = rs.Fields("stdid").Value
Me.cal_dob.Value = rs.Fields("dob").Value

End Sub

