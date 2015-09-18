VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   Caption         =   "HOTEL RESERVATION MANAGEMENT SYSTEM "
   ClientHeight    =   8130
   ClientLeft      =   945
   ClientTop       =   1035
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "PROJECT IP.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   14280
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton REFRESH 
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   22
      Top             =   5640
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "class"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      ItemData        =   "PROJECT IP.frx":43DE
      Left            =   3960
      List            =   "PROJECT IP.frx":43EB
      TabIndex        =   21
      Text            =   "SELECT CLASS"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton CMDUPDATE 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton CMDEXIT 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   6600
      Width           =   3375
   End
   Begin VB.CommandButton CMDPREVIOUS 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      MaskColor       =   &H0000FFFF&
      TabIndex        =   18
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton CMDNEXT 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      MaskColor       =   &H0000FFFF&
      TabIndex        =   17
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton CMDLAST 
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      MaskColor       =   &H0000FFFF&
      TabIndex        =   16
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton CMDFIRST 
      BackColor       =   &H000000FF&
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      MaskColor       =   &H0000FFFF&
      TabIndex        =   15
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   14
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CALCULATE TOTAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   13
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      DataField       =   "total"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   4680
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      DataField       =   "rooms_on_rent"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      DataField       =   "days_of_staying"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   2880
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   7920
      Top             =   1920
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   794
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   50
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"PROJECT IP.frx":4405
      OLEDBString     =   $"PROJECT IP.frx":4499
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "SCOTT"
      Password        =   "TIGER"
      RecordSource    =   "HOTEL"
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
   Begin VB.TextBox Text3 
      DataField       =   "phone_no"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "address"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000014&
      Caption         =   "RECORDS OF SINGLES IN THE HOTEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2040
      TabIndex        =   23
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   "CLASS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000014&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000014&
      Caption         =   "ROOMS ON RENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000014&
      Caption         =   "DAYS OF STAYING"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000014&
      Caption         =   "PHONE NO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000014&
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000014&
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FIELD
Dim exp
Dim val


Private Sub CMDADDNEW_Click()
Form3.Hide
MDIForm1.Hide
Form5.Show
End Sub



Private Sub CMDDELETE_Click()

End Sub

Private Sub CMDEXIT_Click()
A = MsgBox("ARE YOU SURE DO YOU REALLY WANT TO EXIT", vbInformation + vbYesNo)
If A = vbYes Then
End
Else: Form1.Show
End If
End Sub

Private Sub CMDFIRST_Click()
Adodc1.Recordset.MoveFirst
Text1.SetFocus
End Sub

Private Sub CMDLAST_Click()
Adodc1.Recordset.MoveLast
Text1.SetFocus
End Sub

Private Sub CMDNEXT_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
MsgBox "YOU ARE ALREADY ON THE LAST RECORD"
Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub CMDPREVIOUS_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
MsgBox "YOU ARE ALREADY ON THE FIRST RECORD"
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub CMDTOTAL_Click()
Class = Text6.Text
days = val(Text4.Text)
rooms = val(Text5.Text)
If Class = "FIRST" Or Class = "first" Then
total = (rooms * days * 1800)
End If
If Class = "second" Or Class = "SECOND" Then
total = (rooms * days * 1000)
End If
If Class = "third" Or Class = "THIRD" Then
total = (rooms * days * 800)
End If
Text7.Text = total
End Sub

Private Sub CMDUPDATE_Click()
Adodc1.Recordset.Update

End Sub

Private Sub Combo1_Click()

F = Combo1.Text

End Sub





Private Sub Command1_Click()
If Option1.Value = True Then
Form1.Hide
MDIForm1.Hide
Form4.Show
Else
Form1.Hide
MDIForm1.Hide
Form6.Show
End If
End Sub

Private Sub Command2_Click()
A = MsgBox("ARE YOY SURE TOU WANT TO DELETE", vbYesNo)
If A = vbYes Then
Adodc1.Recordset.Delete
End If
End Sub

Private Sub Command3_Click()
val = "'" & Trim(Text9.Text) & "'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = ("select * from hotel where ") & FIELD & exp & val
Adodc1.REFRESH
End Sub

Private Sub Command4_Click()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from hotel"
Adodc1.REFRESH
End Sub

Private Sub Command5_Click()
If Combo1.Text = "first" Or Combo1.Text = "FIRST" Then
Text7.Text = (Text5.Text) * (Text4.Text) * 1800
ElseIf Combo1.Text = "second" Or Combo1.Text = "SECOND" Then
Text7.Text = (Text5.Text) * (Text4.Text) * 1000
ElseIf Combo1.Text = "third" Or Combo1.Text = "THIRD" Then
Text7.Text = (Text5.Text) * (Text4.Text) * 800
End If
End Sub

Private Sub Text9_Change()
val = Text9.Text
End Sub

Private Sub REFRESH_Click()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from HOTEL"
Adodc1.REFRESH

End Sub


