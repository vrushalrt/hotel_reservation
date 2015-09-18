VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   11280
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
      Left            =   3360
      TabIndex        =   23
      Top             =   5760
      Width           =   1815
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
      Left            =   6240
      TabIndex        =   22
      Top             =   5280
      Width           =   2055
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
      Left            =   480
      MaskColor       =   &H0000FFFF&
      TabIndex        =   21
      Top             =   5280
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
      Left            =   1920
      MaskColor       =   &H0000FFFF&
      TabIndex        =   20
      Top             =   5280
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
      Left            =   3360
      MaskColor       =   &H0000FFFF&
      TabIndex        =   19
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   1455
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
      Left            =   4800
      MaskColor       =   &H0000FFFF&
      TabIndex        =   18
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   1455
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
      Left            =   2640
      TabIndex        =   17
      Top             =   6480
      Width           =   3375
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
      Left            =   480
      TabIndex        =   16
      Top             =   5760
      Width           =   1455
   End
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
      Left            =   1920
      TabIndex        =   15
      Top             =   5760
      Width           =   1455
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
      Left            =   4080
      TabIndex        =   6
      Top             =   840
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
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Width           =   3015
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
      Left            =   4080
      TabIndex        =   4
      Top             =   2040
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
      Left            =   4080
      TabIndex        =   3
      Top             =   2640
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
      Left            =   4080
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
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
      Left            =   4080
      TabIndex        =   1
      Top             =   4440
      Width           =   3015
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
      ItemData        =   "COUPLE.frx":0000
      Left            =   4080
      List            =   "COUPLE.frx":000D
      TabIndex        =   0
      Text            =   "SELECT CLASS"
      Top             =   3840
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   7680
      Top             =   2640
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
      Connect         =   $"COUPLE.frx":0027
      OLEDBString     =   $"COUPLE.frx":00BB
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "SCOTT"
      Password        =   "TIGER"
      RecordSource    =   "COUPLE"
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
   Begin VB.Label Label8 
      Caption         =   "RECORDS OF COUPLES IN THE HOTEL"
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
      Left            =   2160
      TabIndex        =   14
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label2 
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
      Left            =   2520
      TabIndex        =   13
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Left            =   2520
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label4 
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
      Left            =   2520
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label5 
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
      Left            =   2520
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label6 
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
      Left            =   2520
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label7 
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
      Left            =   2520
      TabIndex        =   8
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Left            =   2520
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDEXIT_Click()
A = MsgBox("ARE YOU SURE DO YOU REALLY WANT TO EXIT", vbInformation + vbYesNo)
If A = vbYes Then
End
Else: Form6.Show
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

Private Sub CMDUPDATE_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command2_Click()
A = MsgBox("ARE YOY SURE TOU WANT TO DELETE", vbYesNo)
If A = vbYes Then
Adodc1.Recordset.Delete
End If
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

Private Sub REFRESH_Click()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from hotel"
Adodc1.REFRESH
End Sub

