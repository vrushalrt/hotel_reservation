VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   7575
   ScaleWidth      =   11385
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      Left            =   1200
      TabIndex        =   18
      Top             =   5040
      Width           =   2895
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
      ItemData        =   "Form5.frx":3228
      Left            =   1560
      List            =   "Form5.frx":3235
      TabIndex        =   17
      Text            =   "SELECT CLASS"
      Top             =   2760
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   5280
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
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
      Connect         =   $"Form5.frx":324F
      OLEDBString     =   $"Form5.frx":32E3
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
      Height          =   615
      Left            =   2400
      TabIndex        =   16
      Top             =   4440
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GO TO MAIN FORM"
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
      Left            =   2280
      TabIndex        =   15
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
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
      Height          =   735
      Left            =   600
      TabIndex        =   14
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton CMDADDNEW 
      Caption         =   "ADDNEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      MaskColor       =   &H0000FFFF&
      TabIndex        =   13
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   2055
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
      Left            =   1560
      TabIndex        =   5
      Top             =   3480
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
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
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
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
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
      Left            =   1560
      TabIndex        =   1
      Top             =   480
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
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      TabIndex        =   12
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDADDNEW_Click()
Adodc1.Recordset.AddNew

Text1.SetFocus
End Sub

Private Sub CMDUPDATE_Click()
'Adodc2.Recordset.Update
Adodc1.Recordset.Update
End Sub

Private Sub Command1_Click()
If Combo1.Text = "first" Or Combo1.Text = "FIRST" Then
Text7.Text = (Text5.Text) * (Text4.Text) * 1800
ElseIf Combo1.Text = "second" Or Combo1.Text = "SECOND" Then
Text7.Text = (Text5.Text) * (Text4.Text) * 1000
ElseIf Combo1.Text = "third" Or Combo1.Text = "THIRD" Then
Text7.Text = (Text5.Text) * (Text4.Text) * 800
End If
End Sub

Private Sub Command2_Click()
A = MsgBox("ARE YOU SURE DO YOU REALLY WANT TO EXIT", vbInformation + vbYesNo)
If A = vbYes Then
End
Else: Form5.Show
End If
End Sub

Private Sub Command5_Click()
MDIForm1.Hide
Form1.Show

End Sub

Private Sub Form_Load()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
'If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("A")) Then
'KeyAscii = 0
'A = MsgBox("NUMBERS ONLY", vbCritical)
'Text1.Text = ""
'Text1.SetFocus
'End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 0 Or KeyAscii <= 10) Then
A = MsgBox("NUMBERS ONLY", vbCritical)
'Text34.Text = ""
Text4.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
'If Not (KeyAscii >= 0 Or KeyAscii <= 10) Then
'A = MsgBox("NUMBERS ONLY", vbCritical)
'Text5.Text = ""
'Text5.SetFocus


End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
'If (KeyAscii >= 0 Or KeyAscii <= 10) Then
'A = MsgBox("ALPHABETS ONLY", vbCritical)
'Text6.Text = ""
'Text6.SetFocus

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
'If Not (KeyAscii >= 0 Or KeyAscii <= 10) Then
'A = MsgBox("NUMBERS ONLY", vbCritical)
'Text7.Text = ""
'Text7.SetFocus

End Sub
