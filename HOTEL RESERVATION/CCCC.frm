VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
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
      Left            =   3480
      TabIndex        =   22
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Caption         =   "search record "
      Height          =   1935
      Left            =   8280
      TabIndex        =   15
      Top             =   240
      Width           =   3255
      Begin VB.CommandButton Command3 
         Caption         =   "click"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   615
         Left            =   1680
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "CCCC.frx":0000
         Left            =   1800
         List            =   "CCCC.frx":0013
         TabIndex        =   18
         Text            =   "operator"
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "CCCC.frx":0028
         Left            =   240
         List            =   "CCCC.frx":0041
         TabIndex        =   17
         Text            =   "search"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "refresh"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "report"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Width           =   975
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
      Left            =   4680
      TabIndex        =   6
      Top             =   960
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
      Left            =   4680
      TabIndex        =   5
      Top             =   1560
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
      Left            =   4680
      TabIndex        =   4
      Top             =   2160
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
      Left            =   4680
      TabIndex        =   3
      Top             =   2760
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
      Left            =   4680
      TabIndex        =   2
      Top             =   3360
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
      Left            =   4680
      TabIndex        =   1
      Top             =   4560
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
      ItemData        =   "CCCC.frx":008C
      Left            =   4680
      List            =   "CCCC.frx":0099
      TabIndex        =   0
      Text            =   "SELECT CLASS"
      Top             =   3960
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   7800
      Top             =   2760
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
      Connect         =   $"CCCC.frx":00B3
      OLEDBString     =   $"CCCC.frx":0147
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
      Caption         =   "SEARCH RECORDS FOR COUPLES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   21
      Top             =   240
      Width           =   4815
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
      Left            =   3120
      TabIndex        =   13
      Top             =   960
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
      Left            =   3120
      TabIndex        =   12
      Top             =   1560
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
      Left            =   3120
      TabIndex        =   11
      Top             =   2160
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
      Left            =   3120
      TabIndex        =   10
      Top             =   2640
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
      Left            =   3120
      TabIndex        =   9
      Top             =   3360
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
      Left            =   3120
      TabIndex        =   8
      Top             =   4560
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
      Left            =   3120
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FIELD
Dim exp
Dim val
Private Sub CMDEXIT_Click()
A = MsgBox("ARE YOU SURE DO YOU REALLY WANT TO EXIT", vbInformation + vbYesNo)
If A = vbYes Then
End
Else: Form8.Show
End If
End Sub

Private Sub Combo2_Change()
Text9.Text = ""
End Sub

Private Sub Command3_Click()
val = "'" & Trim(Text9.Text) & "'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = ("select * from COUPLE where ") & FIELD & exp & val
Adodc1.REFRESH
End Sub

Private Sub Command4_Click()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from COUPLE"
Adodc1.REFRESH
End Sub

Private Sub Combo3_Click()
exp = Combo3.Text
End Sub

Private Sub Combo2_Click()
FIELD = Combo2.Text
Text9.Text = ""
End Sub

