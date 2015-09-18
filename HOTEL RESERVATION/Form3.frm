VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000080&
   Caption         =   "Form3"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13395
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   8070
   ScaleWidth      =   13395
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "SHOW DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Width           =   4935
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Form3.frx":0D42
         Left            =   1080
         List            =   "Form3.frx":0D4C
         TabIndex        =   3
         Text            =   "SELECT"
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   1
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "                  HOTEL                         MANAGEMENT  SYSTEM"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1215
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   6135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()

End Sub

Private Sub Combo1_Change()
If Combo1.Text = "SINGLES" Then
MDIForm1.Hide
Form1.Show
Else
End If
If Combo1.Text = "COUPLES" Then
MDIForm1.Hide
Form6.Show
End If
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "SINGLES" Then
MDIForm1.Hide
Form1.Show

End If
If Combo1.Text = "COUPLES" Then
MDIForm1.Hide
Form6.Show
End If
End Sub

Private Sub Command1_Click()
MDIForm1.Hide
Form1.Show
Form3.Hide
End Sub

Private Sub Command2_Click()
A = MsgBox("ARE YOU SURE DO YOU REALLY WANT TO EXIT", vbInformation + vbYesNo)
If A = vbYes Then
End
Else: Form3.Show
End If
End Sub

