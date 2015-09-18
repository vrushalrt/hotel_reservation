VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "HOTEL RESERVATION MANAGEMENT SYSTEM "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   2
      Tag             =   "ISHAAN"
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "THIS SOFTWARE IS COPYRIGHT UNDER SECTION 453-C .ANY ILLEGAL COPY OF IT WILL BE A PUNISHABLE ACT."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   5520
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "ENTER PASSWORD"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "WELCOME TO HOTEL RESERVATION              MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = Text1.Tag Then
frmSplash.Show
A = MsgBox("RECORDS OF ALLTHE CUSTOMERS OF OUR HOTEL JUST SECONDS AWAY,ENJOY THE TOUR", vbExclamation)
Else
B = MsgBox("NO NO  NO! INVALID PASSWORD. YOU CANNOT CONTINUE", vbCritical, "PASSWORD VALIDATION")
End If
  C = MsgBox("TRY ONCE MORE OR EXIT", vbRetryCancel)
  If C = vbRetry Then
  Text1.SetFocus
  Text1.Text = ""
  ElseIf C = vbCancel Then
  
  U = MsgBox("BYE,BYE", vbOKOnly)
  End
  End If
  
  
  
  

End Sub

