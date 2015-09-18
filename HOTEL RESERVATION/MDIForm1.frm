VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7200
   ClientLeft      =   1845
   ClientTop       =   1785
   ClientWidth     =   9870
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Menu VIEW 
      Caption         =   "&VIEW"
      Enabled         =   0   'False
      Begin VB.Menu CV 
         Caption         =   "COUPLES RECORDS"
      End
      Begin VB.Menu SINGLES 
         Caption         =   "SINGLES RECORDS"
         Index           =   0
      End
   End
   Begin VB.Menu ADD 
      Caption         =   "&ADD"
      Enabled         =   0   'False
      Begin VB.Menu ADDTO 
         Caption         =   "ADD TO SINGLES RECORDS"
      End
      Begin VB.Menu ADDTO1 
         Caption         =   "ADD TO COUPLES RECORDS"
      End
   End
   Begin VB.Menu SEARCH 
      Caption         =   "&SEARCH RECORDS BY"
      Enabled         =   0   'False
      Begin VB.Menu CX 
         Caption         =   "COUPLES"
         Index           =   2
      End
      Begin VB.Menu SING 
         Caption         =   "SINGLES"
      End
   End
   Begin VB.Menu REP 
      Caption         =   "&REPORTS OF"
      Enabled         =   0   'False
      Begin VB.Menu re 
         Caption         =   "SINGLES RECORDS"
      End
      Begin VB.Menu CUP 
         Caption         =   "COUPLES RECORDS"
      End
   End
   Begin VB.Menu OPT 
      Caption         =   "&OPTIONS"
      Enabled         =   0   'False
      Begin VB.Menu END 
         Caption         =   "END"
      End
      Begin VB.Menu ARR 
         Caption         =   "ARRANGE"
         Begin VB.Menu HOR 
            Caption         =   "TILE HORIZONTAL"
         End
         Begin VB.Menu VER 
            Caption         =   "VERTICAL"
         End
         Begin VB.Menu CAS 
            Caption         =   "CASCADE"
         End
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ADDTO_Click()
MDIForm1.Hide
ADDTO.Checked = True
ADDTO1.Checked = False
CV.Checked = False
SINGLES(0).Checked = False
ADD.Checked = False
SEARCH.Checked = False
CX(2).Checked = False
SING.Checked = False
REP.Checked = False
re.Checked = False
CUP.Checked = False
Form5.Show
End Sub

Private Sub ADDTO1_Click()
MDIForm1.Hide
ADDTO.Checked = False
ADDTO1.Checked = True
ADDTO.Checked = True

CV.Checked = False
SINGLES(0).Checked = False
ADD.Checked = False
SEARCH.Checked = False
CX(2).Checked = False
SING.Checked = False
REP.Checked = False
re.Checked = False
CUP.Checked = False
Form10.Show
End Sub

Private Sub CAS_Click()
MDIForm1.Arrange vbCascade

End Sub

Private Sub CUP_Click()
MDIForm1.Hide
Form11.Show
CUP.Checked = True
CV.Checked = True
SINGLES(0).Checked = False
ADDTO.Checked = False
ADDTO1.Checked = False


ADD.Checked = False
SEARCH.Checked = False
CX(2).Checked = False
SING.Checked = False
REP.Checked = False
re.Checked = False

End Sub

Private Sub CV_Click()
MDIForm1.Hide
CV.Checked = True
SINGLES(0).Checked = False
ADDTO.Checked = False
ADDTO1.Checked = False


ADD.Checked = False
SEARCH.Checked = False
CX(2).Checked = False
SING.Checked = False
REP.Checked = False
re.Checked = False
CUP.Checked = False
Form6.Show
End Sub

Private Sub CX_Click(Index As Integer)
MDIForm1.Hide
CX(2).Checked = True
ADDTO.Checked = True
ADDTO1.Checked = False
CV.Checked = False
SINGLES(0).Checked = False
ADD.Checked = False
SEARCH.Checked = False

SING.Checked = False
REP.Checked = False
re.Checked = False
CUP.Checked = False
SING.Checked = False
Form8.Show
End Sub

Private Sub report_Click()
MDIForm1.Hide


Form2.Show
End Sub

Private Sub reservation_Click()
MDIForm1.Hide

Form1.Show

End Sub


Private Sub SINGLE_Click(Index As Integer)
MDIForm1.Hide

Form4.Show
End Sub

Private Sub END_Click()
A = MsgBox("ARE YOU SURE DO YOU REALLY WANT TO EXIT", vbInformation + vbYesNo)
If A = vbYes Then
End

End If
End Sub

Private Sub HOR_Click()
MDIForm1.Arrange vbTileHorizontal
End Sub

Private Sub re_Click()
MDIForm1.Hide
Form2.Show
re.Checked = True
ADDTO.Checked = True
ADDTO1.Checked = False
CV.Checked = False
SINGLES(0).Checked = False
ADD.Checked = False
SEARCH.Checked = False
CX(2).Checked = False
SING.Checked = False
REP.Checked = False

CUP.Checked = False
CUP.Checked = False
End Sub

Private Sub SING_Click()
MDIForm1.Hide
SING.Checked = True
CX(2).Checked = False
ADDTO.Checked = True
ADDTO1.Checked = False
CV.Checked = False
SINGLES(0).Checked = False
ADD.Checked = False
SEARCH.Checked = False
CX(2).Checked = False

REP.Checked = False
re.Checked = False
CUP.Checked = False
Form4.Show
End Sub

Private Sub SINGLES_Click(Index As Integer)
MDIForm1.Hide
CV.Checked = False
SINGLES(0).Checked = True
ADDTO.Checked = True
ADDTO1.Checked = False
CV.Checked = False

ADD.Checked = False
SEARCH.Checked = False
CX(2).Checked = False
SING.Checked = False
REP.Checked = False
re.Checked = False
CUP.Checked = False
Form1.Show
End Sub

Private Sub VER_Click()
MDIForm1.Arrange vbCascade

End Sub
