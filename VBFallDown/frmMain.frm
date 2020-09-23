VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VbFallDown"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrChecker 
      Interval        =   30
      Left            =   6600
      Top             =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Shape PlankR 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   5
      Left            =   2760
      Top             =   7560
      Width           =   5055
   End
   Begin VB.Shape PlankR 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   4
      Left            =   2760
      Top             =   6240
      Width           =   5055
   End
   Begin VB.Shape PlankR 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   2760
      Top             =   4920
      Width           =   5055
   End
   Begin VB.Shape PlankR 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   2760
      Top             =   3480
      Width           =   5055
   End
   Begin VB.Shape PlankR 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   2760
      Top             =   2160
      Width           =   5055
   End
   Begin VB.Shape PlankR 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   2760
      Top             =   840
      Width           =   5055
   End
   Begin VB.Shape PlankL 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   5
      Left            =   0
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Shape PlankL 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   4
      Left            =   0
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Shape PlankL 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   0
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Shape PlankL 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   0
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Shape PlankL 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Shape PlankL 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Shape Ball 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Quit, Pkey(255) As Boolean
Dim r1, r2, Score As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Pkey(KeyCode) = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Pkey(KeyCode) = False
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) * 0.5
Top = (Screen.Height - Height) * 0.5
Show
Do Until Quit = True
DoEvents: Scroll: HandleBall: RunKeys
Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
Quit = True: End
End Sub

Sub RunKeys()
If Pkey(37) = True And Ball.Left > 0 Then Ball.Left = Ball.Left - 2
If Pkey(39) = True And Ball.Left < 7560 Then Ball.Left = Ball.Left + 2
End Sub

Sub Scroll()
For i = 0 To 5
PlankL(i).Top = PlankL(i).Top - 1: PlankR(i).Top = PlankR(i).Top - 1
Next i
End Sub

Private Sub tmrChecker_Timer()
Score = Score + 1: Label1.Caption = "Score" & vbCrLf & Score
For i = 0 To 5
If PlankL(i).Top < -200 Then ResetPlank Val(i)
Next i
If Ball.Top < 0 Then
Open App.Path + "\High.txt" For Input As #1
Input #1, TheHigh#
If Score > TheHigh# Then MsgBox "You got a high score!", vbInformation + vbOKOnly, "Cool!" Else MsgBox "The high score is " & TheHigh# & ".  Your score was " & Score & ".", vbOKOnly + vbInformation, "Oops!": Unload Me
Close #1
Open App.Path + "\High.txt" For Output As #1
Print #1, Score
Close #1
Unload Me
End If
End Sub

Sub ResetPlank(WP As Integer)
PlankL(WP).Top = 8000
PlankR(WP).Top = 8000
r1 = Int(Rnd * 2)
Select Case r1
Case 0: r2 = Int(Rnd * 6900): PlankL(WP).Width = r2: PlankR(WP).Left = r2 + 855: PlankR(WP).Width = Width
Case 1: r2 = Int(Rnd * 6000): PlankL(WP).Width = r2: PlankR(WP).Left = r2 + 1700: PlankR(WP).Width = Width
Case 2: r2 = Int(Rnd * 4900): PlankL(WP).Width = r2: PlankR(WP).Left = r2 + 2666: PlankR(WP).Width = Width
End Select
End Sub

Sub HandleBall()
If Ball.Top > 7680 Then Ball.Top = Ball.Top - 1
Ball.Top = Ball.Top + 1
For i = 0 To 5
If PlankL(i).Top - Ball.Top = 240 Or PlankL(i).Top - Ball.Top = 241 Then
If Ball.Left < PlankL(i).Width Or Ball.Left > PlankR(i).Left - 240 Then Ball.Top = Ball.Top - 2
End If
Next i
End Sub
