VERSION 5.00
Begin VB.Form image1 
   Caption         =   "Form1"
   ClientHeight    =   12015
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   19950
   LinkTopic       =   "Form1"
   ScaleHeight     =   12015
   ScaleWidth      =   19950
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer12 
      Interval        =   1000
      Left            =   2880
      Top             =   2400
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   1095
      Left            =   16680
      TabIndex        =   18
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   1335
      Left            =   15120
      TabIndex        =   17
      Top             =   6960
      Width           =   4455
   End
   Begin VB.Timer Timer11 
      Interval        =   1000
      Left            =   2520
      Top             =   2400
   End
   Begin VB.Timer Timer10 
      Interval        =   1
      Left            =   2160
      Top             =   2400
   End
   Begin VB.Timer Timer9 
      Interval        =   1
      Left            =   1800
      Top             =   2400
   End
   Begin VB.Timer Timer8 
      Interval        =   1
      Left            =   1440
      Top             =   2400
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Press to Hit!"
      Height          =   855
      Left            =   12960
      TabIndex        =   16
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdPress 
      Caption         =   "Press to Hit!"
      Height          =   855
      Left            =   6000
      TabIndex        =   15
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   1080
      Top             =   2400
   End
   Begin VB.Timer Timer6 
      Interval        =   750
      Left            =   720
      Top             =   2400
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   360
      Top             =   2400
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   0
      Top             =   2400
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   18840
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   18360
      Top             =   1320
   End
   Begin VB.TextBox txtInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   8880
      TabIndex        =   12
      Top             =   1800
      Width           =   10455
   End
   Begin VB.TextBox txtCommentary 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   17880
      Top             =   1320
   End
   Begin VB.Frame Frame5 
      Caption         =   "Timer"
      Height          =   1215
      Left            =   18000
      TabIndex        =   5
      Top             =   0
      Width           =   1215
      Begin VB.TextBox txtTimer 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Hard Mode Highscore"
      Height          =   1695
      Left            =   15360
      TabIndex        =   4
      Top             =   0
      Width           =   2535
      Begin VB.TextBox txtHighscore3 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Normal Mode Highscore"
      Height          =   1695
      Left            =   12600
      TabIndex        =   3
      Top             =   0
      Width           =   2655
      Begin VB.TextBox txtHighscore2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Easy Mode Highscore"
      Height          =   1695
      Left            =   9840
      TabIndex        =   2
      Top             =   0
      Width           =   2655
      Begin VB.TextBox txtHighscore1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   1215
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Score"
      Height          =   1695
      Left            =   7200
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin VB.TextBox txtScore 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   1215
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Image Image5 
      Height          =   7185
      Left            =   4440
      Picture         =   "HitMeForm.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   10170
   End
   Begin VB.Image Image4 
      Height          =   3255
      Left            =   2880
      Picture         =   "HitMeForm.frx":2E14F
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Image Image3 
      Height          =   1560
      Left            =   0
      Picture         =   "HitMeForm.frx":53E05
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   1560
   End
   Begin VB.Image Image2 
      Height          =   1575
      Left            =   0
      Picture         =   "HitMeForm.frx":7761F
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   0
      Picture         =   "HitMeForm.frx":7E9EE
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1560
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HIT ME!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   67.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Menu cmdPresentTime 
      Caption         =   "Date and Time"
   End
   Begin VB.Menu txtHowToPlay 
      Caption         =   "How to Play"
      Begin VB.Menu cmdInstructions 
         Caption         =   "How To Play Game 01"
      End
      Begin VB.Menu cmdInstructions2 
         Caption         =   "How to Play Game 02"
      End
   End
   Begin VB.Menu txtStartGame 
      Caption         =   "Start Game"
      Begin VB.Menu cmdStart 
         Caption         =   "Start Game 01"
      End
      Begin VB.Menu cmdStart2 
         Caption         =   "Start Game 02"
      End
   End
   Begin VB.Menu txtResetGame 
      Caption         =   "Reset Game "
      Begin VB.Menu cmdReset 
         Caption         =   "Reset Game 01 "
      End
      Begin VB.Menu cmdReset2 
         Caption         =   "Reset Game 02"
      End
      Begin VB.Menu cmdResetFull 
         Caption         =   "Reset Full Game"
      End
   End
   Begin VB.Menu txtGame1 
      Caption         =   "Game 01"
      Begin VB.Menu txtLvl 
         Caption         =   "Levels"
         Begin VB.Menu cmdLvl1 
            Caption         =   "Easy"
         End
         Begin VB.Menu cmdLvl2 
            Caption         =   "Normal"
         End
         Begin VB.Menu cmdLvl3 
            Caption         =   "Hard"
         End
      End
   End
   Begin VB.Menu txtGame2 
      Caption         =   "Game 02"
      Begin VB.Menu txtLvl2 
         Caption         =   "Levels"
         Begin VB.Menu cmdLvlA 
            Caption         =   "Easy"
         End
         Begin VB.Menu cmdLvlB 
            Caption         =   "Normal"
         End
         Begin VB.Menu cmdLvlC 
            Caption         =   "Hard"
         End
      End
   End
   Begin VB.Menu cmdExit 
      Caption         =   "Exit Game"
   End
End
Attribute VB_Name = "image1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim timeLegendary As Integer
Dim timeFlies As Integer
Dim time1 As Integer
Dim time2 As Integer
Dim time3 As Integer
Dim time4 As Integer
Dim time5 As Integer
Dim time6 As Integer
Dim pick As Integer
Dim xPosA As Integer
Dim xPosB As Integer
Dim xPosC As Integer
Dim yPosA As Integer
Dim yPosB As Integer
Dim yPosC As Integer
Dim xPosD As Integer
Dim yPosD As Integer
Dim xPosE As Integer
Dim yPosE As Integer
Dim xPosF As Integer
Dim yPosF As Integer
Dim xPosG As Integer
Dim yPosG As Integer
Dim xPosH As Integer
Dim score As Integer
Dim highscore3 As Integer
Dim highscore2 As Integer
Dim highscore1 As Integer

Private Sub cmdPress_Click()
Beep
score = score + 1
txtScore = score
End Sub
Private Sub cmdClick_Click()
Beep
score = score + 1
txtScore = score
End Sub

Private Sub Form_Load()
cmdPresentTime.Visible = False
txtHowToPlay.Visible = False
txtStartGame.Visible = False
txtResetGame.Visible = False
txtGame1.Visible = False
txtGame2.Visible = False
Title.Visible = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
txtScore.Visible = False
txtHighscore1.Visible = False
txtHighscore2.Visible = False
txtHighscore3.Visible = False
txtTimer.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
txtCommentary.Visible = False
txtInstructions.Visible = False
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
cmdPress.Visible = False
cmdClick.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = True
timeLegendary = 20
Text3.Enabled = False
Text4.Enabled = False
End Sub
Private Sub Timer11_Timer()
timeLegendary = timeLegendary - 1
Text4 = timeLegendary
    If timeLegendary = 18 Then
    Image5.Visible = True
End If
    If timeLegendary = 15 Then
    Text3.Visible = True
End If
    If timeLegendary = 14 Then
    Text3.Text = CStr("There Is No Game.")
End If
    If timeLegendary = 12 Then
    Text3.Text = CStr("Just Kidding!")
End If
    If timeLegendary = 10 Then
    Text4.Visible = True
    Text3.Text = CStr("Get Ready")
End If
    If timeLegendary = 5 Then
    Text3.Visible = False
End If
    If timeLegendary <= 0 Then
    Timer11.Enabled = False
    Call LetThereBeLight
End If

End Sub


Private Sub LetThereBeLight()
cmdPresentTime.Visible = True
txtHowToPlay.Visible = True
txtStartGame.Visible = True
txtResetGame.Visible = True
txtGame1.Visible = True
txtGame2.Visible = True
Title.Visible = True
Title.Enabled = False
Frame1.Visible = True
Frame1.Enabled = False
Frame2.Visible = True
Frame2.Enabled = False
Frame3.Visible = True
Frame3.Enabled = False
Frame4.Visible = True
Frame4.Enabled = False
Frame5.Visible = True
Frame5.Enabled = False
txtScore.Visible = True
txtScore.Enabled = False
txtHighscore1.Visible = True
txtHighscore1.Enabled = False
txtHighscore2.Visible = True
txtHighscore2.Enabled = False
txtHighscore3.Visible = True
txtHighscore3.Enabled = False
txtTimer.Visible = True
txtTimer.Enabled = False
Text1.Visible = True
Text1.Enabled = False
Text2.Visible = True
Text2.Enabled = False
Text3.Visible = False
Text4.Visible = False
txtCommentary.Visible = True
txtCommentary.Enabled = False
txtInstructions.Visible = True
txtInstructions.Enabled = False
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
cmdPress.Visible = False
cmdClick.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Call Opening_Disable
End Sub

Private Sub Opening_Disable()
Randomize
txtTimer.Text = CStr("?s")
txtScore.Text = CStr("0")
txtHighscore1.Text = CStr("0")
txtHighscore2.Text = CStr("0")
txtHighscore3.Text = CStr("0")
txtCommentary.Text = CStr("Welcome to HIT ME! Press how to play for game instructions.")
xPosC = 9240
Image3.Left = xPosC
time3 = 20
time2 = 15
time1 = 10
time4 = 10
Text3.Visible = False
Text4.Visible = False
cmdReset.Enabled = False
cmdReset2.Enabled = False
cmdResetFull.Enabled = False
cmdStart2.Enabled = False
cmdStart.Enabled = False
cmdPress.Visible = False
cmdClick.Visible = False
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
End Sub
Private Sub cmdExit_Click()
  Dim x As Integer
  x = MsgBox("Are you sure you want to exit?", vbYesNo, "Exit")
  If x = 6 Then
    End
  End If
End Sub
Private Sub Timer4_Timer()
Time = Time + 1
Text1 = Time

End Sub
Private Sub cmdPresentTime_Click()
Dim thedate As Date
thedate = Date
Text2 = Format(thedate, "dddd, mmm d yyyy")
Text1 = Time
Timer4.Enabled = True
cmdPresentTime.Visible = False

End Sub
Private Sub Game1_Disable()
txtGame2.Enabled = False
cmdLvl1.Enabled = False
cmdLvl2.Enabled = False
cmdLvl3.Enabled = False
txtCommentary.Text = CStr("You chose easy! Press Start To Start!")
txtInstructions.Text = CStr("")
txtTimer.Text = CStr("10s")
cmdStart.Enabled = True
End Sub

Private Sub cmdInstructions_Click()
txtHowToPlay.Enabled = False
txtInstructions.Text = CStr("Please hover over Levels. From there select your difficulty then proceed to press start game to start. After the game's over, press Reset. Have Fun!")
End Sub
Private Sub cmdInstructions2_Click()
txtInstructions.Text = CStr("Please hover over Levels. From there select your difficulty then proceed to press start game to start. After the game's over, press Reset. Have Fun!")
txtHowToPlay.Enabled = False
End Sub
Private Sub cmdLvlA_Click()
Call Game2_Disable
pick = 1
Image4.Visible = True
cmdPress.Visible = True
cmdPress.Enabled = False

End Sub
Private Sub cmdLvlB_Click()
Call Game2_Disable
pick = 2
Image2.Visible = True
xPosD = 5640
yPosD = 6480
Image2.Left = xPosD
Image2.Top = yPosD
cmdClick.Visible = False

End Sub
Private Sub cmdLvlC_Click()
Call Game2_Disable
pick = 3
Image3.Visible = True
xPosE = 0
yPosE = 9480
Image3.Left = xPosE
Image3.Top = yPosE
End Sub

Private Sub Game2_Disable()
txtGame1.Enabled = False
cmdLvlA.Enabled = False
cmdLvlB.Enabled = False
cmdLvlC.Enabled = False
txtCommentary.Text = CStr("You chose easy! Press Start To Start!")
txtTimer.Text = CStr("10s")
cmdStart2.Enabled = True
txtHowToPlay.Enabled = True

End Sub

Private Sub cmdLvl1_Click()
Call Game1_Disable
pick = 1
Image1.Visible = True
Image1.Enabled = False

End Sub

Private Sub cmdLvl2_Click()
Call Game1_Disable
pick = 2
Image2.Visible = True
Image2.Enabled = False

End Sub

Private Sub cmdLvl3_Click()
Call Game1_Disable
pick = 3
Image3.Visible = True
Image3.Enabled = False

End Sub
Private Sub cmdReset_Click()
Call New_Disable

End Sub
Private Sub cmdReset2_Click()
Call New_Disable2
End Sub
Private Sub cmdResetFull_Click()
Randomize
txtTimer.Text = CStr("?s")
txtScore.Text = CStr("0")
txtHighscore1.Text = CStr("0")
txtHighscore2.Text = CStr("0")
txtHighscore3.Text = CStr("0")
txtCommentary.Text = CStr("Welcome to HIT ME! Press how to play for game instructions.")
xPosC = 9240
Image3.Left = xPosC
time3 = 20
time2 = 15
time1 = 10
time4 = 10
time5 = 0
Call Freedom

End Sub
Private Sub Freedom()
cmdReset.Enabled = False
cmdReset2.Enabled = False
cmdResetFull.Enabled = False
cmdStart2.Enabled = False
cmdStart.Enabled = False
cmdPress.Visible = False
cmdClick.Visible = False
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = True
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
txtGame2.Enabled = True
txtGame1.Enabled = True
cmdLvlA.Enabled = True
cmdLvlB.Enabled = True
cmdLvlC.Enabled = True
cmdLvl1.Enabled = True
cmdLvl2.Enabled = True
cmdLvl3.Enabled = True
Text3.Visible = False
Text4.Visible = False
Title.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False
Frame3.Enabled = False
Frame4.Enabled = False
Frame5.Enabled = False
txtScore.Enabled = False
txtHighscore1.Enabled = False
txtHighscore2.Enabled = False
txtHighscore3.Enabled = False
txtTimer.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
txtCommentary.Enabled = False
txtInstructions.Enabled = False

End Sub
Private Sub New_Disable2()
cmdReset2.Enabled = False
Randomize
score = 0
txtTimer.Text = CStr("?s")
txtScore.Text = CStr("0")
txtCommentary.Text = CStr("Choose Another Level or choose the same one!")
xPosC = 9240
Image3.Left = xPosC
time4 = 10
time5 = 10
cmdLvlA.Enabled = True
cmdLvlB.Enabled = True
cmdLvlC.Enabled = True
txtHowToPlay.Enabled = False

End Sub
Private Sub New_Disable()
cmdReset.Enabled = False
Randomize
score = 0
txtTimer.Text = CStr("?s")
txtScore.Text = CStr("0")
txtCommentary.Text = CStr("Choose Another Level or choose the same one!")
xPosC = 9240
Image3.Left = xPosC
time3 = 20
time2 = 15
time1 = 10
cmdLvl1.Enabled = True
cmdLvl2.Enabled = True
cmdLvl3.Enabled = True
txtHowToPlay.Enabled = False

End Sub

Private Sub cmdStart_Click()
cmdReset.Enabled = False
cmdResetFull.Enabled = True
cmdStart.Enabled = False
Image1.Enabled = True
Image2.Enabled = True
Image3.Enabled = True
txtCommentary.Text = CStr("HIT THEM HARD!")
txtHowToPlay.Enabled = False


If pick = 1 Then
Timer1.Enabled = True
End If

If pick = 2 Then
Timer2.Enabled = True
End If

If pick = 3 Then
Timer3.Enabled = True
End If

End Sub
Private Sub cmdStart2_Click()
cmdReset2.Enabled = False
cmdResetFull.Enabled = True
cmdStart2.Enabled = False
Image1.Enabled = True
Image2.Enabled = True
Image3.Enabled = True
txtCommentary.Text = CStr("Have Fun!")
txtHowToPlay.Enabled = False


If pick = 1 Then
Timer5.Enabled = True
End If

If pick = 2 Then
Timer6.Enabled = True
End If

If pick = 3 Then
Timer7.Enabled = True
txtInstructions.Text = CStr("Catch him by clicking on him!!!")
End If

End Sub

Private Sub Image1_Click()
Beep
score = score + 1
txtScore = score
xPosA = Int(Rnd * 18000)
yPosA = Int(Rnd * 500) + 2300
Image1.Left = xPosA
Image1.Top = yPosA
End Sub

Private Sub Image2_Click()
Beep
score = score + 1
txtScore = score

End Sub

Private Sub Image3_Click()
Beep
score = score + 1
txtScore = score

End Sub


Private Sub Timer1_Timer()

time1 = time1 - 1
txtTimer = time1 & "s"
If time1 <= 0 Then
    MsgBox ("Times Up! Your score was: " & score)
    Timer1.Enabled = False
    Image1.Enabled = False
    Image1.Visible = False
    txtCommentary.Text = CStr("Good Job Press Reset To Start Again!")
    cmdReset.Enabled = True
        If score > highscore1 Then
        highscore1 = score
        txtHighscore1 = highscore1
    End If
End If

End Sub

Private Sub Timer2_Timer()
xPosB = Int(Rnd * 18000)
yPosB = Int(Rnd * 5000) + Int(Rnd * 4000)
Image2.Left = xPosB
Image2.Top = yPosB

time2 = time2 - 1
txtTimer = time2 & "s"
If time2 <= 0 Then
    MsgBox ("Times Up! Your score was: " & score)
    Timer2.Enabled = False
    Image2.Enabled = False
    Image2.Visible = False
    txtCommentary.Text = CStr("Good Job Press Reset To Start Again!")
    cmdReset.Enabled = True
        If score > highscore2 Then
        highscore2 = score
        txtHighscore2 = highscore2
    End If
End If
End Sub

Private Sub Timer3_Timer()
xPosC = Int(Rnd * 18000)
yPosC = Int(Rnd * 5000) + Int(Rnd * 4000)
Image3.Left = xPosC
Image3.Top = yPosC

time3 = time3 - 1
txtTimer = time3 & "s"
If time3 <= 0 Then
   MsgBox ("Times Up! Your score was: " & score)
   Timer3.Enabled = False
   Image3.Enabled = False
   Image3.Visible = False
   txtCommentary.Text = CStr("Good Job Press Reset To Start Again!")
   cmdReset.Enabled = True
   If score > highscore3 Then
        highscore3 = score
        txtHighscore3 = highscore3
    End If
End If

End Sub


Private Sub Timer5_Timer()
cmdPress.Enabled = True
time4 = time4 - 1
txtTimer = time4 & "s"
If time4 <= 0 Then
    MsgBox ("Times Up! Your score was: " & score)
    Timer5.Enabled = False
    Image4.Enabled = False
    Image4.Visible = False
    txtCommentary.Text = CStr("Good Job Press Reset To Start Again!")
    txtInstructions.Text = CStr("")
    cmdReset2.Enabled = True
    cmdPress.Visible = False
        If score > highscore1 Then
        highscore1 = score
        txtHighscore1 = highscore1
    End If
End If
End Sub

Private Sub Timer6_Timer()
cmdClick.Enabled = True
cmdClick.Visible = True
xPosB = Int(Rnd * 2000) + 2000 + Int(Rnd * 5000)
yPosB = Int(Rnd * 9000) + 9000
   cmdClick.Left = yPosB
   cmdClick.Top = xPosB

xPosD = Int(Rnd * 200) + 5640
yPosD = Int(Rnd * 200) + 6480
Image2.Left = xPosD
Image2.Top = yPosD
txtInstructions.Text = CStr("Catch me if you can!!!")

time5 = time5 - 1
txtTimer = time5 & "s"
If time5 <= 0 Then
    MsgBox ("Times Up! Your score was: " & score)
    Timer6.Enabled = False
    Image2.Enabled = False
    Image2.Visible = False
    txtCommentary.Text = CStr("Good Job Press Reset To Start Again!")
    txtInstructions.Text = CStr("")
    cmdReset2.Enabled = True
    cmdClick.Visible = False
        If score > highscore2 Then
        highscore2 = score
        txtHighscore2 = highscore2
    End If
End If
End Sub

Private Sub Timer7_Timer()
xPosE = xPosE + Int(Rnd * 500) + 1
    Image3.Left = xPosE

If xPosE >= 17760 Then
       Timer7.Enabled = False
       Timer8.Enabled = True
       xPosE = 0
       Image3.Left = xPosE
       yPosE = 7560
       Image3.Top = yPosE
End If
End Sub

Private Sub Timer8_Timer()
xPosF = xPosF + Int(Rnd * 500) + 1
    Image3.Left = xPosF

If xPosF >= 17760 Then
    Timer8.Enabled = False
    Timer9.Enabled = True
    xPosF = 0
    Image3.Left = xPosF
    yPosF = 5640
    Image3.Top = yPosF
End If
End Sub

Private Sub Timer9_Timer()
xPosG = xPosG + Int(Rnd * 500) + 1
    Image3.Left = xPosG

If xPosG >= 17760 Then
    Timer9.Enabled = False
    Timer10.Enabled = True
    xPosG = 0
    Image3.Left = xPosG
    yPosG = 3720
    Image3.Top = yPosG
End If
End Sub

Private Sub Timer10_Timer()
txtTimer.Text = CStr("?s")
xPosH = xPosH + Int(Rnd * 500) + 1
    Image3.Left = xPosH

If xPosH >= 17760 Then
    Timer10.Enabled = False
    MsgBox ("Times Up! Your score was: " & score)
    Image3.Enabled = False
    Image3.Visible = False
    txtCommentary.Text = CStr("Good Job Press Reset To Start Again!")
    txtInstructions.Text = CStr("")
    cmdReset2.Enabled = True
        If score > highscore3 Then
        highscore3 = score
        txtHighscore3 = highscore3
    End If
End If
End Sub
