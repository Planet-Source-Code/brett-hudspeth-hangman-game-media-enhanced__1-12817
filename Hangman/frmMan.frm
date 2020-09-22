VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmMan 
   BackColor       =   &H00FFFF00&
   Caption         =   "Hangman"
   ClientHeight    =   6120
   ClientLeft      =   2685
   ClientTop       =   1470
   ClientWidth     =   6090
   DrawWidth       =   3
   ForeColor       =   &H00000040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   6090
   Begin VB.Timer tmrLose 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5520
      Top             =   1440
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
   Begin MediaPlayerCtl.MediaPlayer mdpSong 
      Height          =   975
      Left            =   4320
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblLose2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "The Word Was:"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblLose1 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "You Lose!"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Line linMouth 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   3120
      X2              =   3480
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Shape shpEye2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shpEye1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Line linLeg1 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   3480
      X2              =   3480
      Y1              =   2880
      Y2              =   3600
   End
   Begin VB.Line linLeg2 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   3120
      X2              =   3120
      Y1              =   2880
      Y2              =   3600
   End
   Begin VB.Line linArm2 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   3000
      X2              =   2400
      Y1              =   1800
      Y2              =   1440
   End
   Begin VB.Line linArm1 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   3600
      X2              =   4200
      Y1              =   1800
      Y2              =   1440
   End
   Begin VB.Shape shpBody 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpHead 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpRope 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   3240
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shpPole2 
      BorderColor     =   &H00004080&
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   960
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Shape shpPole 
      BorderColor     =   &H00004080&
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   3615
      Left            =   960
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape shpGallows 
      BorderColor     =   &H00404080&
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   960
      Top             =   3840
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Shape shpGround 
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   -120
      Top             =   4800
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "frmMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intY2 As Integer
Dim intX1 As Integer
Dim intX2 As Integer
Dim intCount As Integer

Private Sub cmdBack_Click()
 frmMan.Hide
 frmHangman.Show (vbModeless)
End Sub

Private Sub cmdQuit_Click()
 End
End Sub

Private Sub Form_Activate()
 intY2 = linArm1.Y2
 intX1 = linLeg1.X2
 intX2 = linLeg2.X2
 Dim intWrong As Integer
 intWrong = Wrong()
 Select Case (intWrong)
 Case 1
  shpGround.Visible = True
 Case 2
  shpGround.Visible = True
  shpGallows.Visible = True
 Case 3
  shpGround.Visible = True
  shpGallows.Visible = True
  shpPole.Visible = True
 Case 4
  shpGround.Visible = True
  shpGallows.Visible = True
  shpPole.Visible = True
  shpPole2.Visible = True
 Case 5
  shpGround.Visible = True
  shpGallows.Visible = True
  shpPole.Visible = True
  shpPole2.Visible = True
  shpRope.Visible = True
 Case 6
  shpGround.Visible = True
  shpGallows.Visible = True
  shpPole.Visible = True
  shpPole2.Visible = True
  shpRope.Visible = True
  shpHead.Visible = True
 Case 7
  shpGround.Visible = True
  shpGallows.Visible = True
  shpPole.Visible = True
  shpPole2.Visible = True
  shpRope.Visible = True
  shpHead.Visible = True
  shpBody.Visible = True
 Case 8
  shpGround.Visible = True
  shpGallows.Visible = True
  shpPole.Visible = True
  shpPole2.Visible = True
  shpRope.Visible = True
  shpHead.Visible = True
  shpBody.Visible = True
  linLeg1.Visible = True
  linLeg2.Visible = True
 Case 9
  shpGround.Visible = True
  shpGallows.Visible = True
  shpPole.Visible = True
  shpPole2.Visible = True
  shpRope.Visible = True
  shpHead.Visible = True
  shpBody.Visible = True
  linLeg1.Visible = True
  linLeg2.Visible = True
  linArm1.Visible = True
  linArm2.Visible = True
  lblLose1.Visible = True
  lblLose2.Caption = lblLose2.Caption + vbCrLf + GoingOut()
  lblLose2.Visible = True
  cmdBack.Enabled = False
  mdpSong.FileName = "C:\Windows\Media\In the Hall of the Mountain King.rmi"
  mdpSong.Volume = -150
  mdpSong.Play
  tmrLose.Enabled = True
 End Select
 intCount = 1
End Sub

Private Sub tmrLose_Timer()
 linArm1.Y2 = intY2
 linArm2.Y2 = intY2
 linLeg1.X2 = intX1
 linLeg2.X2 = intX2
 If (intCount Mod 2 = 1) Then
  linArm1.Y2 = intY2 + 100
  linArm2.Y2 = intY2 + 100
  linLeg1.X2 = intX1 + 100
  linLeg2.X2 = intX2 - 100
 Else
  linArm1.Y2 = intY2 - 100
  linArm2.Y2 = intY2 - 100
  linLeg1.X2 = intX1 - 100
  linLeg2.X2 = intX2 + 100
 End If
 intCount = intCount + 1
 
End Sub
