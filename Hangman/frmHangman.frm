VERSION 5.00
Begin VB.Form frmHangman 
   Caption         =   "Hangman"
   ClientHeight    =   6600
   ClientLeft      =   2685
   ClientTop       =   1095
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   6270
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
      Left            =   1920
      TabIndex        =   27
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton cmdZ 
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   26
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdX 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   24
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdY 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   25
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdW 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   23
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdV 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   22
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdR 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   18
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   19
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdT 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   20
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdU 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   21
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdQ 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   17
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdN 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   14
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdM 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   13
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdL 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   12
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdK 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   11
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdJ 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   10
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdP 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   16
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdI 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   9
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdO 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   15
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdH 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdG 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdF 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdE 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   4
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdB 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblword 
      Alignment       =   2  'Center
      Caption         =   "_ _ _ _ _ _ _ _ _ _"
      BeginProperty Font 
         Name            =   "Myriad Condensed"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
End
Attribute VB_Name = "frmHangman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strWord As String
Dim strCaption As String
Dim strCap2 As String
Dim intNumWrong As Integer
Dim strWord2 As String

Private Sub cmdA_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "A") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "A" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdA.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdB_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "B") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "B" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdB.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdC_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "C") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "C" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdC.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdD_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "D") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "D" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdD.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdE_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "E") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "E" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdE.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdF_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "F") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "F" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdF.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdG_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "G") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "G" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdG.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdH_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "H") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "H" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdH.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdI_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "I") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "I" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdI.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdJ_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "J") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "J" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdJ.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdK_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "K") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "K" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdK.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdL_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "L") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "L" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdL.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdM_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "M") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "M" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdM.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdN_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "N") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "N" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdN.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdO_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "O") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "O" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdO.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdP_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "P") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "P" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdP.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdQ_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "Q") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "Q" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdQ.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdQuit_Click()
 End
End Sub

Private Sub cmdR_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "R") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "R" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdR.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdS_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "S") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "S" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdS.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdT_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "T") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "T" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdT.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdU_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "U") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "U" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdU.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdV_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "V") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "V" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdV.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdW_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "W") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "W" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdW.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdX_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "X") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "X" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdX.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdY_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "Y") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "Y" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdY.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub cmdZ_Click()
 Dim blnFound As Boolean
 blnFound = False
 Dim intC2 As Integer
 intC2 = 1
 While intC2 <= Len(strWord)
  If (Mid(strWord, intC2, 1) = "Z") Then
   strCaption = Mid(strCaption, 1, (intC2 * 2) - 2) + "Z" _
     + Mid(strCaption, intC2 * 2, Len(strCaption) - (intC2 * 2) + 1)
   blnFound = True
  End If
  intC2 = intC2 + 1
 Wend
 If (blnFound <> True) Then
  intNumWrong = intNumWrong + 1
  Call WrongOne(intNumWrong)
  frmHangman.Hide
  frmMan.Show (vbModeless)
 End If
 lblword.Caption = strCaption
 cmdZ.Enabled = False
 If (strCaption = strWord2) Then
  frmHangman.Hide
  frmWin.Show (vbModeless)
 End If
End Sub

Private Sub Form_Activate()
 strWord = GoingOut()
 If (strCap2 = strCaption) Then
  strCaption = ""
  Dim intCount As Integer
  While intCount < Len(strWord)
   If (Mid(strWord, intCount + 1, 1) <> " ") Then
    strCaption = strCaption + "_ "
   Else
    strCaption = strCaption + "  "
   End If
   intCount = intCount + 1
  Wend
  lblword.Caption = strCaption
  strCap2 = strCaption
 End If
 strWord = UCase(strWord)
 Dim intC2 As Integer
 intC2 = 0
 While intC2 < Len(strWord)
  If (Mid(strWord, intC2 + 1, 1) <> " ") Then
   strWord2 = strWord2 + Mid(strWord, intC2 + 1, 1) + " "
  Else
   strWord2 = strWord2 + "  "
  End If
  intC2 = intC2 + 1
 Wend
End Sub

Private Sub Form_Load()
 frmHangman.Hide
 frmEnter.Show (vbModeless)
End Sub
