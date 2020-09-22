VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample Mp3 Player by Dbushea"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      Columns         =   1
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1740
      ItemData        =   "Form1.frx":0000
      Left            =   0
      List            =   "Form1.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   2655
      Left            =   4320
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   3015
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
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -110
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MediaPlayer1.FileName = "C:\web\mp3\third eye blind\Third Eye Blind-Deep Insde You.mp3"

MediaPlayer1.Play

End Sub

Private Sub Command2_Click()
CommonDialog1.Filter = "Mp3 Files (*.mp3) | *.mp3"
CommonDialog1.ShowOpen
If Not CommonDialog1.FileName = "" Then
List1.List(List1.ListCount) = Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)
List2.List(List2.ListCount) = CommonDialog1.FileName
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
List1.RemoveItem 0
List2.RemoveItem 0
End Sub

Private Sub List1_DblClick()
MediaPlayer1.Stop
MediaPlayer1.FileName = List2.List(List1.ListIndex)
lstindex = List1.ListIndex
MediaPlayer1.Play
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
newsong
End Sub

Private Sub Timer1_Timer()
Sec = MediaPlayer1.CurrentPosition
minutes = Fix(Sec / 60)
seconds = Fix(Sec - minutes * 60)
If Len(minutes) = 1 Then
minutes = "0" & minutes
End If
If Val(seconds) < 0 Then
seconds = 0
End If

If Len(seconds) = 1 Then
seconds = "0" & seconds
End If

Label1.Caption = minutes & ":" & seconds

End Sub
