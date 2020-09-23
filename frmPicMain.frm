VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Media Viewer"
   ClientHeight    =   8595
   ClientLeft      =   -1140
   ClientTop       =   -450
   ClientWidth     =   11880
   Icon            =   "frmPicMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMoveRight 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   2400
   End
   Begin VB.Timer tmrMoveLeft 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   2400
   End
   Begin VB.CommandButton cmdUseless 
      Height          =   255
      Left            =   11640
      TabIndex        =   8
      Top             =   8280
      Width           =   255
   End
   Begin VB.VScrollBar Yaxis 
      Height          =   8175
      LargeChange     =   1000
      Left            =   11640
      SmallChange     =   100
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar Xaxis 
      Height          =   255
      LargeChange     =   1000
      Left            =   3120
      SmallChange     =   100
      TabIndex        =   7
      Top             =   8280
      Width           =   8535
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00000000&
      Height          =   8175
      Left            =   3120
      ScaleHeight     =   541
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   565
      TabIndex        =   10
      Top             =   120
      Width           =   8535
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   0
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   11
         Top             =   0
         Width           =   1335
      End
      Begin MediaPlayerCtl.MediaPlayer MP1 
         Height          =   8055
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   8535
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   -1  'True
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   -1  'True
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
         Volume          =   -600
         WindowlessVideo =   0   'False
      End
      Begin VB.Image PictureTemp 
         Height          =   975
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
      Height          =   8415
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   8040
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   4770
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.CheckBox chkMP 
      Caption         =   "Media Player (On/Off)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkMP_Click()
    If chkMP.Value = 1 Then
        Picture1.Visible = False
        MP1.Visible = True
        Xaxis.Enabled = False
        Yaxis.Enabled = False
    Else
        Picture1.Visible = True
        MP1.Visible = False
        Xaxis.Enabled = True
        Yaxis.Enabled = True
    End If
End Sub

Private Sub cmdMove_Click()
    If picMain.Left = 208 Then
        Let tmrMoveLeft.Enabled = True
    ElseIf picMain.Left = 28 Then
        Let tmrMoveRight.Enabled = True
    End If
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Dir1_Change()
    Let File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo Error
    Let Dir1.Path = Drive1.Drive
    Exit Sub
Error:
    MsgBox "Device is Unavailable", vbCritical, "Error"
End Sub

Private Sub File1_Click()
    On Error Resume Next
    Dim file As String
    file = Dir1.Path & "\" & File1.FileName
    If chkMP.Value = 1 Then
        Let MP1.FileName = file
        Let Form1.Caption = "Media Viewer - Viewing: [" & file & "]"
    Else
        Let MP1.FileName = "" 'Clear Screen
        Let PictureTemp.Picture = LoadPicture(file)
        Let Picture1.Picture = PictureTemp.Picture
        Let Picture1.Width = PictureTemp.Width
        Let Picture1.Height = PictureTemp.Height
        Let Form1.Caption = "Media Viewer - Viewing: [" & file & "]"
        Let Xaxis.Max = Picture1.ScaleWidth - picMain.ScaleWidth
        Let Yaxis.Max = Picture1.ScaleHeight - picMain.ScaleHeight
    End If
End Sub

Private Sub Form_Load()
    Let Drive1.Drive = "C:\"
    Let Dir1.Path = "C:\"
    Let File1.Path = Dir1.Path
    Let chkMP.Value = 0
    Let MP1.Visible = False
    Let MP1.Width = picMain.Width
    Let MP1.Height = picMain.Height
End Sub

Private Sub tmrMoveLeft_Timer()
    picMain.Left = picMain.Left - 10
    picMain.Width = picMain.Width + 10
    Xaxis.Left = picMain.Left
    Xaxis.Width = picMain.Width
    If picMain.Left = 28 Then
        tmrMoveLeft.Enabled = False
        Let cmdMove.Left = 8
        Let cmdMove.Caption = ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
        Let Drive1.Visible = False
        Let Dir1.Visible = False
        Let File1.Visible = False
        Let cmdQuit.Visible = False
        Let chkMP.Visible = False
    ElseIf picMain.Left < 28 Then
        picMain.Left = 28
        Xaxis.Left = 28
    End If
    Let Xaxis.Max = Picture1.ScaleWidth - picMain.ScaleWidth
    Let Yaxis.Max = Picture1.ScaleHeight - picMain.ScaleHeight
    Let MP1.Width = picMain.Width
    Let MP1.Height = picMain.Height
End Sub

Private Sub tmrMoveRight_Timer()
    Let cmdMove.Visible = False
    Let Drive1.Visible = True
    Let Dir1.Visible = True
    Let File1.Visible = True
    Let cmdQuit.Visible = True
    Let chkMP.Visible = True
    picMain.Left = picMain.Left + 10
    picMain.Width = picMain.Width - 10
    Xaxis.Left = picMain.Left
    Xaxis.Width = picMain.Width
    If picMain.Left = 208 Then
        Let tmrMoveRight = False
        Let cmdMove.Visible = True
        Let cmdMove.Left = 184
        Let cmdMove.Caption = "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
    End If
    Let Xaxis.Max = Picture1.ScaleWidth - picMain.ScaleWidth
    Let Yaxis.Max = Picture1.ScaleHeight - picMain.ScaleHeight
    Let MP1.Width = picMain.Width
    Let MP1.Height = picMain.Height
End Sub

Private Sub Xaxis_Change()
    Let Picture1.Left = -Xaxis.Value
End Sub


Private Sub Yaxis_Change()
    Let Picture1.Top = -Yaxis.Value
End Sub
