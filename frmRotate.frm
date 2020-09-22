VERSION 5.00
Begin VB.Form Rotate 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotate by BattleStorm"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRotate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   75
      Top             =   4500
   End
   Begin VB.CommandButton cmdRealtimeRotation 
      Caption         =   "Turn Realtime Rotation On"
      Height          =   300
      Left            =   75
      TabIndex        =   8
      Tag             =   "0"
      Top             =   3150
      Width           =   4815
   End
   Begin VB.CommandButton cmdRotateBits 
      Caption         =   "Bit Rotate - Very Fast"
      Height          =   300
      Left            =   75
      TabIndex        =   7
      ToolTipText     =   "Click to rotate using bits"
      Top             =   2775
      Width           =   1740
   End
   Begin VB.CommandButton cmdRotatePoints 
      Caption         =   "Point Rotate - Slow"
      Height          =   300
      Left            =   75
      TabIndex        =   6
      ToolTipText     =   "Click to Rotate using points"
      Top             =   2175
      Width           =   1740
   End
   Begin VB.CheckBox chkAntiAlias 
      Caption         =   "AA"
      Height          =   240
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "Anti-Alias"
      Top             =   1875
      Value           =   1  'Checked
      Width           =   540
   End
   Begin VB.CommandButton cmdRotatePixels 
      Caption         =   "Pixel Rotate - Medium"
      Height          =   300
      Left            =   75
      TabIndex        =   3
      ToolTipText     =   "Click to rotate using pixels"
      Top             =   2475
      Width           =   1740
   End
   Begin VB.PictureBox picDestination 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   1875
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   2
      Top             =   75
      Width           =   3000
   End
   Begin VB.TextBox txtAngle 
      Height          =   240
      Left            =   525
      TabIndex        =   1
      Text            =   "45.000"
      ToolTipText     =   "Angle in Degrees"
      Top             =   1875
      Width           =   615
   End
   Begin VB.PictureBox picSource 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   75
      Picture         =   "frmRotate.frx":0000
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   115
      TabIndex        =   0
      Top             =   75
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Angle"
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   1875
      Width           =   465
   End
End
Attribute VB_Name = "Rotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* CODED BY: BattleStorm
'* EMAIL: battlestorm@cox.net
'* UPDATED: 08/02/2002
'* PURPOSE: Rotates a picture from one
'*     picturebox to another using an
'*     angle specified in degrees from
'*     -359.999° to 359.999°.
'* COPYRIGHT: This program and source
'*     code is freeware and can be copied
'*     and/or distributed as long as you
'*     mention the original author. I am
'*     not responsible for any harm as the
'*     outcome of using any of this code.

'* Most of the code is original and coded by me.
'* Some of the rotate functionality and variable
'* names may be inspired by other coders submissions,
'* but no code has been cut and pasted.
'* The only object borrowed is the "Hell Inside"
'* picture from one of XasanSoft's submissions.

'* If you use any of my code (modified or unmodified),
'* please mention me somewhere in your program. I
'* worked really hard to get the Anti-Alias and
'* centering of the destination picture to work.

Option Explicit

'Variables
Private idx As Integer
Private CodeTimer As clsTimer

'Rotate picture using points
Private Sub cmdRotatePoints_Click()
  'Initialize timer
  Set CodeTimer = New clsTimer
  
  'Set mouse pointer to hourglass
  Me.MousePointer = 11
  
  'Clear destination and check for events
  picDestination.Cls
  DoEvents

  'Start timer. Rotate picture. Stop timer.
  CodeTimer.StartTimer
  PointRotate picSource, picDestination, Val(txtAngle.Text), chkAntiAlias.Value
  CodeTimer.StopTimer
  
  'Set mouse pointer back to default
  Me.MousePointer = 0
  
  'Display elapsed processing time in form's caption
  Me.Caption = "Processing Time: " & CodeTimer.Elasped & " ms"
End Sub

'Rotate picture using pixels
Private Sub cmdRotatePixels_Click()
  'Initialize timer
  Set CodeTimer = New clsTimer
  
  'Set mouse pointer to hourglass
  Me.MousePointer = 11
  
  'Clear destination and check for events
  picDestination.Cls
  DoEvents

  'Start timer. Rotate picture. Stop timer.
  CodeTimer.StartTimer
  PixelRotate picSource, picDestination, Val(txtAngle.Text), chkAntiAlias.Value
  CodeTimer.StopTimer
  
  'Set mouse pointer back to default
  Me.MousePointer = 0
  
  'Display elapsed processing time in form's caption
  Me.Caption = "Processing Time: " & CodeTimer.Elasped & " ms"
End Sub

'Rotate picture using bits
Private Sub cmdRotateBits_Click()
  'Initialize timer
  Set CodeTimer = New clsTimer
  
  'Set mouse pointer to hourglass
  Me.MousePointer = 11
  
  'Clear destination and check for events
  picDestination.Cls
  DoEvents

  'Start timer. Rotate picture. Stop timer.
  CodeTimer.StartTimer
  BitRotate picSource, picDestination, Val(txtAngle.Text), chkAntiAlias.Value
  CodeTimer.StopTimer
  
  'Set mouse pointer back to default
  Me.MousePointer = 0
  
  'Display elapsed processing time in form's caption
  Me.Caption = "Processing Time: " & CodeTimer.Elasped & " ms"
End Sub

'Rotate picture in realtime using bits
Private Sub cmdRealtimeRotation_Click()
  If cmdRealtimeRotation.Tag = 0 Then
    'Turn realtime rotation on
    cmdRealtimeRotation.Tag = 1
    cmdRealtimeRotation.Caption = "Turn Realtime Rotation Off"
    tmrRotate.Enabled = True
  Else
    'Turn realtime rotation off
    cmdRealtimeRotation.Tag = 0
    cmdRealtimeRotation.Caption = "Turn Realtime Rotation On"
    tmrRotate.Enabled = False
  End If
End Sub

'Rotate timer for realtime rotation
Private Sub tmrRotate_Timer()
  idx = idx + 1
  'If angle is greater thn 359°, reset it to 0°
  If idx > 359 Then idx = 0
  'Rotate picture
  BitRotate picSource, picDestination, idx, chkAntiAlias.Value
  'Display current angle in degrees
  Me.Caption = "Angle: " & Trim(Str(idx))
End Sub

Private Sub Form_Load()
  'Ensure that program is compiled before running
  If App.LogMode = 0 Then
    MsgBox "Compile Me - I'll Run Faster"
    End
  End If
End Sub
