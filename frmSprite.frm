VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSprite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sprite Tutorial"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   282
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ani 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   4320
      Picture         =   "frmSprite.frx":0000
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   450
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   6750
   End
   Begin VB.PictureBox d_mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   4320
      Picture         =   "frmSprite.frx":EDEA
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox d_sprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   4320
      Picture         =   "frmSprite.frx":14D18
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Timer tmrDraw 
      Interval        =   10
      Left            =   4200
      Top             =   1320
   End
   Begin VB.Timer tmrAni 
      Interval        =   100
      Left            =   4200
      Top             =   1800
   End
   Begin VB.Frame Frame2 
      Caption         =   "Animation"
      Height          =   735
      Left            =   1200
      TabIndex        =   7
      Top             =   1440
      Width           =   2655
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         SelStart        =   1
         Value           =   10
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change sprite on direction"
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   2655
      Begin VB.CommandButton cmdDir 
         Caption         =   "Right"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "Left"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "Down"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "Up"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   1
      Left            =   360
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   1440
      Width           =   675
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   360
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   0
      Top             =   360
      Width           =   675
   End
End
Attribute VB_Name = "frmSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sDir As Long 'the sprite's direction
Dim sFrame As Long 'the current frame of the animation

Private Sub cmdDir_Click(Index As Integer)
sDir = Index 'set the direction of the unit
End Sub

Private Sub cmdExit_Click()
tmrAni.Enabled = False
tmrDraw.Enabled = False
board(0).Cls
board(1).Cls
Unload Me
End Sub

Private Sub Form_Load()
sDir = 1 'set the unit's direction so you can see it's face
End Sub

Private Sub sldSpeed_Click()
tmrAni.Interval = 1100 - (sldSpeed.Value * 100) 'set the speed for the animation
End Sub

Private Sub tmrAni_Timer()
sFrame = sFrame + 1: If sFrame > 9 Then sFrame = 0 'update the frame
End Sub

Private Sub tmrDraw_Timer()
'draw the direction box
board(0).Cls
BitBlt board(0).hDC, 0, 0, 45, 45, d_mask.hDC, 45 * sDir, 0, SRCAND
BitBlt board(0).hDC, 0, 0, 45, 45, d_sprite.hDC, 45 * sDir, 0, SRCINVERT

'draw the animation box
board(1).Cls
BitBlt board(1).hDC, 0, 0, 45, 45, ani.hDC, 45 * sFrame, 0, SRCCOPY
End Sub
