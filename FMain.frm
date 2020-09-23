VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TrackMouseDEMO"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "FMAIN.frx":0000
      Top             =   120
      Width           =   5295
   End
   Begin VB.PictureBox picLink 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   2520
      Width           =   5295
      Begin VB.Label lblLink 
         Alignment       =   2  'Center
         Caption         =   "visit decadence of evolution"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents tmLink As CTrackMouse
Attribute tmLink.VB_VarHelpID = -1
Private Sub Form_Load()
Set tmLink = New CTrackMouse
Set tmLink.TrackObject = picLink
End Sub


Private Sub tmLink_MouseOut()
With lblLink
    .ForeColor = 0
    With .Font
        .Underline = False
    End With
End With
End Sub

Private Sub tmLink_MouseOver()
With lblLink
    .ForeColor = &HFF0000
    With .Font
        .Underline = True
    End With
End With
End Sub


