VERSION 5.00
Begin VB.UserControl Timeline 
   ClientHeight    =   1380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10230
   ScaleHeight     =   1380
   ScaleWidth      =   10230
   ToolboxBitmap   =   "Timeline.ctx":0000
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   440
      Left            =   6120
      TabIndex        =   4
      ToolTipText     =   "Scroll Left"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   440
      Left            =   5880
      TabIndex        =   3
      ToolTipText     =   "Scroll Right"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C00000&
      Height          =   440
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   105
         TabIndex        =   5
         Top             =   0
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   105
         TabIndex        =   2
         Top             =   0
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   4440
         ScaleHeight     =   345
         ScaleWidth      =   105
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
   End
End
Attribute VB_Name = "Timeline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TIMELINE CONTROL
'COPYRIGHT 2002(C) NICK BRABANT & LOGIC MULTIMEDIA
'LOGICMEDIA.TK

Dim N1 As Integer
Dim N2 As Integer
Public BandCount
Public InitBand As Boolean

Private Sub Command1_Click()
For x = 0 To (Picture2.Count - 1)
Picture2(x).Left = Picture2(x).Left - Picture2(x).Width
Next x
End Sub

Private Sub Command2_Click()
If Picture2(1).Left < 0 Then
For x = 0 To (Picture2.Count - 1)
Picture2(x).Left = Picture2(x).Left + Picture2(x).Width
Next x
End If
End Sub

Private Sub Picture2_Click(Index As Integer)
N1 = Index
N2 = Index
If Int(N1 / 5) <> (N2 / 5) Then
    If Picture2(Index).BackColor = &HC0C0C0 Then
    Picture2(Index).BackColor = &HFF0000
    Picture2(Index).Tag = &HFF0000
    Else
    Picture2(Index).BackColor = &HC0C0C0
    Picture2(Index).Tag = &HC0C0C0
    End If
Else
    If Picture2(Index).BackColor = &H808080 Then
    Picture2(Index).BackColor = &H800000
    Picture2(Index).Tag = &H800000
    Else
    Picture2(Index).BackColor = &H808080
    Picture2(Index).Tag = &H808080
    End If
End If
End Sub

Private Sub UserControl_Initialize()
BandCount = 20
End Sub

Private Sub UserControl_Resize()
UserControl.Height = Picture1.Height
Picture1.Width = UserControl.Width - (Command1.Width + Command2.Width)
Command1.Left = Picture1.Width
Command2.Left = Command1.Left + Command1.Width
End Sub

Public Function Highlight(Band As Integer)
On Error Resume Next
'For y = 1 To Picture2.Count
'Refresh
'Next y
'Picture2(Band).BackColor = vbYellow
Picture3.Visible = True
Picture3.ZOrder 0
Picture3.Left = Picture2(Band).Left

If Picture2(Band).Tag = "" Then
InitBand = False
Else
If Picture2(Band).Tag = &HFF0000 Then
InitBand = True
GoTo Theend
Else
InitBand = False
End If

If Picture2(Band).Tag = &H800000 Then
InitBand = True
GoTo Theend
Else
InitBand = False
End If
End If

Theend:
DoEvents
End Function

Public Function Refresh()
On Error Resume Next
Picture3.Visible = False
For y = 1 To Picture2.Count
If Int(y / 5) = (y / 5) Then
    If Picture2(y).Tag = "" Then
    Picture2(y).BackColor = &H808080
    Else
    Picture2(y).BackColor = Picture2(y).Tag
    End If
Else
    If Picture2(y).Tag = "" Then
    Picture2(y).BackColor = &HC0C0C0
    Else
    Picture2(y).BackColor = Picture2(y).Tag
    End If
End If
Next y
End Function

Public Function AddBand()
Load Picture2(Picture2.Count)
With Picture2(Picture2.Count - 1)
    .Top = 0
    .Left = Picture2(Picture2.Count - 2).Left + Picture2(Picture2.Count - 2).Width - 15
    .Visible = True
    .ZOrder 1
End With
BandCount = Picture2.Count - 1
End Function
