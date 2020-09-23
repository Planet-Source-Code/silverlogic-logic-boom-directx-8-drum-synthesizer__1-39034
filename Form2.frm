VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Logic Boom"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Uses Microsoft(R) DirectX(R) 8.0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   360
      Picture         =   "Form2.frx":030A
      Top             =   3840
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E8C1A8&
      X1              =   240
      X2              =   5760
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Logic Boom Copyright 2002(C) Logic Multimedia"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   720
      Picture         =   "Form2.frx":0614
      Top             =   120
      Width           =   4500
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label2.Caption = Label2.Caption & vbCrLf & vbCrLf & "Permission is expressly granted to redistribute this software in a bundle, online, or any other digital storage, as long as all original files remain intact, and unchanged, exept for saved sound files. You are free to use this program as you wish." & vbCrLf & vbCrLf & "http://www.LogicMedia.tk/"
End Sub
