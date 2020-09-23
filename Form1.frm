VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logic Boom - Dance Beat Machine"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   9870
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   6285
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   4680
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3960
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add &Column"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5880
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1800
      Width           =   1935
   End
   Begin Logicboom.Timeline Timeline3 
      Height          =   435
      Left            =   2160
      TabIndex        =   13
      Top             =   1800
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   767
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1080
      Width           =   1935
   End
   Begin Logicboom.Timeline Timeline2 
      Height          =   435
      Left            =   2160
      TabIndex        =   11
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   767
   End
   Begin VB.CommandButton Command2 
      DownPicture     =   "Form1.frx":104D0
      Height          =   615
      Left            =   9000
      Picture         =   "Form1.frx":112E2
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Stop"
      Top             =   5520
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "Form1.frx":120F4
      Height          =   615
      Left            =   8280
      Picture         =   "Form1.frx":12F06
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Play"
      Top             =   5520
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   4320
      Top             =   4200
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   5760
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Min             =   1
      Max             =   355
      SelStart        =   75
      TickStyle       =   3
      Value           =   75
      TextPosition    =   1
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Reverb"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin Logicboom.Timeline Timeline1 
      Height          =   435
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   767
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin Logicboom.Timeline Timeline4 
      Height          =   435
      Left            =   2160
      TabIndex        =   16
      Top             =   2520
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   767
   End
   Begin Logicboom.Timeline Timeline5 
      Height          =   435
      Left            =   2160
      TabIndex        =   17
      Top             =   3240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   767
   End
   Begin Logicboom.Timeline Timeline6 
      Height          =   435
      Left            =   2160
      TabIndex        =   18
      Top             =   3960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   767
   End
   Begin Logicboom.Timeline Timeline7 
      Height          =   435
      Left            =   2160
      TabIndex        =   19
      Top             =   4680
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   767
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "75"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3945
      TabIndex        =   8
      Top             =   5520
      Width           =   210
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1920
      TabIndex        =   6
      Top             =   5520
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   5250
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Beat Timeline"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   2160
      TabIndex        =   3
      Top             =   90
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Instrument"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   1050
   End
   Begin VB.Image Image7 
      Height          =   150
      Left            =   120
      Picture         =   "Form1.frx":13948
      Top             =   120
      Width           =   9600
   End
   Begin VB.Image Image4 
      Height          =   150
      Left            =   120
      Picture         =   "Form1.frx":1848A
      Top             =   5280
      Width           =   9600
   End
   Begin VB.Image Image5 
      Height          =   3840
      Left            =   3840
      Picture         =   "Form1.frx":1CFCC
      Top             =   0
      Width           =   3840
   End
   Begin VB.Image Image3 
      Height          =   3840
      Left            =   3840
      Picture         =   "Form1.frx":2D192
      Top             =   3840
      Width           =   3840
   End
   Begin VB.Image Image2 
      Height          =   3840
      Left            =   7680
      Picture         =   "Form1.frx":3D358
      Top             =   0
      Width           =   3840
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   0
      Picture         =   "Form1.frx":4D51E
      Top             =   3840
      Width           =   3840
   End
   Begin VB.Image Image6 
      Height          =   3840
      Left            =   7680
      Picture         =   "Form1.frx":5D6E4
      Top             =   3840
      Width           =   3840
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveas 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFileHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpSpecs 
         Caption         =   "Program &Specs"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim dx As New DirectX8
Dim perf As DirectMusicPerformance8
Dim loader As DirectMusicLoader8
Dim style As DirectMusicStyle8
Dim Band As DirectMusicBand8
Dim composer As DirectMusicComposer8
Dim seg As DirectMusicSegment8
Dim segBand As DirectMusicSegment8
Dim segMotif() As DirectMusicSegment8
Dim mediapath As String

Dim mtTime As Long
Dim CurBand As Integer

'Private Sub cmdExit_Click()
'On Error Resume Next
'    Stop_Click
'    Unload Me
'End Sub

'Private Sub Drum_Click(Index As Integer)
'On Error Resume Next
'    Call perf.PlaySegmentEx(segMotif(Index), DMUS_SEGF_SECONDARY, 0)
'End Sub

'Private Sub EDIT_Tempo_KeyPress(KeyAscii As Integer)
'On Error Resume Next
'    If KeyAscii = vbKeyReturn Then
'        If Val(EDIT_Tempo.Text) > 0 And Val(EDIT_Tempo.Text) < 1001 And IsNumeric(EDIT_Tempo.Text) Then
'            UpDown_Tempo.Value = EDIT_Tempo.Text
'            ChangeTempo (UpDown_Tempo.Value)
'        Else
'            EDIT_Tempo.Text = UpDown_Tempo.Value
'        End If
'    End If
'    If KeyAscii = vbKeyReturn Then KeyAscii = 0
'End Sub

'Private Sub EDIT_Tempo_LostFocus()
'On Error Resume Next
'    If Val(EDIT_Tempo.Text) > 0 And Val(EDIT_Tempo.Text) < 1001 And IsNumeric(EDIT_Tempo.Text) Then
'        UpDown_Tempo.Value = EDIT_Tempo.Text
'        ChangeTempo (UpDown_Tempo.Value)
'    Else
'        EDIT_Tempo.Text = UpDown_Tempo.Value
'   End If
'End Sub

Private Sub EDIT_Volume_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(EDIT_Volume.Text) And Val(EDIT_Volume.Text) >= 0 And Val(EDIT_Volume.Text) < 101 Then
            UpDown_Volume.Value = EDIT_Volume.Text
            ChangeVolume UpDown_Volume.Value
        Else
            EDIT_Volume.Text = UpDown_Volume.Value
        End If
    End If
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub EDIT_Volume_LostFocus()
On Error Resume Next
    If IsNumeric(EDIT_Volume.Text) And Val(EDIT_Volume.Text) >= 0 And Val(EDIT_Volume.Text) < 101 Then
        UpDown_Volume.Value = EDIT_Volume.Text
        ChangeVolume UpDown_Volume
    Else
        EDIT_Volume.Text = UpDown_Volume.Value
    End If

End Sub

Private Sub Command1_Click()
CurBand = 0
Timer1.Enabled = True
End Sub

Private Sub Check1_Click()
    'Ok, they want to switch the default audio paths
    Dim dmPath As DirectMusicAudioPath8

    If Check1.Value = vbUnchecked Then
        Set dmPath = perf.CreateStandardAudioPath(DMUS_APATH_DYNAMIC_STEREO, 128, True)
    Else
        Set dmPath = perf.CreateStandardAudioPath(DMUS_APATH_SHARED_STEREOPLUSREVERB, 128, True)
    End If
    perf.SetDefaultAudioPath dmPath
    ChangeBands
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Timeline1.Refresh
Timeline2.Refresh
Timeline3.Refresh
Timeline4.Refresh
Timeline5.Refresh
Timeline6.Refresh
Timeline7.Refresh
End Sub

Private Sub Command3_Click()
Timeline1.AddBand
Timeline2.AddBand
Timeline3.AddBand
Timeline4.AddBand
Timeline5.AddBand
Timeline6.AddBand
Timeline7.AddBand
Timeline1.Refresh
Timeline2.Refresh
Timeline3.Refresh
Timeline4.Refresh
Timeline5.Refresh
Timeline6.Refresh
Timeline7.Refresh
End Sub

Private Sub Form_Load()
CurBand = 0
Timeline1.Refresh
Timeline2.Refresh
Timeline3.Refresh
Timeline4.Refresh
Timeline5.Refresh
Timeline6.Refresh
Timeline7.Refresh
    Dim dmA As DMUS_AUDIOPARAMS, lCount As Long
    Dim MotifName As String
    
    mediapath = FindMediaDir("Drums!.sgt")
    
    Set perf = dx.DirectMusicPerformanceCreate()
    Set loader = dx.DirectMusicLoaderCreate()
    Set composer = dx.DirectMusicComposerCreate()
    
    'Make sure we can init the audio as well
    On Error GoTo FailedInit
    ' Initialize performance object to use its own DirectSound object
    perf.InitAudio Me.hWnd, DMUS_AUDIOF_ALL, dmA, , DMUS_APATH_SHARED_STEREOPLUSREVERB, 128
    
    ' SetMasterAutoDownload indicates we the perofmance object
    ' to attempt to auto download DLS collections when reference in
    ' sgt and sty files
    Call perf.SetMasterAutoDownload(True)
    
    Set style = loader.LoadStyle(mediapath & "drums!.sty")

    Set seg = loader.LoadSegment(mediapath & "drums!.sgt")
    
    Get_Bands

    'LIST_Grooves.AddItem ("Alternative")
    'LIST_Grooves.AddItem ("Blues")
    'LIST_Grooves.AddItem ("Country")
    'LIST_Grooves.AddItem ("Dance - Pop")
    'LIST_Grooves.AddItem ("Hard Rock")
    'LIST_Grooves.AddItem ("Hip Hop")
    'LIST_Grooves.AddItem ("Jazz")
    'LIST_Grooves.AddItem ("Latin")
    'LIST_Grooves.AddItem ("R & B")
    'LIST_Grooves.AddItem ("Rap")
    'LIST_Grooves.AddItem ("Soft Rock")
    'LIST_Grooves.AddItem ("World")
    
For x = 0 To 6
With Combo1(x)
    .AddItem "Kick"
    .AddItem "Snare"
    .AddItem "Low Tom"
    .AddItem "Mid Tom"
    .AddItem "High Tom"
    .AddItem "Ride"
    .AddItem "Splash"
    .AddItem "Crash"
    .AddItem "Low Conga"
    .AddItem "High Conga"
    .AddItem "Timbale"
    .AddItem "Agogo"
    .AddItem "Guiro"
    .AddItem "Low Block"
    .AddItem "High Block"
    .AddItem "Cuica"
    .AddItem "Triangle"
    .AddItem "Shaker"
    .AddItem "Castanets"
    .AddItem "Jingle Bells"
    .AddItem "Tambourine"
    .AddItem "Hand Clap"
    .AddItem "Sticks"
    .AddItem "Scratch"
    .AddItem "High Q"
    .ListIndex = x
End With
Next x
    
    ' Download the default band so that we can play the drum pads immediately
    ChangeBands
    'ChangeVolume UpDown_Volume.Value
    
    ReDim segMotif(style.GetMotifCount() - 1)
    For lCount = 0 To style.GetMotifCount() - 1
        MotifName = style.GetMotifName(lCount)
        'We could set the drum name here (but we'll just leave them hard coded)
        'Drum(lCount).Caption = MotifName
        Set segMotif(lCount) = style.GetMotif(MotifName)
    Next
    
    'LIST_Grooves.ListIndex = 0
    'LIST_Bands.ListIndex = 0
    
    Exit Sub
    
FailedInit:
    MsgBox "Could not initialize DirectMusic." & vbCrLf & "This sample will exit.", vbOKOnly Or vbInformation, "Exiting..."
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Dim lCount As Long
    
    On Error Resume Next
    If Not (segBand Is Nothing) Then
        perf.StopEx segBand, 0, 0
        segBand.Unload perf.GetDefaultAudioPath
    End If
    If Not (seg Is Nothing) Then perf.StopEx seg, 0, 0
    Set seg = Nothing
    For lCount = LBound(segMotif) To UBound(segMotif)
        If Not (segMotif(lCount) Is Nothing) Then perf.StopEx segMotif(lCount), 0, 0
        Set segMotif(lCount) = Nothing
    Next
    Set segBand = Nothing
    Set style = Nothing
    Set composer = Nothing
    Set loader = Nothing
    If Not (Band Is Nothing) Then
        Call Band.Unload(perf)
    End If
    Set Band = Nothing
    If Not (perf Is Nothing) Then perf.CloseDown
    Set perf = Nothing

End Sub

Private Sub Get_Bands()
On Error Resume Next
    Dim BandCount As Integer
    Dim counter As Integer
    BandCount = style.GetBandCount()
    For counter = 0 To (BandCount - 1)
    '    Combo1.AddItem (style.GetBandName(BandCount - counter - 1))
    Next counter
End Sub

Private Sub LIST_Bands_Click()
On Error Resume Next
    ChangeBands
End Sub

Private Sub LIST_Grooves_Click()
On Error Resume Next
    perf.SetMasterGrooveLevel ((LIST_Grooves.ListIndex * 8) + 1)
End Sub

Private Sub Play_Click()
On Error Resume Next
    PlaySeg
    ChangeBands
    chkReverb.Enabled = False
End Sub

Private Sub Stop_Click()
On Error Resume Next
    perf.StopEx seg, 0, 0
    chkReverb.Enabled = True
End Sub

Private Sub UPDOWN_Tempo_Change()
On Error Resume Next
    EDIT_Tempo.Text = UpDown_Tempo.Value
    ChangeTempo (UpDown_Tempo.Value)
End Sub

'Private Sub UPDOWN_Volume_Change()
'On Error Resume Next
'    EDIT_Volume.Text = UpDown_Volume.Value
'    Call ChangeVolume(UpDown_Volume.Value)
'End Sub

Private Sub ChangeBands()
On Error Resume Next
    If Not (Band Is Nothing) Then
        Call Band.Unload(perf)
    End If

    If LIST_Bands = vbNullString Then
        Set Band = style.GetBand("Standard")
    Else
        Set Band = style.GetBand(LIST_Bands)
    End If
    Call Band.Download(perf)
    Set segBand = Band.CreateSegment()
    segBand.Download perf.GetDefaultAudioPath
    Call perf.PlaySegmentEx(segBand, DMUS_SEGF_SECONDARY, 0)
End Sub

Private Sub PlaySeg()
On Error Resume Next
    Call perf.PlaySegmentEx(seg, 0, 0)
End Sub

Private Sub ChangeTempo(Tempo As Integer)
On Error Resume Next
    perf.SendTempoPMSG 0, DMUS_PMSGF_REFTIME, Tempo
End Sub

Sub ChangeVolume(ByVal n As Long)
On Error Resume Next
    If n = 0 Then
        n = -10000
    Else
        n = (-50 * (100 - n))
    End If

    perf.SetMasterVolume n
End Sub

Private Sub mnuHelpAbout_Click()
Form2.Show vbModal
End Sub

Private Sub mnuHelpSpecs_Click()
On Error Resume Next
Shell "notepad " & App.Path & "\spcs.dat", vbNormalFocus
End Sub

Private Sub Slider1_Change()
Label5.Caption = Slider1.Value - Slider1.Value - Slider1.Value
Timer1.Interval = (Slider1.Value - Slider1.Value - Slider1.Value) + 1
End Sub

Private Sub Slider1_Scroll()
Label5.Caption = Slider1.Value - Slider1.Value - Slider1.Value
Timer1.Interval = (Slider1.Value - Slider1.Value - Slider1.Value) + 1
End Sub

Private Sub Timer1_Timer()
If CurBand < Timeline1.BandCount Then
CurBand = CurBand + 1
Timeline1.Highlight CurBand
Timeline2.Highlight CurBand
Timeline3.Highlight CurBand
Timeline4.Highlight CurBand
Timeline5.Highlight CurBand
Timeline6.Highlight CurBand
Timeline7.Highlight CurBand

    If Timeline1.InitBand = True Then
    Call perf.PlaySegmentEx(segMotif(Combo1(0).ListIndex), DMUS_SEGF_SECONDARY, 0)
    End If
    
    If Timeline2.InitBand = True Then
    Call perf.PlaySegmentEx(segMotif(Combo1(1).ListIndex), DMUS_SEGF_SECONDARY, 0)
    End If
    
    If Timeline3.InitBand = True Then
    Call perf.PlaySegmentEx(segMotif(Combo1(2).ListIndex), DMUS_SEGF_SECONDARY, 0)
    End If
    
    If Timeline4.InitBand = True Then
    Call perf.PlaySegmentEx(segMotif(Combo1(3).ListIndex), DMUS_SEGF_SECONDARY, 0)
    End If
    
    If Timeline5.InitBand = True Then
    Call perf.PlaySegmentEx(segMotif(Combo1(4).ListIndex), DMUS_SEGF_SECONDARY, 0)
    End If
    
    If Timeline6.InitBand = True Then
    Call perf.PlaySegmentEx(segMotif(Combo1(5).ListIndex), DMUS_SEGF_SECONDARY, 0)
    End If
    
    If Timeline7.InitBand = True Then
    Call perf.PlaySegmentEx(segMotif(Combo1(6).ListIndex), DMUS_SEGF_SECONDARY, 0)
    End If

Else
CurBand = 0
Timer1_Timer
End If
End Sub
