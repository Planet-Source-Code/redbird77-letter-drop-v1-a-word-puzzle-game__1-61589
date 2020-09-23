VERSION 5.00
Begin VB.Form fGame 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7560
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7350
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   StartUpPosition =   2  'CenterScreen
   Begin pLetterDrop.ucTitleBar ucTitleBar1 
      Height          =   405
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   714
      Caption         =   "LETTER DROP v1"
      CaptionForeColor=   0
      CaptionBackColor=   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Dungeon"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraControls 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   4575
      TabIndex        =   8
      Top             =   5700
      Width           =   2325
      Begin VB.Label lblScores 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HI SCORES"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         TabIndex        =   11
         Top             =   840
         Width           =   2325
      End
      Begin VB.Label lblSettings 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SETTINGS"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         TabIndex        =   10
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label lblPlay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PLAY"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2325
      End
   End
   Begin VB.ListBox lstWords 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   4575
      TabIndex        =   3
      Top             =   915
      Width           =   2370
   End
   Begin VB.Timer tmrPlay 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   180
      Top             =   165
   End
   Begin VB.PictureBox pGrid 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Dungeon"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7410
      Left            =   0
      ScaleHeight     =   494
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label lblLevelCap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LEVEL"
      BeginProperty Font 
         Name            =   "Dungeon"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4575
      TabIndex        =   6
      Top             =   5250
      Width           =   1140
   End
   Begin VB.Label lblWordsCap 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WORDS "
      BeginProperty Font 
         Name            =   "Dungeon"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   465
      Width           =   2370
   End
   Begin VB.Label lblScoreCap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SCORE"
      BeginProperty Font 
         Name            =   "Dungeon"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4575
      TabIndex        =   4
      Top             =   4785
      Width           =   1125
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4575
      TabIndex        =   5
      Top             =   4785
      Width           =   2370
   End
   Begin VB.Label lblLevel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4575
      TabIndex        =   7
      Top             =   5250
      Width           =   2370
   End
End
Attribute VB_Name = "fGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' fGame.frm \ redbird77@earthlink.net \ 2005 July 08
' _______________________________________________________________________________

Option Explicit

Private Sub Form_Load()

    Randomize
    
    Set g_Game.Canvas = pGrid
    
     ' Load the game's saved settings from an INI file.
    Call mGame.LoadSettings(App.Path & "\Data\Settings.ini")
    Call ApplySettings
    Call InitalizeGUI
        
    ' Load the high score list.
    Call mScores.LoadScores(App.Path & "\Data\HiScores.dat")
    
    ' Load the word file.
    Call mWords.LoadFile(App.Path & "\Data\Words.txt")
    
    ' Start subclassing.  (Check mUtility.Subclass and mSound.NotifyProc for details.)
    Call mUtility.Subclass(fGame.hWnd)
    
End Sub

Public Sub InitalizeGUI()

' Set, size, and move all the controls associated with the game's GUI.

    With g_Game
    
        Call mGame.ApplySettings
    
        ' Canvas.
        .Canvas.Move 0, ucTitleBar1.Height - 1
    
        ' Words ListBox.
        lstWords.Clear
        lstWords.Move (.Canvas.Left + .Canvas.Width + 8), _
                       lblWordsCap.Top + lblWordsCap.Height + 4
        
        ' Words Caption.
        lblWordsCap.Move lstWords.Left, ucTitleBar1.Height + 4
        
        ' Score Caption.
        lblScoreCap.Move lstWords.Left, lstWords.Top + lstWords.Height + 4
        
        ' Score Value.
        lblScore.Move lblScoreCap.Left, lblScoreCap.Top
        
        ' Level Caption.
        lblLevelCap.Move lstWords.Left, lblScoreCap.Top + lblScoreCap.Height + 4
        
        ' Level Value.
        lblLevel.Move lblLevelCap.Left, lblLevelCap.Top
        
        ' Form.
        Me.Width = (lstWords.Left + lstWords.Width + 6) * Screen.TwipsPerPixelX
        Me.Height = (.Canvas.Top + .Canvas.Height) * Screen.TwipsPerPixelY
        
        ' Buttons.
        fraControls.Move lstWords.Left, Me.Height \ Screen.TwipsPerPixelY - fraControls.Height
    
        ' TitleBar.
        ucTitleBar1.Width = Me.Width \ Screen.TwipsPerPixelX
       
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call mSound.StopMusic
    Call mUtility.UnSubclass(fGame.hWnd)
    Call mScores.SaveScores(App.Path & "\Data\HiScores.dat")
    
End Sub

Private Sub pGrid_KeyDown(KeyCode As Integer, Shift As Integer)

    Call mBlocks.HandleKey(KeyCode, Shift)
    
End Sub

Private Sub tmrPlay_Timer()

    Call mBlocks.DoTimer
    
End Sub

Private Sub lblPlay_Click()

' Handles the Play/Pause/Resume bit.
' In my future? class version, perhaps a OnModeChange Event and a
' PlayModeConstants enumeration (pmPause, pmRun, etc.).  This way I could
' the mode would not be tied to a label control.  I could represent it any way
' I wanted, but it would be stored within the class.

    If lblPlay.Caption = "PLAY" Then
    
        Call mGame.Start
        lblPlay.Caption = "PAUSE"
        
    Else
    
        If lblPlay.Caption = "RESUME" Then
            lblPlay.Caption = "PAUSE"
        Else
            lblPlay.Caption = "RESUME"
        End If
        
        tmrPlay.Enabled = Not tmrPlay.Enabled
        
    End If
    
End Sub

Private Sub lblScores_Click()

     ' Only show the scores if the game is not active.
    If lblPlay.Caption <> "PLAY" Then Exit Sub

    fHiScores.Show vbModal, fGame
    
End Sub

Private Sub lblSettings_Click()

    ' Only show the settings the game is not active.
    If lblPlay.Caption <> "PLAY" Then Exit Sub
    
    fSettings.Show vbModal, Me
    
End Sub
