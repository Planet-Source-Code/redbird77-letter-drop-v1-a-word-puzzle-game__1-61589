VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fSettings 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5955
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3375
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   3375
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   1860
      Top             =   2385
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   5505
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2220
      TabIndex        =   20
      Top             =   5505
      Width           =   1065
   End
   Begin VB.Frame fraPlay 
      Caption         =   "Game Play"
      Height          =   1230
      Left            =   75
      TabIndex        =   17
      Top             =   4185
      Width           =   3210
      Begin VB.CheckBox chkSearch 
         Caption         =   "Search for Words Backwards"
         Height          =   375
         Left            =   165
         TabIndex        =   18
         Top             =   285
         Width           =   2745
      End
      Begin VB.Label lblCap 
         Caption         =   "Note: This causes more words to be found."
         Height          =   450
         Index           =   4
         Left            =   165
         TabIndex        =   21
         Top             =   645
         Width           =   2340
      End
   End
   Begin VB.Frame fraAudio 
      Caption         =   "Audio"
      Height          =   1485
      Left            =   75
      TabIndex        =   12
      Top             =   2595
      Width           =   3210
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   315
         Left            =   2715
         TabIndex        =   16
         Top             =   1035
         Width           =   390
      End
      Begin VB.TextBox txtMusic 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1035
         Width           =   2520
      End
      Begin VB.CheckBox chkAudio 
         Caption         =   "Background Music"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chkAudio 
         Caption         =   "Sound Effects"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraCanvas 
      Caption         =   "Playing Field"
      Height          =   2055
      Left            =   75
      TabIndex        =   1
      Top             =   465
      Width           =   3210
      Begin VB.ComboBox ddlLevel 
         Height          =   315
         Left            =   705
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1635
         Width           =   525
      End
      Begin VB.TextBox txtSize 
         Height          =   315
         Left            =   705
         TabIndex        =   8
         Text            =   "50"
         Top             =   1170
         Width           =   525
      End
      Begin VB.TextBox txtCols 
         Height          =   315
         Left            =   1740
         TabIndex        =   6
         Text            =   "10"
         Top             =   735
         Width           =   525
      End
      Begin VB.TextBox txtRows 
         Height          =   315
         Left            =   705
         TabIndex        =   4
         Text            =   "10"
         Top             =   735
         Width           =   525
      End
      Begin VB.Label lblBGColorCap 
         Caption         =   "Background Color"
         Height          =   240
         Left            =   675
         TabIndex        =   3
         Top             =   330
         Width           =   1650
      End
      Begin VB.Label lblBGColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   165
         TabIndex        =   2
         Top             =   270
         Width           =   375
      End
      Begin VB.Label lblCap 
         Caption         =   "Level:"
         Height          =   240
         Index           =   3
         Left            =   165
         TabIndex        =   10
         Top             =   1620
         Width           =   765
      End
      Begin VB.Label lblCap 
         Caption         =   "Size:"
         Height          =   240
         Index           =   2
         Left            =   165
         TabIndex        =   9
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblCap 
         Caption         =   "Cols:"
         Height          =   240
         Index           =   1
         Left            =   1305
         TabIndex        =   7
         Top             =   780
         Width           =   765
      End
      Begin VB.Label lblCap 
         Caption         =   "Rows:"
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   5
         Top             =   780
         Width           =   765
      End
   End
   Begin pLetterDrop.ucTitleBar ucTitleBar1 
      Height          =   405
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   714
      Caption         =   "SETTINGS"
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
End
Attribute VB_Name = "fSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' fSettings.frm \ redbird77@earthlink.net \ 2005 July 08
' _______________________________________________________________________________

Option Explicit

Private Sub cmdBrowse_Click()

' Select MIDI file to serve as background music.
    
    With cdlMain
    
        .CancelError = False
        .InitDir = App.Path & "\Music\"
        .Filter = "MIDI Files|*.mid"
        
        .ShowOpen
        
        If .FileName <> "" Then txtMusic.Text = .FileTitle
        
    End With
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()

' Set game settings from control values.
    
    With g_Game
    
        .BGColor = lblBGColor.BackColor
        .Rows = txtRows.Text
        .Cols = txtCols.Text
        .Size = txtSize.Text
        .Level = ddlLevel.Text
    
        .Effects = chkAudio(0).Value
        .Music = chkAudio(1).Value
        .MusicFile = txtMusic.Text
        
        .ReverseSearch = chkSearch.Value
        
    End With

    Call mGame.SaveSettings(App.Path & "\Data\Settings.ini")
    Call fGame.InitalizeGUI
    
    Unload Me

End Sub

Private Sub Form_Load()

' Set control values from game settings.

Dim i   As Integer

    For i = 1 To 8: ddlLevel.AddItem i: Next

    With g_Game
    
        lblBGColor.BackColor = .BGColor
        txtRows.Text = .Rows
        txtCols.Text = .Cols
        txtSize.Text = .Size
        ddlLevel.ListIndex = .Level - 1
        
        chkAudio(0).Value = .Effects
        chkAudio(1).Value = .Music
        txtMusic.Text = .MusicFile
        
        chkSearch.Value = .ReverseSearch
        
    End With

End Sub

Private Sub lblBGColor_Click()

On Error GoTo ErrExit

    With cdlMain
    
        .CancelError = True
        .ShowColor
        lblBGColor.BackColor = .Color
        
    End With
    
    Exit Sub
    
ErrExit:

End Sub
