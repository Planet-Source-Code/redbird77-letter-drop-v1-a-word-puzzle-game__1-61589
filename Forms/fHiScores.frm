VERSION 5.00
Begin VB.Form fHiScores 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3525
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6150
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
   ScaleHeight     =   3525
   ScaleWidth      =   6150
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   3855
      TabIndex        =   2
      Top             =   3090
      Width           =   1065
   End
   Begin pLetterDrop.ucTitleBar ucTitleBar1 
      Height          =   405
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   714
      Caption         =   "HI SCORES"
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4995
      TabIndex        =   3
      Top             =   3090
      Width           =   1065
   End
   Begin VB.ListBox lstScores 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   75
      TabIndex        =   1
      Top             =   480
      Width           =   5985
   End
End
Attribute VB_Name = "fHiScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' fHiScores.frm \ redbird77@earthlink.net \ 2005 July 08
' _______________________________________________________________________________

Option Explicit

Private Sub Form_Load()
                
    Call mScores.Display(lstScores)
    
End Sub

Private Sub cmdOK_Click()

    Unload Me
    
End Sub

Private Sub cmdReset_Click()

    If MsgBox("Are you sure?", vbYesNo + vbExclamation, "Confirm Reset") = vbYes Then
    
        Call mScores.Clear
        Call mScores.Display(lstScores)
    
    End If

End Sub
