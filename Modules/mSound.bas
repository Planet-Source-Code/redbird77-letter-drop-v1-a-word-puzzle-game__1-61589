Attribute VB_Name = "mSound"
' mSound.bas \ redbird77@earthlink.net \ 2005 July 08
' _______________________________________________________________________________
'
' This module contains functions to play/stop a midi and a wave file.
'
' Before I attempted this, I had never written a line of code to play any type
' of media before, so take the code with a grain of salt ;)
'
' Things I've learned that might help others...
' 1. All mci, mid, wavXXX functions do not accept file names with spaces.
'    Use the GetShortPath API function or surround the file name with quotes.
'
' 2. I could have nixed the PlaySound function and only used the mciXXX
'    functions to play the effects and music, but letting PlaySound do all
'    the work for me was nice.
'
' 3. I don't know if it necessary to first stop the midi before you close it,
'    but I got the desired results and I didn't crash VB so I kept it in.
'
' 4. Looping a midi... didn't know what a task that would be.  I decided to
'    go about the right(?) way and use subclassing and not a timer.
'    If I wanted to be notified of multiple events (open, play, seek, etc.),
'    would I have to use multiple VB forms or create my own windows, or
'    rig up some Select Case with different NotifyProc addresses...?? since there
'    seems to be no way of telling what command string triggered the notification.

Option Explicit

' PlaySound flag constants.
Private Const SND_ASYNC             As Long = &H1
Private Const SND_NODEFAULT         As Long = &H2

' Media Control Interface (mci) messages and constants.
Private Const MM_MCINOTIFY          As Long = &H3B9
Private Const MCI_NOTIFY_ABORTED    As Long = &H4
Private Const MCI_NOTIFY_FAILURE    As Long = &H8
Private Const MCI_NOTIFY_SUCCESSFUL As Long = &H1
Private Const MCI_NOTIFY_SUPERSEDED As Long = &H2

Public Enum EffectConstants
    ecBomb
    ecBlockLand
    ecWordFound
    ecLose
End Enum

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Sub PlayEffect(ByVal eEffect As EffectConstants)
    
' This could be easily customizeable.  You could load the paths into an array
' read from a INI file.  Or you could store the sounds in a resource file.

    If g_Game.Effects = 0 Then Exit Sub

    Select Case eEffect

        Case ecBomb
            Call pPlayEffect(App.Path & "\Effects\Bomb.wav")

        Case ecBlockLand
            Call pPlayEffect(App.Path & "\Effects\Block.wav")
            
        Case ecWordFound
            Call pPlayEffect(App.Path & "\Effects\Found.wav")
            
    End Select
    
End Sub

Public Sub PlayMusic()

' No need for parameter since one and only one midi will be playing at a
' time and its alias will always be "MUSIC".
  
' ToDo: Read default (binary?) midi from resource file if not user-specified.

    With g_Game
    
        If .Music = 0 Then Exit Sub
    
        If .MusicFile = "" Then
            
            MsgBox "No midi file specified.  Go to ""Settings"" to select a midi to play as background music.", _
                   vbExclamation, "Error."
        Else
        
            ' Always located in Music folder.
            Call pPlayMusic(App.Path & "\Music\" & .MusicFile)
        
        End If
        
    End With
    
End Sub

Public Function StopMusic() As Long
    
    pStopMusic
      
End Function

Private Function pPlayEffect(ByVal sFile As String) As Long

    pPlayEffect = PlaySound(sFile, 0&, SND_ASYNC Or SND_NODEFAULT)
    
    Debug.Assert pPlayEffect
    
End Function

Private Sub pError(ByVal lErr As Long, ByVal sCmd As String)

Dim lRet    As Long
Dim sErr    As String

    lRet = mciGetDeviceID("MUSIC")
    
    ' If still valid alias, stop and close the device.
    If lRet Then
    
        ' Could this cause a circular reference?
        pStopMusic
        
        ' Display error message.
        sErr = String$(256, vbNullChar)
        lRet = mciGetErrorString(lErr, sErr, Len(sErr))
        MsgBox "Command String: " & sCmd & vbCrLf & vbCrLf & _
               "MCI Error String: " & Left$(sErr, InStr(sErr, vbNullChar) - 1), vbExclamation, "Error Occurred"
    
    Else
    
        ' Alias never got associated with device, so nothing to do(?).
        
    End If
        
End Sub

Private Function pPlayMusic(ByVal sFile As String) As Long

Dim sCmd    As String
Dim lRet    As Long

    ' Open device with alias MUSIC.
    sCmd = "open """ & sFile & """" & " type sequencer alias MUSIC"
    lRet = mciSendString(sCmd, 0&, 0, 0)
    
    ' If error..
    If lRet Then
    
        Call pError(lRet, sCmd)
    
    ' Else play (sending notification messages to fGame).
    Else
    
        sCmd = "play MUSIC notify"
        lRet = mciSendString(sCmd, 0&, 0, fGame.hWnd)
        If lRet Then Call pError(lRet, sCmd)
        
    End If
    
End Function

Private Function pStopMusic() As Long

Dim sCmd    As String
Dim lRet    As Long
Dim sStatus As String

    ' Get status of device.
    sStatus = String$(256, vbNullChar)
    sCmd = "status MUSIC mode"
    lRet = mciSendString(sCmd, sStatus, Len(sStatus), 0)
    
    ' If error..
    If lRet Then
    
        Call pError(lRet, sCmd)
        
    ' Else status was successfully retrieved.
    Else
    
        Select Case Left$(sStatus, InStr(sStatus, vbNullChar) - 1)
        
            Case "stopped"
    
                'MsgBox "Music is stopped.  Now closing."
                lRet = mciSendString("close MUSIC", 0&, 0, 0)
            
            Case "playing"
        
                'MsgBox "Music is playing.  Now stopping and closing."
                lRet = mciSendString("stop MUSIC", 0&, 0, 0)
                lRet = mciSendString("close MUSIC", 0&, 0, 0)
    
        End Select
        
    End If
    
End Function

Public Function NotifyProc(ByVal hWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If lMsg = MM_MCINOTIFY Then

        ' If midi finished playing successfully..
        If wParam = MCI_NOTIFY_SUCCESSFUL Then
        
            ' Close device.
            Call StopMusic
            
            ' Reopen.
            Call PlayMusic
            
        End If

    End If
    
    ' I'm just checking out the status, so let the original window
    ' procedure take care of things.
    NotifyProc = CallWindowProc(m_lpProcOld, hWnd, lMsg, wParam, lParam)

End Function
