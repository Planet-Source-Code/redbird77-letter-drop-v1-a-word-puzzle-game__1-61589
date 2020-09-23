Attribute VB_Name = "mUtility"
' mUtility.bas \ redbird77@earthlink.net \ 2005 July 08
' _______________________________________________________________________________
'
' Various uncategorizeable "helper" procedures.

Option Explicit

Private Const GWL_WNDPROC   As Long = -4

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public m_lpProcOld As Long

Public Function Normalize(ByVal uVal As Single, _
                          ByVal uLo As Single, ByVal uHi As Single, _
                          ByVal nLo As Single, ByVal nHi As Single) As Long
                          
' Return a value in a specified output range based it's position in a
' specified input range.  Thanks to Dr. Math for help with this function.
    
    Normalize = nLo + (uVal - uLo) * (nHi - nLo) / (uHi - uLo)

End Function

Public Sub Pause(ByVal zDelay As Single)

' Pauses execution for "zDelay" seconds.

Dim zEnd    As Single

    zEnd = Timer + zDelay
    
    Do
        'DoEvents
    Loop Until Timer > zEnd
    
End Sub

Public Function HLStoLNG(ByVal H As Single, ByVal L As Single, ByVal s As Single) As Long

Dim M1  As Single
Dim M2  As Single
Dim R   As Single
Dim G   As Single
Dim B   As Single

    If s = 0 Then
        R = L
        B = L
        G = L
    Else
        If L <= 0.5 Then
            M2 = L * (1 + s)
        Else
            M2 = L + s - L * s
        End If

        M1 = 2 * L - M2

        R = V(M1, M2, H + 1 / 3)
        G = V(M1, M2, H)
        B = V(M1, M2, H - 1 / 3)
    End If
    
    HLStoLNG = RGB(R * 255, G * 255, B * 255)

End Function

Private Function V(ByVal M1 As Single, ByVal M2 As Single, ByVal H As Single) As Single
    
    If H > 1 Then H = H - 1
    If H < 0 Then H = H + 1
    
    If (6 * H < 1) Then
        V = (M1 + (M2 - M1) * H * 6)
    ElseIf (2 * H < 1) Then
        V = M2
    ElseIf (3 * H < 2) Then
        V = (M1 + (M2 - M1) * ((2 / 3) - H) * 6)
    Else
        V = M1
    End If
    
End Function

Public Function Subclass(ByVal hWnd As Long) As Long

' OK, I know only a slight bit more re: subclassing than my dog, but I think I
' finally got this simple method down.  (Please correct me!)

' What we are doing is giving the window with the handle "hWnd" (in this case
' fGame) a new address for it's window procedure (the procedure that takes care
' of processing all messages).

' We are giving it the address of mSound.NotifyProc and also saving it's
' original value - "m_lpProcOld".  We will assign this original value back
' before we exit the application.

' In the meantime, the new window procedure will take care of message processing.
' However I'm really only interested in one message in this app - "MM_MCINOTIFY".
' This message will be sent when you do something with a MIDI - play it, queue the
' status, etc. (as long as you use the "notify" flag in the command string).

' When I get this notification, I check to see if it was successful.  If so, I
' re-start the MIDI.

' If any other (there are zillions) messages are sent, I call CallWindowProc
' with the original window procedure address - "m_lpProcOld" to let the
' default processing of those messages take place.

    m_lpProcOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf mSound.NotifyProc)
    
End Function

Public Function UnSubclass(ByVal hWnd As Long) As Long

    UnSubclass = SetWindowLong(hWnd, GWL_WNDPROC, m_lpProcOld)
    
End Function
