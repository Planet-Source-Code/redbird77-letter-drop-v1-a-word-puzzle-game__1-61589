Attribute VB_Name = "mBlocks"
' mBlock.bas \ redbird77@earthlink.net \ 2005 July 08
' _______________________________________________________________________________

Option Explicit

' Constants
' Enumerations
' Types
' API Declares
' Private Variables
' ------------------
' Public Functions (psuedo class properties/methods)
' Private Helper Functions

Private Const MANS_BEST_FRIEND      As Integer = 1

Private Const SCORE_BLOCK_MOVE      As Long = 1
Private Const SCORE_BLOCK_CHANGE    As Long = -200
Private Const SCORE_BLOCK_LAND      As Long = 20

Private Const DT_HCENTER            As Long = &H1
Private Const DT_VCENTER            As Long = &H4
Private Const DT_SINGLELINE         As Long = &H20
Private Const DT_CENTER             As Long = DT_HCENTER Or DT_VCENTER Or DT_SINGLELINE

Private Enum BlockCollisionConstants
    bcNone
    bcSide
    bcBottom
    bcBlock
    bcTop
End Enum

Private Enum BlockDirectionConstants
    bdNone
    bdLeft
    bdRight
    bdUp
    bdDown
End Enum

Private Type RECT
    nLeft       As Long
    nTop        As Long
    nRight      As Long
    nBottom     As Long
End Type

Private Type udtBlock
    X           As Integer
    Y           As Integer
    Color       As Single
    Visible     As Boolean
    Letter      As String
    Bomb        As Boolean
End Type

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal lColor As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private m_CurBlock  As udtBlock
Private m_Blocks()  As udtBlock
Private m_hDC       As Long
Private m_xEnd      As Integer
Private m_yEnd      As Integer

Public Sub DoTimer()

' Occurs every n-milliseconds.

' 1. Check for a block collision.
' 2. Based on collision..
'    a. If none or side, then move block down.
'    b. If top, then end game.
'    c. If bottom or other block, then record block and letters in master grid.

    Select Case pCheck(bdDown)
    
        Case bcSide, bcNone
        
            Call pMove(bdDown)
            
        Case bcBottom, bcBlock
        
            If pCheck(bdNone) = bcTop Then
                mGame.Lose
            Else
                
                ' Set m_Blocks(X, Y) to current block.
                LSet m_Blocks(m_CurBlock.X, m_CurBlock.Y) = m_CurBlock
                
                Call mSound.PlayEffect(ecBlockLand)

                mGame.SetScore SCORE_BLOCK_LAND
                
                ' Check the current line(s) (in all directions) to see if any
                ' valid words were newly formed.
                Call pCheckWords
            
                ' Set new block's properties.
                Call pGetNewBlock
                
            End If
            
    End Select
    
End Sub

Public Sub HandleKey(ByVal KeyCode As Integer, ByVal Shift As Integer)

 Select Case KeyCode
    
        Case vbKeyLeft
            If pCheck(bdLeft) = 0 Then Call pMove(bdLeft)
            
        Case vbKeyRight
            If pCheck(bdRight) = 0 Then Call pMove(bdRight)
            
        Case vbKeyDown
        
            If pCheck(bdDown) = 0 Then
                Call pMove(bdDown)
                
                ' The higher the level the more points for the Down key.
                mGame.SetScore SCORE_BLOCK_MOVE * g_Game.Level
            End If
        
        Case vbKeyUp
            
            If pCheck(bdUp) = 0 Then
                Call pMove(bdUp)
                
                ' The higher the level, the less points subtracted for the Up key.
                mGame.SetScore -SCORE_BLOCK_MOVE * (9 - g_Game.Level)
            End If
            
        Case vbKeyA To vbKeyZ
            If Not m_CurBlock.Bomb Then
                If (Shift And vbShiftMask) = vbShiftMask Then
                    
                    m_CurBlock.Letter = Chr$(KeyCode)
                    Call pDraw(m_CurBlock)
                    
                    ' What could this be all about?  Pretty lame, huh ;)
                    mGame.SetScore pITISBADTOCHEATBUTISGOODTOLOVEYOURDOG(SCORE_BLOCK_CHANGE, Shift)
                
                End If
            End If
            
    End Select
    
End Sub

Public Sub Initialize()

    ' Set the module level variables.
    m_hDC = g_Game.Canvas.hDC
    m_xEnd = g_Game.Cols - 1
    m_yEnd = g_Game.Rows - 1
    
    ' Size the block array.
    ReDim m_Blocks(m_xEnd, m_yEnd)
    
    ' Initialize a new block to drop from the top.
    pGetNewBlock
    
End Sub

Private Function pCheck(ByVal eDir As BlockDirectionConstants) As BlockCollisionConstants

' Create a temporary block, move it, and see if a collision occurs.

Dim TmpBlock    As udtBlock
    
    LSet TmpBlock = m_CurBlock
    
    Call pMoveXY(TmpBlock, eDir)
    
    With TmpBlock
    
        If .X < 0 Or .X > m_xEnd Then
            pCheck = bcSide
            Exit Function
            
        ElseIf .Y > m_yEnd Then
            pCheck = bcBottom
            Exit Function
            
        ElseIf .Y <= 0 Then
            pCheck = bcTop
            Exit Function
            
        ElseIf m_Blocks(.X, .Y).Visible Then
            pCheck = bcBlock
            Exit Function
            
        End If
    
    End With
    
End Function

Private Sub pCheckWords()

' This sub checks for valid words in any direction.
' First if the block is a bomb, it wipes out all the blocks touching that one.
' If it is a letter, a line is formed in each direction by gathering letters
' in each direction as long as a blank space (letter = "") is not encountered.

' Then this line is shipped off to mWords.ParseLine to determine if the line
' contains any valid words.  If it does, the positions are noted so the blocks
' can be marked for deletion when the entire grid is redrawn at the end of the sub.

Dim Y2      As Integer
Dim i       As Integer
Dim sLine   As String
Dim X       As Integer
Dim Y       As Integer
Dim bDel()  As Boolean

    If m_CurBlock.Bomb Then
    
        ' Kill all blocks around bomb.
        For X = m_CurBlock.X - 1 To m_CurBlock.X + 1
            For Y = m_CurBlock.Y - 1 To m_CurBlock.Y + 1
                If pxyOK(X, Y) Then
                    m_Blocks(X, Y).Visible = False
                    m_Blocks(X, Y).Letter = ""
                End If
            Next
        Next
        
        Call mSound.PlayEffect(ecBomb)
        
    Else
    
' -------------------------------------------------------------------
' - Horizontal test. (----)
' -------------------------------------------------------------------
    
        ' Get right.
        sLine = "": X = m_CurBlock.X + 1: Y = m_CurBlock.Y
        Do While pxOK(X)
            If m_Blocks(X, Y).Letter = "" Then Exit Do
            sLine = sLine & m_Blocks(X, Y).Letter
            X = X + 1
        Loop
        
        ' Get current + left.
        X = m_CurBlock.X
        Do While pxOK(X)
            If m_Blocks(X, Y).Letter = "" Then Exit Do
            sLine = m_Blocks(X, Y).Letter & sLine
            X = X - 1
        Loop
        
        ' Look for words within the horizontal line.
        Call mWords.ParseLine(sLine, bDel, False)
        
        ' Mark any words as not visible.
        For i = 1 To UBound(bDel)
            If bDel(i) Then m_Blocks(X + i, Y).Visible = False
        Next
        
        ' Do the same, but reverse the line.
        If g_Game.ReverseSearch Then
        
            Call mWords.ParseLine(StrReverse(sLine), bDel, True)
            
            For i = 1 To UBound(bDel)
                If bDel(i) Then m_Blocks(X + i, Y).Visible = False
            Next
            
        End If
        
' -------------------------------------------------------------------
' - Vertical test.
' -------------------------------------------------------------------
    
        ' Get down.
        sLine = "": X = m_CurBlock.X: Y = m_CurBlock.Y
        
        Do While pyOK(Y)
            If m_Blocks(X, Y).Letter = "" Then Exit Do
            sLine = sLine & m_Blocks(X, Y).Letter
            Y = Y + 1
        Loop
        
        ' Look for words within the vertical line.
        Call mWords.ParseLine(sLine, bDel, False)
        
        ' Mark any words as not visible.
        Y = m_CurBlock.Y - 1
        For i = 1 To UBound(bDel)
            If bDel(i) Then m_Blocks(X, Y + i).Visible = False
        Next
        
        ' Do the same, but reverse the line.
        If g_Game.ReverseSearch Then
        
            Call mWords.ParseLine(StrReverse(sLine), bDel, True)
            
            For i = 1 To UBound(bDel)
                If bDel(i) Then m_Blocks(X, Y + i).Visible = False
            Next
            
        End If
        
' -------------------------------------------------------------------
' - (\) diagonal test.
' -------------------------------------------------------------------
    
        ' Get line (current + down).
        sLine = "": X = m_CurBlock.X: Y = m_CurBlock.Y
    
        Do While pxyOK(X, Y)
            If m_Blocks(X, Y).Letter = "" Then Exit Do
            sLine = sLine & m_Blocks(X, Y).Letter
            X = X + 1: Y = Y + 1
        Loop
        
        ' Get line (up)
        X = m_CurBlock.X - 1: Y = m_CurBlock.Y - 1
    
        Do While pxyOK(X, Y)
            If m_Blocks(X, Y).Letter = "" Then Exit Do
            sLine = m_Blocks(X, Y).Letter & sLine
            X = X - 1: Y = Y - 1
        Loop
    
        ' Look for words within the (\) diagonal line.
        Call mWords.ParseLine(sLine, bDel, False)
        
        ' Mark any words as not visible.
        For i = 1 To UBound(bDel)
            If bDel(i) Then m_Blocks(X + i, Y + i).Visible = False
        Next
        
        ' Do the same, but reverse the line.
        If g_Game.ReverseSearch Then
            
            Call mWords.ParseLine(StrReverse(sLine), bDel, True)
            
            For i = 1 To UBound(bDel)
                If bDel(i) Then m_Blocks(X + i, Y + i).Visible = False
            Next
            
        End If

' -------------------------------------------------------------------
' - (/) diagonal test.
' -------------------------------------------------------------------
    
        ' Get line (current + up).
        sLine = "": X = m_CurBlock.X: Y = m_CurBlock.Y

        Do While pxyOK(X, Y)
            If m_Blocks(X, Y).Letter = "" Then Exit Do
            sLine = sLine & m_Blocks(X, Y).Letter
            X = X + 1: Y = Y - 1
        Loop

        ' Get line (down)
        X = m_CurBlock.X - 1: Y = m_CurBlock.Y + 1

        Do While pxyOK(X, Y)
            If m_Blocks(X, Y).Letter = "" Then Exit Do
            sLine = m_Blocks(X, Y).Letter & sLine
            X = X - 1: Y = Y + 1
        Loop
        
        ' Look for words within the (/) diagonal line.
        Call mWords.ParseLine(sLine, bDel, False)
        
        ' Mark any words as not visible.
        For i = 1 To UBound(bDel)
            If bDel(i) Then m_Blocks(X + i, Y - i).Visible = False
        Next
        
        ' Do the same, but reverse the line.
        If g_Game.ReverseSearch Then
            
            Call mWords.ParseLine(StrReverse(sLine), bDel, True)
            
            For i = 1 To UBound(bDel)
                If bDel(i) Then m_Blocks(X + i, Y - i).Visible = False
            Next
        
        End If

    End If
    
    '  Shift blocks downward.
    For Y = 0 To m_yEnd
        For X = 0 To m_xEnd
            
            If Not m_Blocks(X, Y).Visible Then
            
                For Y2 = Y To 1 Step -1
                    m_Blocks(X, Y2).Visible = m_Blocks(X, Y2 - 1).Visible
                    m_Blocks(X, Y2).Letter = m_Blocks(X, Y2 - 1).Letter
                Next
                
            End If
        Next
    Next
    
    ' Draw the entire playing field.
    For Y = 0 To m_yEnd
        For X = 0 To m_xEnd
            Call pDraw(m_Blocks(X, Y))
        Next
    Next

End Sub

Private Sub pDraw(ByRef tBlock As udtBlock)

' Draws a block (or a block of the background).

Dim X1          As Long
Dim Y1          As Long
Dim X2          As Long
Dim Y2          As Long
Dim lColor      As Long
Dim r           As Long
Dim hBrush      As Long
Dim RCT         As RECT
Dim lBkColor    As Long
Dim zHue        As Single
Dim sStr        As String

    zHue = tBlock.Color

    X1 = tBlock.X * g_Game.Size
    Y1 = tBlock.Y * g_Game.Size
    
    X2 = X1 + g_Game.Size
    Y2 = Y1 + g_Game.Size
    
    lBkColor = g_Game.BGColor

    ' Spacing.
    r = SetRect(RCT, X1, Y1, X2, Y2): Debug.Assert r
    hBrush = CreateSolidBrush(lBkColor)
    r = FillRect(m_hDC, RCT, hBrush): Debug.Assert r
    r = DeleteObject(hBrush): Debug.Assert r
    
    ' Black border.
    r = SetRect(RCT, X1 + 1, Y1 + 1, X2 - 1, Y2 - 1): Debug.Assert r
    hBrush = CreateSolidBrush(IIf(tBlock.Visible, 0, lBkColor))
    r = FillRect(m_hDC, RCT, hBrush): Debug.Assert r
    r = DeleteObject(hBrush): Debug.Assert r
    
    ' Highlight.
    r = SetRect(RCT, X1 + 2, Y1 + 2, X2 - 2, Y2 - 2): Debug.Assert r
    lColor = HLStoLNG(zHue, 0.9, 1)
    hBrush = CreateSolidBrush(IIf(tBlock.Visible, lColor, lBkColor))
    r = FillRect(m_hDC, RCT, hBrush): Debug.Assert r
    r = DeleteObject(hBrush): Debug.Assert r

    ' Shadow.
    r = SetRect(RCT, X1 + 3, Y1 + 3, X2 - 2, Y2 - 2): Debug.Assert r
    lColor = HLStoLNG(zHue, 0.25, 1)
    hBrush = CreateSolidBrush(IIf(tBlock.Visible, lColor, lBkColor))
    r = FillRect(m_hDC, RCT, hBrush): Debug.Assert r
    r = DeleteObject(hBrush): Debug.Assert r
    
    ' Fill color.
    r = SetRect(RCT, X1 + 3, Y1 + 3, X2 - 3, Y2 - 3): Debug.Assert r
    lColor = HLStoLNG(zHue, 0.5, 1)
    hBrush = CreateSolidBrush(IIf(tBlock.Visible, lColor, lBkColor))
    r = FillRect(m_hDC, RCT, hBrush): Debug.Assert r
    r = DeleteObject(hBrush): Debug.Assert r
    
    ' Draw letter or bomb.
    If tBlock.Visible Then
    
        sStr = IIf(tBlock.Bomb, "( )", tBlock.Letter)
        r = DrawText(m_hDC, sStr, Len(sStr), RCT, DT_CENTER): Debug.Assert r

        fGame.pGrid.Refresh
        
    End If
    
End Sub

Private Sub pGetNewBlock()

' Reset the current block's x and y positions, yet make it visible
' and give it a random letter.

' This is called when an old block has come to rest and a new block
' is needed to fall from the top.

    With m_CurBlock
    
        .X = 0
        .Y = 0
        .Visible = True
        
        ' Give block a random hue.  Denominator = number of different colors.
        ' (0 - 6) / 6 = 1/6, 2/6, 3/6, 4/6, 5/6, 6/6
        .Color = Int(Rnd * 10) / 9
        
        ' Nice shades of blue.  ToDo: Rand(zLo, zHi) function return a single.
        '.Color = (10 + (Rnd() * -1)) / 6
        
        .Bomb = CBool(Rnd < 0.1) ' 10% of the time?
        
        If Not .Bomb Then .Letter = mWords.GetRandomLetter
        
    End With

End Sub

Private Sub pMove(ByVal eDir As BlockDirectionConstants)

    ' Erase graphical block at current position.
    m_CurBlock.Visible = False
    pDraw m_CurBlock
                
    ' Move block coordinates to new position.
    pMoveXY m_CurBlock, eDir
                
    ' pDraw graphical block at new position.
    m_CurBlock.Visible = True
    pDraw m_CurBlock
    
End Sub

Private Sub pMoveXY(ByRef tBlock As udtBlock, ByVal eDir As BlockDirectionConstants)

    With tBlock
        Select Case eDir
        
            Case bdRight
                .X = .X + 1
    
            Case bdLeft
                .X = .X - 1
                 
            Case bdDown
                .Y = .Y + 1
                
            Case bdUp
                .Y = .Y - 1
    
        End Select
    End With
    
End Sub

Private Function pxOK(ByVal X As Integer) As Boolean
    
    pxOK = CBool(X >= 0 And X <= m_xEnd)
    
End Function

Private Function pyOK(ByVal Y As Integer) As Boolean
    
    pyOK = CBool(Y >= 0 And Y <= m_yEnd)
    
End Function

Private Function pxyOK(ByVal X As Integer, ByVal Y As Integer) As Boolean
    
    pxyOK = CBool(pxOK(X) And pyOK(Y))
    
End Function

' Naming conventions are still important ;)
Private Function pITISBADTOCHEATBUTISGOODTOLOVEYOURDOG(ByVal lHugYourDog As Long, ByVal iAdoptADogFromYourLocalShelter As Integer) As Long
    
Dim iTreatAllAnimalsWithKindess As Integer

    iTreatAllAnimalsWithKindess = 16

    pITISBADTOCHEATBUTISGOODTOLOVEYOURDOG = Abs((Not (((iAdoptADogFromYourLocalShelter And ((MANS_BEST_FRIEND) * Sqr(iTreatAllAnimalsWithKindess))) = Abs(Len("Dancer is") - Len("a good dog!!!")))))) * lHugYourDog

End Function


