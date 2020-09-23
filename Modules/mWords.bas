Attribute VB_Name = "mWords"
' mWords.bas \ redbird77@earthlink.net \ 2005 July 07
' _______________________________________________________________________________
'
' Manages the word file.
'
' ToDo:
' Make the 3 To 5 bit a setting.  Read thru a file and gather the word lengths.

Option Explicit

Private Const SCORE_LETTER_HI   As Long = 80
Private Const SCORE_LETTER_LO   As Long = 10
Private Const SCORE_BONUS       As Long = 100

Private Type udtLetterFrequency
    Letter      As String
    Frequency   As Single
End Type

Private Type udtWordList
    Words()         As String
    Offsets(3 To 5) As Long         ' Change (3 To 5) if other word lengths needed.
End Type

Private m_FreqsF(25)    As udtLetterFrequency
Private m_FreqsA(25)    As udtLetterFrequency
Private m_Words         As udtWordList

Public Function GetRandomLetter() As String

' Return a random letter based on the frequencies of the letters in the word file.

Dim i       As Integer
Dim zRand   As Single
Dim zTotal  As Single

    zRand = Rnd
    
    For i = UBound(m_FreqsF) To 0 Step -1
        zTotal = zTotal + m_FreqsF(i).Frequency
        If zRand < zTotal Then Exit For
    Next

    GetRandomLetter = m_FreqsF(i).Letter
    
End Function

Public Function LoadFile(ByVal sFile As String) As Long

' 1. Load the word file into an array.
' 2. Store the indicies where the 3-letter, 4-letter, etc. words start.

Dim f       As Integer
Dim i       As Long
Dim iLen    As Integer
Dim iLenOld As Integer

    ' Load word file into an array.
    f = FreeFile
    
    Open sFile For Input As #f
        m_Words.Words = Split(UCase$(Input(LOF(f), #f)), vbCrLf)
    Close #f
    
    ' For each word..
    For i = 0 To UBound(m_Words.Words)
    
        iLen = Len(m_Words.Words(i))
        
        ' If the word's length differs from the old length, then
        ' record that as the offset for that length.
        If iLen <> iLenOld Then m_Words.Offsets(iLen) = i
        
        iLenOld = iLen
        
    Next
    
    Call pAnalyzeFile
    
End Function

Public Sub ParseLine(ByVal sLine As String, ByRef bDel() As Boolean, ByVal bRev As Boolean)

' This procedure takes a line (generated from the letters of contiguous blocks)
' and parses it, looking for 3 to 5 letter length words.

' It is called every time a new letter is dropped.

Dim i           As Long
Dim j           As Integer
Dim k           As Integer
Dim sWord       As String
Dim iCurLen     As Integer
Dim lEnd        As Long
Dim lScore      As Long
Dim lScoreChr   As Long

    ReDim bDel(1 To Len(sLine))
    
    For i = 1 To UBound(bDel)
        bDel(i) = False
    Next

    ' For each possible word length..
    For iCurLen = LBound(m_Words.Offsets) To UBound(m_Words.Offsets)
    
        ' Starting with each character in the line..
        For j = 1 To Len(sLine) - iCurLen + 1
        
            ' Create a "word" that is iCurLen characters in length.
            sWord = Mid$(sLine, j, iCurLen)
            
            'fGame.lstWords.AddItem "Testing: " & sWord
            
            If iCurLen = UBound(m_Words.Offsets) Then
                lEnd = UBound(m_Words.Words)
            Else
                lEnd = m_Words.Offsets(iCurLen + 1) - 1
            End If
            
            ' Look through word list for a match, but only look at words
            ' that are the same length.
            For i = m_Words.Offsets(iCurLen) To lEnd
            
                ' If match, exit after marking the blocks to delete.
                If sWord = m_Words.Words(i) Then
                
                    ' Reset score.
                    lScore = 0
            
                    ' k iterates thru each characer in word.
                    For k = 0 To iCurLen - 1
                        
                        ' j = index of first character in word.
                        
                        If bRev Then
                            bDel((Len(sLine) + 1) - (j + k)) = True
                        Else
                            bDel(j + k) = True
                        End If
                        
                        lScoreChr = Normalize(m_FreqsA(Asc(Mid$(sWord, k + 1, 1)) - 65).Frequency, _
                                              m_FreqsF(0).Frequency, m_FreqsF(25).Frequency, _
                                              SCORE_LETTER_LO, SCORE_LETTER_HI)
                        
                        'fGame.lstWords.AddItem m_FreqsA(Asc(Mid$(sWord, k + 1, 1)) - 65).Letter & " worth " & lScoreChr & " points."
                        
                        lScore = lScore + lScoreChr
                    Next
                    
                    ' ToDo: Whenever I get around to making all this with classes.
                    ' I might have an OnFound event and add the word to the ListBox
                    ' in the main form, like cWord_OnFound(sWord As String)
                    fGame.lstWords.AddItem Right$("     " & sWord, 5) & " - " & lScore
                    
                    ' Keep last word made as currently selected.
                    fGame.lstWords.ListIndex = fGame.lstWords.NewIndex
                    
                    ' Get a bonus for longer letter words.
                    ' Bonus =   0 for 3 letters.
                    '       = 100 for 4 letters.
                    '       = 200 for 5 letters.
                    
                    ' ToDo: Un-hardcode "3".
                    Call mGame.SetScore(lScore + SCORE_BONUS * (iCurLen - 3))
                    Call mSound.PlayEffect(ecWordFound)
                    
                    g_Game.Words = g_Game.Words + 1
                    
                    Exit For
                    
                End If
            Next
        Next
    Next

End Sub

Private Sub pAnalyzeFile()

' Calcuate the frequencies of each letter in the 'words.txt' file.

' Yes, this is very English-centric.  I'm from Texas, whatta ya'll expect ;).

Dim i       As Long
Dim j       As Long
Dim iChr    As Integer
Dim lChars  As Long
Dim Freq    As udtLetterFrequency
Dim z       As Single

    z = Timer

    ' For each word in file..
    For i = 0 To UBound(m_Words.Words)

        ' Keep a running total of the occurrences of each character.
        For j = 1 To Len(m_Words.Words(i))
        
            iChr = Asc(Mid$(m_Words.Words(i), j, 1)) - 65
            m_FreqsF(iChr).Frequency = m_FreqsF(iChr).Frequency + 1

        Next

        ' Update the total character count.
        lChars = lChars + j - 1
        
    Next
    
    ' Calculate the frequency of each letter.
    For i = 0 To UBound(m_FreqsF)
    
        With m_FreqsF(i)
            .Letter = Chr$(i + 65)
            .Frequency = .Frequency / lChars
        End With
        
    Next
    
    ' Make copy to keep sorted by letter (A-Z) not by frequency.
    ' (Note: This is so I can easily lookup a frequency by letter when
    ' calculating the score.  I know I could ditch the 2 arrays in favor
    ' of a collection with the letter as the key... put that on my ToDo list.)
    For i = 0 To UBound(m_FreqsF)
    
        With m_FreqsA(i)
            .Letter = m_FreqsF(i).Letter
            .Frequency = m_FreqsF(i).Frequency
        End With
        
    Next
    
    ' Sort by frequency (descending).
    For i = 0 To UBound(m_FreqsF) - 1
        For j = i + 1 To UBound(m_FreqsF)

            If m_FreqsF(i).Frequency < m_FreqsF(j).Frequency Then

                LSet Freq = m_FreqsF(i)
                LSet m_FreqsF(i) = m_FreqsF(j)
                LSet m_FreqsF(j) = Freq

            End If

        Next
    Next

'    For i = 0 To UBound(m_FreqsF)
'        fGame.lstWords.AddItem m_FreqsF(i).Letter & ": " & Format$(m_FreqsF(i).Frequency, "Percent")
'    Next
    
'    Debug.Print "pAnalyzeFile: " & Format$(Timer - z, "0.000")
    
End Sub
