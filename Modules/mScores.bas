Attribute VB_Name = "mScores"
' mScores.bas \ redbird77@earthlink.net \ 2005 July 07
' _______________________________________________________________________________
'
' A _simple_ implementation of a _simple_ high score list.
'
' ToDo:
' I would like to make
'  1.  Good Dog    1000
'      Happy Cat   1000
'  2.  Fancy Fish   750
' ...
' 10.  Fluffy Bunny 125
'      Wild Goose   125

Option Explicit

Private Const SCORES_COUNT  As Integer = 10

' For proper alignment name must be <= 10 characters when display count <= 99.
Private Const NAME_FORMAT   As String = "!@@@@@@@@@@@@@@"
Private Const SCORES_FORMAT As String = "@@@@@@@@@@"

Private Type udtHiScore
    
    Name    As String
    Score   As Long
    
    Level   As Integer
    Words   As Long
    
End Type

Private m_Scores()  As udtHiScore

Public Function Add(ByVal lScore As Long, ByVal sName As String, _
                    ByVal iLevel As Integer, ByVal lWords As Long) As Boolean

' Scores are stored/added by score descending.

Dim i   As Integer
Dim j   As Integer

    For i = 0 To UBound(m_Scores)

        If lScore > m_Scores(i).Score Then

            ' From the end of the list to just before the incoming score position,
            ' move each list entry down a position.
            For j = UBound(m_Scores) To i + 1 Step -1
                LSet m_Scores(j) = m_Scores(j - 1)
            Next

            ' Insert incoming.
            With m_Scores(i)
                .Score = lScore
                .Name = sName
                .Level = iLevel
                .Words = lWords
            End With
            
            Add = True
            Exit For

        End If
    Next

End Function

Public Sub Clear()

' ToDo: Make backup of HiScores.dat before clear.

    ReDim m_Scores(SCORES_COUNT - 1)
    
End Sub

Public Sub Display(ByRef List As ListBox)

' ToDo: Something nicer than a standard issue ListBox with a plain monospaced font.

Dim i   As Integer
Dim s   As String

    List.Clear
    
    List.AddItem Format$("NAME", NAME_FORMAT) & " " & _
                 Format$("SCORE", SCORES_FORMAT) & " " & _
                 Format$("LEVEL", SCORES_FORMAT) & " " & _
                 Format$("WORDS", SCORES_FORMAT)

    For i = 0 To UBound(m_Scores)

        With m_Scores(i)
        
            'If .Name <> "" Then
            s = Right$("  " & (i + 1), 2) & ". "
            
            List.AddItem Format$(s & .Name, NAME_FORMAT) & " " & _
                         Format$(.Score, SCORES_FORMAT) & " " & _
                         Format$(.Level, SCORES_FORMAT) & " " & _
                         Format$(.Words, SCORES_FORMAT)
            'End If
            
        End With
        
    Next
    
End Sub

Public Function IsValid(ByVal lScore As Long) As Boolean

' Is the score high enough to make it on the list?

    IsValid = CBool(lScore > m_Scores(UBound(m_Scores)).Score)
    
End Function

Public Function LoadScores(ByVal sFile As String) As Boolean

' Load from HiScores.dat.  Called once at start of application.

Dim f   As Integer

    ReDim m_Scores(SCORES_COUNT - 1)

    f = FreeFile
    
    Open sFile For Binary As #f
        Get #f, , m_Scores()
    Close #f
    
End Function

Public Function SaveScores(ByVal sFile As String) As Boolean

' Save scores in HiScores.dat.

Dim f   As Integer

    f = FreeFile
    
    Open sFile For Binary As #f
        Put #f, , m_Scores()
    Close #f

End Function
