Attribute VB_Name = "mINI"
' mScores.bas \ redbird77@earthlink.net \ 2005 July 07
' _______________________________________________________________________________
'
' A couple of basic INI file procedures (and trying out a little different naming convention.)

Option Explicit

Private Declare Function API_WriteString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function API_GetString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function DelValue(ByVal sSect As String, ByVal sKey As String, ByVal sFile As String) As Long

    DelValue = API_WriteString(sSect, sKey, vbNullString, sFile)
    
    Debug.Assert DelValue
    
End Function

Public Function GetValue(ByVal sSect As String, ByVal sKey As String, ByVal sFile As String, Optional ByVal sDefault As String = "") As String

Dim sBuf    As String
Dim lRet    As Long
    
    sBuf = String$(1024, vbNullChar)
    
    ' lRet = copied characters excluding \0.
    lRet = API_GetString(sSect, sKey, sDefault, sBuf, Len(sBuf), sFile)
    
    If lRet Then GetValue = Left$(sBuf, lRet)
    
End Function

Public Function PutValue(ByVal sSect As String, ByVal sKey As String, ByVal sVal As String, ByVal sFile As String) As Long

    PutValue = API_WriteString(sSect, sKey, sVal, sFile)
    
    Debug.Assert PutValue
    
End Function
