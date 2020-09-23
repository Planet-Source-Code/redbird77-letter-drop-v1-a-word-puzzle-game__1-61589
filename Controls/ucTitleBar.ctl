VERSION 5.00
Begin VB.UserControl ucTitleBar 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2370
   ScaleHeight     =   420
   ScaleWidth      =   2370
End
Attribute VB_Name = "ucTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ucTitleBar.ctl \ redbird77@earthlink.net \ 2005 July 06
' ___________________________________________________________________
'
' A simple TitleBar control for a form with no titlebar.
'
' Possible enhancements:
' 1. More exposed properties and events.  More customizeable.
' 2. More buttons other than "exit".  Perhaps an enumeration of standard
'    titlebar buttons (minimize, maximize, system menu, help, etc.)
' 3. To add to #2, some nifty buttons like always-on-top and minimize-to-
'    system-tray.
' 4. All the above buttons represented with icons or text. (+ tooltips).
' 5. An Align property like a picturebox.
' 6. But of course - gradients!

Option Explicit

Private Const HTCAPTION        As Long = 2
Private Const WM_NCLBUTTONDOWN As Long = &HA1

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const DEF_CAPTION           As String = "Caption"
Private Const DEF_CAPTIONFORECOLOR  As Long = vbWindowText
Private Const DEF_CAPTIONBACKCOLOR  As Long = vbButtonFace
Private Const DEF_PADDING           As Long = 6

Private m_lCaptionForeColor As OLE_COLOR
Private m_lCaptionBackColor As OLE_COLOR
Private m_sCaption          As String
Private m_oFont             As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_lPadding          As Long
Private m_lWidth            As Long

Private Sub UserControl_Show()

    ' Must I wait til here to get this info, or I get a "Client Site Not
    ' Available" error?
    m_lWidth = UserControl.Parent.Width
    
End Sub

Private Sub UserControl_Terminate()

    Set m_oFont = Nothing

End Sub

Private Sub UserControl_InitProperties()

' This sub is called only once, when the control is first placed on a form.
' Subsequently, the UserControl_Paint event is fired.

    ' Set the default UserControl properties.
    m_sCaption = DEF_CAPTION
    m_lCaptionForeColor = DEF_CAPTIONFORECOLOR
    m_lCaptionBackColor = DEF_CAPTIONBACKCOLOR
    m_lPadding = DEF_PADDING
    
    ' The default font is the parent's font.
    Set m_oFont = Ambient.Font
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
Dim r   As Long

    ' Clicked in caption-area, move parent form.
    If X < UserControl.Width - UserControl.TextWidth("X") Then
    
        r = ReleaseCapture()
        r = SendMessage(UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
    
    ' Clicked in exit-area, close parent form.
    Else
    
        Unload UserControl.Parent
    
    End If

End Sub

Private Sub UserControl_Paint()

    With UserControl

        Set .Font = m_oFont
        .Height = .TextHeight(m_sCaption) + (m_lPadding * Screen.TwipsPerPixelY)
    
        ' Draw caption.
        .Cls
        
        .CurrentX = (4 * Screen.TwipsPerPixelY) ' 4 is hardcoded for now.
        .CurrentY = (.Height \ 2) - (.TextHeight(m_sCaption) \ 2)
        
        .ForeColor = m_lCaptionForeColor
        .BackColor = m_lCaptionBackColor
        
        UserControl.Print m_sCaption
        
        ' Draw exit.
        .CurrentX = (.Width - (4 * Screen.TwipsPerPixelX)) - .TextWidth("X")
        .CurrentY = (.Height \ 2) - (.TextHeight("X") \ 2)
        
        UserControl.Print "X"
        
        ' Draw border.
        UserControl.Line (0, 0)- _
                         (.Width - 1 * Screen.TwipsPerPixelX, _
                          .Height - 1 * Screen.TwipsPerPixelY), m_lCaptionForeColor, B
        
    End With

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

' Still trying to get the whole Font property implemented correctly.

    m_sCaption = PropBag.ReadProperty("Caption", DEF_CAPTION)
    m_lCaptionForeColor = PropBag.ReadProperty("CaptionForeColor", DEF_CAPTIONFORECOLOR)
    m_lCaptionBackColor = PropBag.ReadProperty("CaptionBackColor", DEF_CAPTIONBACKCOLOR)
    Set m_oFont = PropBag.ReadProperty("Font", Ambient.Font)
    m_lPadding = PropBag.ReadProperty("Padding", DEF_PADDING)
    
    Call UserControl.Refresh
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", m_sCaption, DEF_CAPTION)
    Call PropBag.WriteProperty("CaptionForeColor", m_lCaptionForeColor, DEF_CAPTIONFORECOLOR)
    Call PropBag.WriteProperty("CaptionBackColor", m_lCaptionBackColor, DEF_CAPTIONBACKCOLOR)
    Call PropBag.WriteProperty("Font", m_oFont) ', DEF_FONT)
    Call PropBag.WriteProperty("Padding", m_lPadding, DEF_PADDING)
    
End Sub

Public Property Get Caption() As String

    Caption = m_sCaption

End Property

Public Property Let Caption(ByVal n As String)

    m_sCaption = n
    
    Call UserControl.PropertyChanged("Caption")
    Call UserControl.Refresh

End Property

Public Property Get CaptionBackColor() As OLE_COLOR

    CaptionBackColor = m_lCaptionBackColor
    
End Property

Public Property Let CaptionBackColor(ByVal n As OLE_COLOR)

    m_lCaptionBackColor = n

    Call UserControl.PropertyChanged("CaptionBackColor")
    Call UserControl.Refresh

End Property

Public Property Get CaptionForeColor() As OLE_COLOR

    CaptionForeColor = m_lCaptionForeColor

End Property

Public Property Let CaptionForeColor(ByVal n As OLE_COLOR)

    m_lCaptionForeColor = n

    Call UserControl.PropertyChanged("CaptionForeColor")
    Call UserControl.Refresh

End Property

Public Property Get Font() As StdFont

    Set Font = m_oFont

End Property

Public Property Set Font(ByRef n As StdFont)

    With m_oFont
        .Bold = n.Bold
        .Italic = n.Italic
        .Name = n.Name
        .Size = n.Size
    End With

    Call UserControl.PropertyChanged("Font")
    Call UserControl.Refresh

End Property

Public Property Get Padding() As Long

    Padding = m_lPadding

End Property

Public Property Let Padding(ByVal n As Long)

    m_lPadding = n

    Call UserControl.PropertyChanged("Padding")
    Call UserControl.Refresh
    
End Property
