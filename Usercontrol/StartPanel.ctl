VERSION 5.00
Begin VB.UserControl StartPanel 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "StartPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Developed By HACKPRO TM.
Option Explicit

Private Const VER_PLATFORM_WIN32_NT = 2

Private Type SIZE
   cX As Long
   cY As Long
End Type

Public Enum THEMESIZE
    TS_MIN             '// minimum size
    TS_TRUE            '// size without stretching
    TS_DRAW            '// size that theme mgr will use to draw part
End Enum

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128 ' Maintenance string for PSS usage.
End Type

Private Declare Function GetVersionEx Lib "kernel32" _
        Alias "GetVersionExA" ( _
        lpVersionInformation As OSVERSIONINFO) As Long

Private mWindowsNT As Boolean

Private Type RECT
    hLeft        As Long
    hTop         As Long
    hRight       As Long
    hBottom      As Long
End Type

Private Const S_OK = 0

Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal _
    hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" _
    (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function GetThemePartSize Lib "uxtheme.dll" _
    (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, prc As RECT, ByVal eSize As THEMESIZE, _
    psz As SIZE) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" _
    (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
    
Private m_BackColor         As OLE_COLOR            ' Background.
Private m_CaptionColor      As OLE_COLOR            ' Text color of caption.
Private m_Caption           As String               ' Caption text.
Private m_CaptionFont       As StdFont              ' Font used to display header text.

Private Const DT_RIGHT                As Long = &H2
Private Const DT_SINGLELINE           As Long = &H20      ' strip cr/lf from string before draw.
Private Const DT_WORD_ELLIPSIS        As Long = &H40000

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" _
    (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, _
    ByRef lpRect As RECT, ByVal wFormat As Long) As Long

Private m_lPartId  As Long
Private m_lStateId As Long
Private m_sClass   As String
Private FaultDraw  As Boolean

Public Property Get BackColor() As OLE_COLOR
   BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   m_BackColor = New_BackColor
   PropertyChanged "BackColor"
   ReDraw
End Property

Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   ReDraw
End Property

Public Property Get CaptionColor() As OLE_COLOR
   CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
   m_CaptionColor = New_CaptionColor
   PropertyChanged "CaptionColor"
   ReDraw
End Property

Public Property Get CaptionFont() As Font
   Set CaptionFont = m_CaptionFont
End Property

Public Property Set CaptionFont(ByVal vNewCaptionFont As Font)
   Set m_CaptionFont = vNewCaptionFont
   PropertyChanged "CaptionFont"
   ReDraw
End Property

'---------------------------------------------------------------------------------------
' Function  : DrawTheme
' Purpose   : Try to open Uxtheme.dll.
'---------------------------------------------------------------------------------------
Private Function DrawTheme(ByVal hWnd As Long, ByVal hDC As Long, sClass As String, ByVal iPart As Long, ByVal iState As Long, rtRect As RECT, Optional ByVal CloseTheme As Boolean = False) As Boolean
 
  Dim lResult As Long '* Temp Variable.
  Dim hTheme  As Long
 
    '* If a error occurs then or we are not running XP or the visual style is Windows Classic.
On Error GoTo NoXP
    '* Get out hTheme Handle.
    hTheme = OpenThemeData(hWnd, StrPtr(sClass))
    '* Did we get a theme handle?.
    If (hTheme) Then
        '* Yes! Draw the control Background.
        lResult = DrawThemeBackground(hTheme, hDC, iPart, iState, rtRect, rtRect)
        '* If drawing was successful, return true, or false If not.
        DrawTheme = IIf(lResult, False, True)
        FaultDraw = DrawTheme
    Else
        '* No, we couldn't get a hTheme, drawing failed.
        DrawTheme = False
        FaultDraw = False
    End If
    
    '* Close theme.
    If (CloseTheme = True) Then
        CloseThemeData hTheme
    End If
    
    '* Exit the function now.
    Exit Function
NoXP:
    '* An Error was detected, drawing Failed.
    DrawTheme = False
    FaultDraw = False
On Error GoTo 0

End Function

Public Function GetFaultDraw() As Boolean
    
    GetFaultDraw = FaultDraw
    
End Function

Private Function GetOsSystemVersion()

  Dim OS As OSVERSIONINFO
    
    OS.dwOSVersionInfoSize = Len(OS)
    GetVersionEx OS
    GetOsSystemVersion = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Function

Public Property Get hWnd()
    
    hWnd = UserControl.hWnd
    
End Property

Private Property Get PartWidth(Optional ByVal eWidthOptions As THEMESIZE = TS_TRUE) As Long
  
  Dim tSize As SIZE
  Dim tR As RECT
  Dim hTheme As Long
  Dim lR As Long
  
    hTheme = OpenThemeData(hWnd, StrPtr(m_sClass))
    If (hTheme) Then
        lR = GetThemePartSize(hTheme, hDC, m_lPartId, m_lStateId, tR, eWidthOptions, tSize)
        If (lR = S_OK) Then
            PartWidth = tSize.cX
        End If
        CloseThemeData hTheme
   End If
   
End Property

Private Property Get PartHeight(Optional ByVal eWidthOptions As THEMESIZE = TS_TRUE) As Long
  
  Dim tSize As SIZE
  Dim tR As RECT
  Dim hTheme As Long
  Dim lR As Long
  
    hTheme = OpenThemeData(hWnd, StrPtr(m_sClass))
    If (hTheme) Then
        lR = GetThemePartSize(hTheme, hDC, m_lPartId, m_lStateId, tR, eWidthOptions, tSize)
        If (lR = S_OK) Then
            PartHeight = tSize.cY
        End If
        CloseThemeData hTheme
    End If
   
End Property

Public Function PutForm(ByRef fForm As Form)

  Dim lRect As RECT

On Error Resume Next
    lRect.hLeft = 0
    lRect.hTop = 0
    lRect.hBottom = fForm.ScaleHeight
    lRect.hRight = fForm.ScaleWidth
    DrawTheme fForm.hWnd, fForm.hDC, "StartPanel", 6, 1, lRect, True
On Error GoTo 0

End Function

Public Sub ReDraw()
    
  Dim lRect As RECT

    If (m_CaptionFont Is Nothing) Then
        Set m_CaptionFont = Ambient.Font
    End If
    
    Set UserControl.Font = m_CaptionFont
    UserControl.BackColor = m_BackColor
    If (GetOsSystemVersion = True) Then
        m_sClass = "StartPanel"
        m_lStateId = 1
        With UserControl
            .Cls
            .AutoRedraw = True
            
            lRect.hLeft = 0
            lRect.hTop = 0
            lRect.hBottom = .ScaleHeight
            lRect.hRight = .ScaleWidth + 1
            DrawTheme .hWnd, hDC, m_sClass, 6, 1, lRect, True
            
            lRect.hBottom = 32
            lRect.hLeft = 0
            lRect.hTop = 0
            lRect.hRight = .ScaleWidth
            ' UserPane
            DrawTheme .hWnd, hDC, m_sClass, 1, 1, lRect, True
            
            lRect.hBottom = .ScaleHeight
            lRect.hTop = 33
                                   
            lRect.hLeft = 2
            lRect.hRight = .ScaleWidth - 1
            ' BackgroundRight
            DrawTheme .hWnd, hDC, "TaskBar", 2, 1, lRect, True
            
            ' ProgList
            lRect.hLeft = 0
            lRect.hRight = .ScaleWidth - 2
            lRect.hBottom = .ScaleHeight - 40
            DrawTheme .hWnd, hDC, m_sClass, 4, 1, lRect, True
            
            m_lPartId = 8
            lRect.hLeft = 0
            lRect.hTop = .ScaleHeight - 40
            lRect.hBottom = PartHeight + lRect.hTop
            lRect.hRight = .ScaleWidth
            ' LogOff
            DrawTheme .hWnd, hDC, m_sClass, 8, 1, lRect, True
            
        End With
        
    End If
    
    ' Draw the caption.
    lRect.hBottom = ScaleHeight
    lRect.hLeft = 8
    lRect.hTop = ScaleHeight - 22
    lRect.hRight = ScaleWidth
    DrawText hDC, m_Caption & "     ", -1, lRect, DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_RIGHT
    
End Sub

Private Sub UserControl_InitProperties()
    
    Set m_CaptionFont = Ambient.Font
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set m_CaptionFont = .ReadProperty("CaptionFont", Ambient.Font)
        m_BackColor = .ReadProperty("BackColor", vbButtonFace)
        m_CaptionColor = .ReadProperty("CaptionColor", vbButtonText)
        m_Caption = .ReadProperty("Caption", "")
    End With

End Sub

Private Sub UserControl_Resize()
    
    ReDraw
    
End Sub

Private Sub UserControl_Show()

    ReDraw
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
      .WriteProperty "BackColor", m_BackColor, vbButtonFace
      .WriteProperty "CaptionColor", m_CaptionColor, vbButtonText
      .WriteProperty "Caption", m_Caption, ""
      .WriteProperty "CaptionFont", m_CaptionFont, Ambient.Font
   End With
    
End Sub
