VERSION 5.00
Begin VB.UserControl ucGradContainer 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   235
   ToolboxBitmap   =   "ucGradContainer.ctx":0000
End
Attribute VB_Name = "ucGradContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'*************************************************************************
'* MorphContainer - Owner-drawn gradient container control.              *
'* Matthew R. Usner et. al., April, 2005.                                *
'*************************************************************************
'* You are welcome to use this control in your projects, as long as      *
'* comments containing all people named in the credits remain intact.    *
'* Only lowlife code thieves like Ilia HD (MoMoYa) download code, remove *
'* the comments, and claim they wrote it.  If you are distributing this  *
'* as part of a compiled application and the application has a 'Credits' *
'* section, named credit is appreciated, but not required.               *
'*************************************************************************
'* A completely owner-drawn replacement for VB's dull frame control.     *
'* Features include:                                                     *
'* - Separate gradients for header and container.                        *
'* - Container background can be a gradient or bitmap.                   *
'* - Icon display capability.                                            *
'* - Ability to round each corner to user-specified curvature amounts.   *
'* - 12 different XP-style color themes are incorporated.                *
'* - Container can be collapsed and expanded. by double-clicking header. *
'* - Container and header gradients can be drawn at any angle.           *
'* - Icon can be displayed in the left or right of the header.           *
'* - Container may be rendered transparent.                              *
'*************************************************************************
'* Credits and Thanks:                                                   *
'* Originally based on XP Container Control written by Cameron Groves.   *
'* Initially modified by Jim Jose April, 2005 (API enhancements).        *
'* Many enhancements (expand/collapse), property reordering made by      *
'* Franck Nunes.                                                         *
'* Carles P.V. supplied the gradient routine and the code for the        *
'* ability to round each corner individually (idea by Franck Nunes).     *
'* original gradient draw routine by Carles P.V. at txtCodeID=60580      *
'* Dana Seaman provided the code to make container transparent no matter *
'* what contains the container (form, another MorphContainer, etc.).     *
'*************************************************************************

Option Explicit

Private Const DT_SINGLELINE           As Long = &H20      ' strip cr/lf from string before draw.
Private Const DT_WORD_ELLIPSIS        As Long = &H40000

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'  declares for container transparency.
Private Const SRCCOPY   As Long = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
   ByVal nXDest As Long, _
   ByVal nYDest As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal hSrcDC As Long, _
   ByVal xSrc As Long, _
   ByVal ySrc As Long, _
   ByVal dwRop As Long) As Long

' declares for Carles P.V.'s gradient paint routine.
Private Type BITMAPINFOHEADER
   biSize          As Long
   biWidth         As Long
   biHeight        As Long
   biPlanes        As Integer
   biBitCount      As Integer
   biCompression   As Long
   biSizeImage     As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed       As Long
   biClrImportant  As Long
End Type

' used to define the text drawing area.
Private Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

'  property enums.
Public Enum IconSizeEnum
   [Display Full Size] = 0
   [Size To Header] = 1
End Enum

Public Enum XPThemes
   [XP Blue] = 0
   [XP Dark Blue] = 1
   [XP Dark Green] = 2
   [XP Green] = 3
   [XP Light Blue] = 4
   [XP Light Green] = 5
   [XP Orange] = 6
   [XP Pastel Green] = 7
   [XP Purple] = 8
   [XP Red] = 9
   [XP Silver] = 10
   [XP Yellow] = 11
End Enum

Public Enum CaptionAlignment
   [Left Justify] = 0
   [Right Justify] = 1
   [Center] = 2
End Enum

Public Enum AutoResizeEvent
   [Double Left Button Click] = 0
   [Double Right Button Click] = 1
End Enum

Public Enum IconAlignmentOptions
   [Align Left] = 0
   [Align Right] = 1
End Enum

'  property variables and constants.
Private m_Picture As Picture
Private m_Transparent As Boolean
Private m_IconAlignment     As IconAlignmentOptions ' icon can be displayed in left or right of header.
Private m_Movable           As Boolean              ' allows container to be dragged using header.
Private m_CurveTopLeft      As Long                 ' the curvature of the top left corner.
Private m_CurveTopRight     As Long                 ' the curvature of the top right corner.
Private m_CurveBottomLeft   As Long                 ' the curvature of the bottom left corner.
Private m_CurveBottomRight  As Long                 ' the curvature of the bottom right corner.
Private m_HeaderVisible     As Boolean              ' flag that shows/hides header.
Private m_BackMiddleOut     As Boolean              ' flag for container background middle-out gradient.
Private m_HeaderMiddleOut   As Boolean              ' flag for header middle-out gradient.
Private m_Enabled           As Boolean              ' enabled/disabled flag.
Private m_ShowBorders       As Boolean              ' Show/Hide the control borders.
Private m_HeaderAngle       As Single               ' the angle of the header gradient.
Private m_BackAngle         As Single               ' background gradient display angle
Private m_Iconsize          As IconSizeEnum         ' icon size - full or size to header
Private m_HeaderColor1      As OLE_COLOR            ' the first gradient color of the header.
Private m_HeaderColor2      As OLE_COLOR            ' the second gradient color of the header.
Private m_BackColor1        As OLE_COLOR            ' the first gradient color of the background.
Private m_BackColor2        As OLE_COLOR            ' the second gradient color of the background.
Private m_BorderWidth       As Integer              ' width, in pixels, of border.
Private m_BorderColor       As OLE_COLOR            ' color of border.
Private m_CaptionColor      As OLE_COLOR            ' text color of caption.
Private m_Caption           As String               ' caption text.
Private m_HeaderHeight      As Long                 ' height, in pixels, of the header.
Private m_CaptionFont       As StdFont              ' font used to display header text.
Private m_Alignment         As CaptionAlignment     ' caption alignment (left, center, right).
Private m_Icon              As Picture              ' the icon or bitmap to display in the header.
Private m_Theme             As XPThemes             ' XP-style color schemes.
Private m_Expanded          As Boolean              ' informs user when container is full size or collapsed.
Private m_AutoResize        As Boolean              ' container collapses/expands without any code from user.
Private m_AutoResizeEvent   As AutoResizeEvent      ' container autocolapses on single click or double click.

Private Const m_def_Transparent = False
Private Const m_def_IconAlignment = 0               ' initalize icon to left alignment.
Private Const m_def_Movable = False                 ' initialize container to fixed position.
Private Const m_def_CurveTopLeft = 0                ' initialize top left curvature to 0.
Private Const m_def_CurveTopRight = 0               ' initialize top right curvature to 0.
Private Const m_def_CurveBottomLeft = 0             ' initialize bottom left curvature to 0.
Private Const m_def_CurveBottomRight = 0            ' initialize bottom right curvature to 0.
Private Const m_def_HeaderVisible = True            ' initialize the header to be visible.
Private Const m_def_BackMiddleOut = True            ' initialize to a middle-out background gradient.
Private Const m_def_HeaderMiddleOut = True          ' initialize to a middle-out header gradient.
Private Const m_def_Enabled = 0                     ' initialize to disabled.
Private Const m_def_HeaderAngle = 90                ' initialize to horizontal header gradient.
Private Const m_def_BackAngle = 90                  ' initialize to horizontal background gradient.
Private Const m_def_Iconsize = 1                    ' initialize to 'size to header'
Private Const m_def_HeaderColor2 = &HF7E0D3
Private Const m_def_HeaderColor1 = &HEDC5A7
Private Const m_def_BackColor2 = &HFCF4EF
Private Const m_def_BackColor1 = &HFAE8DC
Private Const m_def_Caption = "MorphContainer"      ' default caption text.
Private Const m_def_BorderWidth = 1                 ' initialize border width to 1 pixel.
Private Const m_def_BorderColor = &HDCC1AD
Private Const m_def_Alignment = 0                   ' initalize text to left justification.
Private Const m_def_CaptionColor = &H7B2D02
Private Const m_def_hHeight = 25                    ' initialize header to 25 pixels in height.
Private Const m_def_Theme = 0
Private Const m_def_Expanded = True
Private Const m_def_AutoResize = True
Private Const m_def_AutoResizeEvent = 1

'  events.
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event Resize()

'  miscellaneous control variables and constants.
Private Const RGN_DIFF       As Long = 4
Private Const DIB_RGB_COLORS As Long = 0
Private Const PI             As Single = 3.14159265358979
Private Const TO_DEG         As Single = 180 / PI
Private Const TO_RAD         As Single = PI / 180
Private Const INT_ROT        As Long = 1000
Private m_hMod               As Long
Private PreviousHeight       As Long    ' for collapsing/expanding container, container original height.
Private MousePosY            As Single  ' stores y coordinate of mouse (for collapse & move).
Private MouseButton          As Integer ' stores last mouse button clicked.
Private MouseButtonDown      As Boolean ' for dragging of container.

Private x               As Long, y As Long, h1 As Long, h2 As Long, h3 As Long
Private wid             As Long, hgt As Long
'Property Variables:


'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Events >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()

'*************************************************************************
'* handles expansion or collapse of container on double-click.           *
'*************************************************************************

   If m_HeaderVisible = False Then
'     no header, don't allow autoresizing.
      RaiseEvent DblClick
      Exit Sub
   End If

   If m_AutoResize And MousePosY < m_HeaderHeight Then
      If (m_AutoResizeEvent = [Double Right Button Click] And MouseButton = vbRightButton) Or _
         (m_AutoResizeEvent = [Double Left Button Click] And MouseButton = vbLeftButton) Then
         If m_Expanded = True Then
            CollapseContainer
         Else
            ExpandContainer
         End If
      End If
   End If

   RaiseEvent DblClick

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   MouseButtonDown = False
   RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Initialize()
   m_hMod = LoadLibrary("shell32.dll") ' Used to prevent crashes on Windows XP
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   MouseButtonDown = True
   MouseButton = Button 'Capture clicked button.
   RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'*************************************************************************
'* saves mouse Y position for collapse/expand, and also allows the cont- *
'* ainer to be dragged if the appropriate conditions are met.            *
'*************************************************************************

   Dim Ret As Long

'  capture mouse vertical position for possible use in collapse/expand.
   MousePosY = y

'  allow the container to be dragged if all conditions are met.
   If MouseButtonDown And HeaderVisible And MousePosY < m_HeaderHeight And m_Movable Then
      If Button = vbLeftButton Then
         ReleaseCapture
         Ret = SendMessage(UserControl.hWnd, &H112, &HF012, 0)
         'If m_Transparent Then RedrawControl
      End If
   End If

   RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Private Sub UserControl_Paint()
   RedrawControl
End Sub

Private Sub UserControl_Show()
   RedrawControl
End Sub

Private Sub UserControl_Resize()

   Dim H As Single

   On Error GoTo ErrHandler

   H = m_HeaderHeight * Screen.TwipsPerPixelY
   If UserControl.Height < H Then
      UserControl.Height = H
   End If
   RedrawControl
   RaiseEvent Resize

ErrHandler:

End Sub

Private Sub UserControl_Terminate()
   FreeLibrary m_hMod ' Used to prevent crashes on Windows XP
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Graphics >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub MakeContainerTransparent()

   '*************************************************************************
   '* displays what's behind the container control, thereby effectively     *
   '* rendering the container transparent.                                  *
   '*************************************************************************

   On Error Resume Next

   Dim pX               As Long, pY As Long, r As Long
   Dim ctl              As Control

   h2 = UserControl.hDC               ' the dc of the gradient container control.

   If UserControl.Parent.hWnd = UserControl.ContainerHwnd Then
'     the container resides on a Form, so use parent hDC.
      UserControl.Parent.AutoRedraw = True ' added by mru
      h1 = UserControl.Parent.hDC
   Else
'     the container resides in another container, so use parent container's hDC.
      For Each ctl In Parent.Controls
'        find container.
         If ctl.hWnd = UserControl.ContainerHwnd Then 'Found our container
            ctl.AutoRedraw = True 'AutoRedraw must be True
            h1 = ctl.hDC 'Get the containers hDC
            Exit For
         End If
      Next
   End If

'  get offsets for BitBlt.
   If UserControl.Extender.Container.ScaleMode = vbTwips Then
      pX = UserControl.Extender.Left \ Screen.TwipsPerPixelX
      pY = UserControl.Extender.Top \ Screen.TwipsPerPixelY
   Else
      pX = UserControl.Extender.Left
      pY = UserControl.Extender.Top
   End If

'  copy background to our usercontrol hDC.
   r = BitBlt(h2, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, h1, pX, pY, vbSrcCopy)

End Sub

Private Sub RedrawControl()
'*************************************************************************
'* master routine for painting of container.                             *
'*************************************************************************

   UserControl.Cls

   If m_Transparent Then
      MakeContainerTransparent
      If m_ShowBorders = True Then
         CreateBorder
      End If
      If m_HeaderVisible Then
         SetHeader
      End If
   Else
      If m_Expanded Then
         SetBackGround
      End If
      If m_HeaderVisible Then
         SetHeader
      End If
      If m_ShowBorders = True Then
         CreateBorder
      End If
   End If
   'UserControl.Refresh
   'DoEvents

End Sub

Private Sub ApplyTheme()

'*************************************************************************
'* XP-style color schemes. By Cameron Groves.                            *
'*************************************************************************

    Select Case m_Theme

       Case [XP Blue]
          HeaderColor2 = &HF7E0D3
          HeaderColor1 = &HEDC5A7
          BackColor2 = &HFCF4EF
          BackColor1 = &HFAE8DC
          BorderColor = &HDCC1AD
          CaptionColor = &H7B2D02

       Case [XP Dark Blue]
          HeaderColor2 = &HECDCD3
          HeaderColor1 = &HDABAA8
          BackColor2 = &HF8F2EF
          BackColor1 = &HF1E5DD
          BorderColor = &HD6B4A0
          CaptionColor = &H4B2A17

       Case [XP Dark Green]
          HeaderColor2 = &HD8E5C8
          HeaderColor1 = &HB1CB92
          BackColor2 = &HF1F5EB
          BackColor1 = &HE1EBD5
          BorderColor = &HAAC688
          CaptionColor = &H213B00

       Case [XP Green]
          HeaderColor2 = &HE0EAE8
          HeaderColor1 = &HC2D6D1
          BackColor2 = &HF4F8F7
          BackColor1 = &HE7EFED
          BorderColor = &HBCD3CD
          CaptionColor = &H324741

       Case [XP Light Blue]
          HeaderColor2 = &HF1E3C8
          HeaderColor1 = &HE4C992
          BackColor2 = &HFAF5EB
          BackColor1 = &HF5EAD5
          BorderColor = &HE2C488
          CaptionColor = &H553900

       Case [XP Light Green]
          HeaderColor2 = &HDAF2E3
          HeaderColor1 = &HB5E5C8
          BackColor2 = &HF1FAF5
          BackColor1 = &HE3F5EA
          BorderColor = &HAEE3C3
          CaptionColor = &H245738

       Case [XP Orange]
          HeaderColor2 = &HD2E2FD
          HeaderColor1 = &HA7C6FA
          BackColor2 = &HEFF5FE
          BackColor1 = &HDDE9FD
          BorderColor = &H9FC0FA
          CaptionColor = &H16366D

       Case [XP Pastel Green]
          HeaderColor2 = &HE3E3D6
          HeaderColor1 = &HC9C9AE
          BackColor2 = &HF5F5F0
          BackColor1 = &HEAEAE0
          BorderColor = &HC4C4A6
          CaptionColor = &H39391D

       Case [XP Purple]
          HeaderColor2 = &HEAD7DF
          HeaderColor1 = &HD5B0BF
          BackColor2 = &HF7F1F3
          BackColor1 = &HEFE1E6
          BorderColor = &HD1A9B9
          CaptionColor = &H46202F

       Case [XP Red]
          HeaderColor2 = &HD6D2FB
          HeaderColor1 = &HAEA6F8
          BackColor2 = &HF0EFFE
          BackColor1 = &HE0DDFC
          BorderColor = &HA79EF7
          CaptionColor = &H1D156A

       Case [XP Silver]
          HeaderColor2 = &HECEAE9
          HeaderColor1 = &HD9D6D3
          BackColor2 = &HF8F7F7
          BackColor1 = &HF1EFEE
          BorderColor = &HD6D2CF
          CaptionColor = &H4A4744

       Case [XP Yellow]
          HeaderColor2 = &HE4FAFC
          HeaderColor1 = &HB9EEF4
          BackColor2 = &HEEFCFD
          BackColor1 = &HDCF7FA
          BorderColor = &H95E1EA
          CaptionColor = &H66D5E1

    End Select

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

'*************************************************************************
'* converts color long COLORREF for api coloring purposes.               *
'*************************************************************************

   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If

End Function

Private Sub SetBackGround()

'*************************************************************************
'* displays the control's background gradient.                           *
'*************************************************************************

'  if there is a visible header, the top of the background gradient is the row of pixels
'  under the header.  Otherwise, it is the top of the control, plus the border width.
   If IsPictureThere(m_Picture) Then
      Set UserControl.Picture = m_Picture
   Else
      If m_HeaderVisible Then
         PaintGradient hDC, 0, m_HeaderHeight, ScaleWidth, ScaleHeight - m_HeaderHeight - m_BorderWidth, _
                       TranslateColor(m_BackColor1), TranslateColor(m_BackColor2), m_BackAngle, m_BackMiddleOut
      Else
         PaintGradient hDC, 0, 0, ScaleWidth, ScaleHeight, _
                       TranslateColor(m_BackColor1), TranslateColor(m_BackColor2), m_BackAngle, m_BackMiddleOut
      End If
   End If

End Sub

Private Function IsPictureThere(ByVal Pic As StdPicture) As Boolean

'*************************************************************************
'* checks for existence of a picture.  Thanks to Roger Gilchrist.        *
'*************************************************************************

   If Not Pic Is Nothing Then
      If Pic.Height <> 0 Then
         IsPictureThere = Pic.Width <> 0
      End If
   End If

End Function

Private Sub CreateBorder()

'*************************************************************************
'* draws the border around the control, using appropriate curvatures.    *
'*************************************************************************

   Dim hRgn1   As Long   ' the outer region of the border.
   Dim hRgn2   As Long   ' the inner region of the border.
   Dim hBrush  As Long   ' the solid-color brush used to paint the combined border regions.

'  create the outer region.
   hRgn1 = pvGetRoundedRgn(0, 0, ScaleWidth, ScaleHeight, _
                           m_CurveTopLeft, m_CurveTopRight, _
                           m_CurveBottomLeft, m_CurveBottomRight)

'  create the inner region.
   hRgn2 = pvGetRoundedRgn(m_BorderWidth, m_BorderWidth, _
                           ScaleWidth - m_BorderWidth, ScaleHeight - m_BorderWidth, _
                           m_CurveTopLeft, m_CurveTopRight, _
                           m_CurveBottomLeft, m_CurveBottomRight)

'  combine the outer and inner regions.
   CombineRgn hRgn2, hRgn1, hRgn2, RGN_DIFF
'  create the brush used to color the combined regions.
   hBrush = CreateSolidBrush(TranslateColor(m_BorderColor))
'  color the combined regions.
   FillRgn hDC, hRgn2, hBrush

'  set the container's visibility region.
   SetWindowRgn hWnd, hRgn1, True

'  delete created objects to restore memory.
   DeleteObject hBrush
   DeleteObject hRgn1
   DeleteObject hRgn2

End Sub

Private Function pvGetRoundedRgn(ByVal x1 As Long, ByVal y1 As Long, _
                                 ByVal x2 As Long, ByVal y2 As Long, _
                                 ByVal TopLeftRadius As Long, _
                                 ByVal TopRightRadius As Long, _
                                 ByVal BottomLeftRadius As Long, _
                                 ByVal BottomRightRadius As Long _
                                 ) As Long

'*************************************************************************
'* allows each corner of the container to have its own curvature.        *
'* Code by the Amazing Carles P.V.  Thanks a million (as usual) Carles.  *
'*************************************************************************

   Dim hRgnMain As Long   ' the original "starting point" region.
   Dim hRgnTmp1 As Long   ' the first region that defines a corner's radius.
   Dim hRgnTmp2 As Long   ' the second region that defines a corner's radius.

'  bounding region.
   hRgnMain = CreateRectRgn(x1, y1, x2, y2)

'  top-left corner.
   hRgnTmp1 = CreateRectRgn(x1, y1, x1 + TopLeftRadius, y1 + TopLeftRadius)
   hRgnTmp2 = CreateEllipticRgn(x1, y1, x1 + 2 * TopLeftRadius, y1 + 2 * TopLeftRadius)
   Call CombineRgn(hRgnTmp1, hRgnTmp1, hRgnTmp2, RGN_DIFF)
   Call CombineRgn(hRgnMain, hRgnMain, hRgnTmp1, RGN_DIFF)
   Call DeleteObject(hRgnTmp1)
   Call DeleteObject(hRgnTmp2)

'  top-right corner.
   hRgnTmp1 = CreateRectRgn(x2, y1, x2 - TopRightRadius, y1 + TopRightRadius)
   hRgnTmp2 = CreateEllipticRgn(x2 + 1, y1, x2 + 1 - 2 * TopRightRadius, y1 + 2 * TopRightRadius)
   Call CombineRgn(hRgnTmp1, hRgnTmp1, hRgnTmp2, RGN_DIFF)
   Call CombineRgn(hRgnMain, hRgnMain, hRgnTmp1, RGN_DIFF)
   Call DeleteObject(hRgnTmp1)
   Call DeleteObject(hRgnTmp2)

'  bottom-left corner.
   hRgnTmp1 = CreateRectRgn(x1, y2, x1 + BottomLeftRadius, y2 - BottomLeftRadius)
   hRgnTmp2 = CreateEllipticRgn(x1, y2 + 1, x1 + 2 * BottomLeftRadius, y2 + 1 - 2 * BottomLeftRadius)
   Call CombineRgn(hRgnTmp1, hRgnTmp1, hRgnTmp2, RGN_DIFF)
   Call CombineRgn(hRgnMain, hRgnMain, hRgnTmp1, RGN_DIFF)
   Call DeleteObject(hRgnTmp1)
   Call DeleteObject(hRgnTmp2)

'  bottom-right corner.
   hRgnTmp1 = CreateRectRgn(x2, y2, x2 - BottomRightRadius, y2 - BottomRightRadius)
   hRgnTmp2 = CreateEllipticRgn(x2 + 1, y2 + 1, x2 + 1 - 2 * BottomRightRadius, y2 + 1 - 2 * BottomRightRadius)
   Call CombineRgn(hRgnTmp1, hRgnTmp1, hRgnTmp2, RGN_DIFF)
   Call CombineRgn(hRgnMain, hRgnMain, hRgnTmp1, RGN_DIFF)
   Call DeleteObject(hRgnTmp1)
   Call DeleteObject(hRgnTmp2)

   pvGetRoundedRgn = hRgnMain

End Function

Private Sub SetHeader()

'*************************************************************************
'* displays the header gradient, header caption, and an icon if used.    *
'*************************************************************************

   Dim Clearance As Long

   If Not m_CaptionFont Is Nothing Then

      If Not m_Transparent Then
'        fill header gradient.
         PaintGradient hDC, 0, 0, ScaleWidth, m_HeaderHeight, _
                       TranslateColor(m_HeaderColor1), TranslateColor(m_HeaderColor2), m_HeaderAngle, m_HeaderMiddleOut
      End If

'     obtain the width of one letter to use as a left/right caption display clearance.
      Clearance = TextWidth("A")

'     draw the caption.
      Dim TextRect As RECT  ' will define the text drawing region.

'     apply the font and text color.
      Set UserControl.Font = m_CaptionFont
      UserControl.ForeColor = TranslateColor(m_CaptionColor)

      With TextRect
'        define the text drawing area rectangle.
         If m_Alignment = vbCenter Then
            .Left = (ScaleWidth - TextWidth(m_Caption)) / 2
         ElseIf m_Alignment = vbLeftJustify Then
            If IsThere(m_Icon) Then
'              provide for a left-hand clearance of one character plus height of header.
               .Left = Clearance + m_HeaderHeight
            Else
'              provide for a left-hand clearance of on character width.
               .Left = Clearance
            End If
         Else
'           provide a right-hand clearance of one character width.
            .Left = (ScaleWidth - TextWidth(m_Caption)) - Clearance
         End If
'        define the rest of the text drawing rectangle.
         .Top = (m_HeaderHeight - TextHeight(m_Caption)) / 2
         .Bottom = .Top + TextHeight(m_Caption)
         .Right = ScaleWidth
      End With

'     draw the caption.
      DrawText hDC, m_Caption, -1, TextRect, DT_SINGLELINE Or DT_WORD_ELLIPSIS
      
'     draw the icon, if one has been specified.
      If IsThere(m_Icon) Then
'        if specified, display the icon in its original size.
         If m_Iconsize = [Display Full Size] Then
            PaintPicture m_Icon, IconX, 2
         Else
'           otherwise, fit it into the confines defined by the header's height.
            PaintPicture m_Icon, IconX, 2, m_HeaderHeight - 2, m_HeaderHeight - 3
         End If
      End If

   End If

End Sub

Private Function IconX() As Long

'*************************************************************************
'* returns the X coordinate of icon based on IconAlignment property.     *
'*************************************************************************

   If m_IconAlignment = [Align Left] Then
      IconX = m_BorderWidth + 3 + m_BorderWidth
   Else
      If m_Iconsize = [Size To Header] Then
         IconX = ScaleWidth - m_HeaderHeight - 3 - m_BorderWidth
      Else
         IconX = ScaleWidth - ScaleX(m_Icon.Width, vbHimetric, vbPixels) - 3 - m_BorderWidth
      End If
   End If

End Function

Private Function IsThere(ByVal Pic As StdPicture) As Boolean

'*************************************************************************
'* checks for existence of a picture by checking dimensions.             *
'*************************************************************************

   If Not Pic Is Nothing Then
      If Pic.Height <> 0 Then
         IsThere = Pic.Width <> 0
      End If
   End If

End Function

Private Sub PaintGradient(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, _
                         ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, _
                         ByVal Angle As Single, ByVal bMOut As Boolean)

'*************************************************************************
'* Carles P.V.'s routine, modified by Matthew R. Usner for middle-out    *
'* gradient capability.  Original submission at PSC, txtCodeID=60580.    *
'*************************************************************************

   Dim uBIH      As BITMAPINFOHEADER
   Dim lBits()   As Long
   Dim lGrad()   As Long, lGrad2() As Long

   Dim lClr      As Long
   Dim R1        As Long, G1 As Long, b1 As Long
   Dim R2        As Long, G2 As Long, b2 As Long
   Dim dR        As Long, dG As Long, dB As Long

   Dim Scan      As Long
   Dim i         As Long, j As Long, k As Long
   Dim jIn       As Long
   Dim iEnd      As Long, jEnd As Long
   Dim Offset    As Long

   Dim lQuad     As Long
   Dim AngleDiag As Single
   Dim AngleComp As Single

   Dim g         As Long
   Dim luSin     As Long, luCos As Long
 
   If (Width > 0 And Height > 0) Then

'     when angle is >= 91 and <= 270, the colors
'     invert in MiddleOut mode.  This corrects that.
      If bMOut And Angle >= 91 And Angle <= 270 Then
         g = Color1
         Color1 = Color2
         Color2 = g
      End If

'     -- Right-hand [+] (ox=0º)
      Angle = -Angle + 90

'     -- Normalize to [0º;360º]
      Angle = Angle Mod 360
      If (Angle < 0) Then
         Angle = 360 + Angle
      End If

'     -- Get quadrant (0 - 3)
      lQuad = Angle \ 90

'     -- Normalize to [0º;90º]
        Angle = Angle Mod 90

'     -- Calc. gradient length ('distance')
      If (lQuad Mod 2 = 0) Then
         AngleDiag = Atn(Width / Height) * TO_DEG
      Else
         AngleDiag = Atn(Height / Width) * TO_DEG
      End If
      AngleComp = (90 - Abs(Angle - AngleDiag)) * TO_RAD
      Angle = Angle * TO_RAD
      g = Sqr(Width * Width + Height * Height) * Sin(AngleComp) 'Sinus theorem

'     -- Decompose colors
      If (lQuad > 1) Then
         lClr = Color1
         Color1 = Color2
         Color2 = lClr
      End If
      R1 = (Color1 And &HFF&)
      G1 = (Color1 And &HFF00&) \ 256
      b1 = (Color1 And &HFF0000) \ 65536
      R2 = (Color2 And &HFF&)
      G2 = (Color2 And &HFF00&) \ 256
      b2 = (Color2 And &HFF0000) \ 65536

'     -- Get color distances
      dR = R2 - R1
      dG = G2 - G1
      dB = b2 - b1

'     -- Size gradient-colors array
      ReDim lGrad(0 To g - 1)
      ReDim lGrad2(0 To g - 1)

'     -- Calculate gradient-colors
      iEnd = g - 1
      If (iEnd = 0) Then
'        -- Special case (1-pixel wide gradient)
         lGrad2(0) = (b1 \ 2 + b2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
         For i = 0 To iEnd
            lGrad2(i) = b1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
         Next i
      End If

'     'if block' added by Matthew R. Usner - accounts for possible MiddleOut gradient draw.
      If bMOut Then
         k = 0
         For i = 0 To iEnd Step 2
            lGrad(k) = lGrad2(i)
            k = k + 1
         Next i
         For i = iEnd - 1 To 1 Step -2
            lGrad(k) = lGrad2(i)
            k = k + 1
         Next i
      Else
         For i = 0 To iEnd
            lGrad(i) = lGrad2(i)
         Next i
      End If

'     -- Size DIB array
      ReDim lBits(Width * Height - 1) As Long
      iEnd = Width - 1
      jEnd = Height - 1
      Scan = Width

'     -- Render gradient DIB
      Select Case lQuad

         Case 0, 2
            luSin = Sin(Angle) * INT_ROT
            luCos = Cos(Angle) * INT_ROT
            Offset = 0
            jIn = 0
            For j = 0 To jEnd
               For i = 0 To iEnd
                  lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
               Next i
               jIn = jIn + luCos
               Offset = Offset + Scan
            Next j

         Case 1, 3
            luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
            luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
            Offset = jEnd * Scan
            jIn = 0
            For j = 0 To jEnd
               For i = 0 To iEnd
                  lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
               Next i
               jIn = jIn + luCos
               Offset = Offset - Scan
            Next j

      End Select

'     -- Define DIB header
      With uBIH
         .biSize = 40
         .biPlanes = 1
         .biBitCount = 32
         .biWidth = Width
         .biHeight = Height
      End With

'     -- Paint it!
      Call StretchDIBits(hDC, x, y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)

    End If
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Public Methods >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Sub ExpandContainer()
   If Expanded Then
      Exit Sub
   End If
   UserControl.Cls
   m_Expanded = True
   UserControl.Height = PreviousHeight
   DoEvents
End Sub

Public Sub CollapseContainer()
   If Not (Expanded) Then
      Exit Sub
   End If
   PreviousHeight = UserControl.Height
   UserControl.Cls
   UserControl.BackColor = UserControl.Ambient.BackColor
   m_Expanded = False
   UserControl.Height = m_HeaderHeight
   DoEvents
End Sub

Public Sub Refresh()
   UserControl.Refresh
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Properties >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_InitProperties()

'*************************************************************************
'* initialize properties to the default constants.                       *
'*************************************************************************

   Set m_Icon = Nothing
   Set m_CaptionFont = Ambient.Font
   m_HeaderAngle = m_def_HeaderAngle
   m_BackAngle = m_def_BackAngle
   m_HeaderColor2 = m_def_HeaderColor2
   m_HeaderColor1 = m_def_HeaderColor1
   m_BackColor2 = m_def_BackColor2
   m_BackColor1 = m_def_BackColor1
   m_BorderColor = m_def_BorderColor
   m_CaptionColor = m_def_CaptionColor
   m_Caption = m_def_Caption
   m_Alignment = m_def_Alignment
   m_HeaderHeight = m_def_hHeight
   m_Enabled = m_def_Enabled
   m_BorderWidth = m_def_BorderWidth
   m_BackMiddleOut = m_def_BackMiddleOut
   m_HeaderMiddleOut = m_def_HeaderMiddleOut
   m_HeaderVisible = m_def_HeaderVisible
   m_CurveTopLeft = m_def_CurveTopLeft
   m_CurveTopRight = m_def_CurveTopRight
   m_CurveBottomLeft = m_def_CurveBottomLeft
   m_CurveBottomRight = m_def_CurveBottomRight
   m_Theme = m_def_Theme
   m_Expanded = m_def_Expanded
   m_AutoResize = m_def_AutoResize
   m_AutoResizeEvent = m_def_AutoResizeEvent
   m_Movable = m_def_Movable
   m_IconAlignment = m_def_IconAlignment
   m_Transparent = m_def_Transparent
   m_ShowBorders = True
   Set m_Picture = LoadPicture("")

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'*************************************************************************
'* read properties in the property bag.                                  *
'*************************************************************************

    With PropBag
        Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
        Set m_Icon = .ReadProperty("HeaderIcon", Nothing)
        Set m_CaptionFont = .ReadProperty("CaptionFont", Ambient.Font)
        m_Iconsize = .ReadProperty("IconSize", m_def_Iconsize)
        m_HeaderAngle = .ReadProperty("HeaderAngle", m_def_HeaderAngle)
        m_BackAngle = .ReadProperty("BackAngle", m_def_BackAngle)
        m_HeaderColor2 = .ReadProperty("HeaderColor2", m_def_HeaderColor2)
        m_HeaderColor1 = .ReadProperty("HeaderColor1", m_def_HeaderColor1)
        m_BackColor2 = .ReadProperty("BackColor2", m_def_BackColor2)
        m_BackColor1 = .ReadProperty("BackColor1", m_def_BackColor1)
        m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
        m_CaptionColor = .ReadProperty("CaptionColor", m_def_CaptionColor)
        m_Caption = .ReadProperty("Caption", m_def_Caption)
        m_Alignment = .ReadProperty("CaptionAlignment", m_def_Alignment) 'modified by Franck Nunes
        m_HeaderHeight = .ReadProperty("HeaderHeight", m_def_hHeight)
        m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
        m_BorderWidth = .ReadProperty("BorderWidth", m_def_BorderWidth)
        m_BackMiddleOut = .ReadProperty("BackMiddleOut", m_def_BackMiddleOut)
        m_HeaderMiddleOut = .ReadProperty("HeaderMiddleOut", m_def_HeaderMiddleOut)
        m_HeaderVisible = .ReadProperty("HeaderVisible", m_def_HeaderVisible)
        m_CurveTopLeft = .ReadProperty("CurveTopLeft", m_def_CurveTopLeft)
        m_CurveTopRight = .ReadProperty("CurveTopRight", m_def_CurveTopRight)
        m_CurveBottomLeft = .ReadProperty("CurveBottomLeft", m_def_CurveBottomLeft)
        m_CurveBottomRight = .ReadProperty("CurveBottomRight", m_def_CurveBottomRight)
        m_Theme = .ReadProperty("Theme", m_def_Theme)
        m_Expanded = .ReadProperty("Expanded", m_def_Expanded)
        m_AutoResize = .ReadProperty("AutoResize", m_def_AutoResize)
        m_AutoResizeEvent = .ReadProperty("AutoResizeOn", m_def_AutoResizeEvent)
        m_Movable = .ReadProperty("Movable", m_def_Movable)
        m_IconAlignment = .ReadProperty("IconAlignment", m_def_IconAlignment)
        m_Transparent = .ReadProperty("Transparent", m_def_Transparent)
        m_ShowBorders = .ReadProperty("ShowBorders", True)

    End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'*************************************************************************
'* write the properties in the property bag.                             *
'*************************************************************************

   With PropBag
      .WriteProperty "HeaderAngle", m_HeaderAngle, m_def_HeaderAngle
      .WriteProperty "BackAngle", m_BackAngle, m_def_BackAngle
      .WriteProperty "IconSize", m_Iconsize, m_def_Iconsize
      .WriteProperty "HeaderColor2", m_HeaderColor2, m_def_HeaderColor2
      .WriteProperty "HeaderColor1", m_HeaderColor1, m_def_HeaderColor1
      .WriteProperty "BackColor2", m_BackColor2, m_def_BackColor2
      .WriteProperty "BackColor1", m_BackColor1, m_def_BackColor1
      .WriteProperty "BorderColor", m_BorderColor, m_def_BorderColor
      .WriteProperty "CaptionColor", m_CaptionColor, m_def_CaptionColor
      .WriteProperty "Caption", m_Caption, m_def_Caption
      .WriteProperty "CaptionAlignment", m_Alignment, m_def_Alignment
      .WriteProperty "HeaderHeight", m_HeaderHeight, m_def_hHeight
      .WriteProperty "CaptionFont", m_CaptionFont, Ambient.Font
      .WriteProperty "HeaderIcon", m_Icon, Nothing
      .WriteProperty "Enabled", m_Enabled, m_def_Enabled
      .WriteProperty "BorderWidth", m_BorderWidth, m_def_BorderWidth
      .WriteProperty "BackMiddleOut", m_BackMiddleOut, m_def_BackMiddleOut
      .WriteProperty "HeaderMiddleOut", m_HeaderMiddleOut, m_def_HeaderMiddleOut
      .WriteProperty "HeaderVisible", m_HeaderVisible, m_def_HeaderVisible
      .WriteProperty "CurveTopLeft", m_CurveTopLeft, m_def_CurveTopLeft
      .WriteProperty "CurveTopRight", m_CurveTopRight, m_def_CurveTopRight
      .WriteProperty "CurveBottomLeft", m_CurveBottomLeft, m_def_CurveBottomLeft
      .WriteProperty "CurveBottomRight", m_CurveBottomRight, m_def_CurveBottomRight
      .WriteProperty "Theme", m_Theme, m_def_Theme
      .WriteProperty "Expanded", m_Expanded, m_def_Expanded
      .WriteProperty "AutoResize", m_AutoResize, m_def_AutoResize
      .WriteProperty "AutoResizeOn", m_AutoResizeEvent, m_def_AutoResizeEvent
      .WriteProperty "Movable", m_Movable, m_def_Movable
      .WriteProperty "IconAlignment", m_IconAlignment, m_def_IconAlignment
      .WriteProperty "Transparent", m_Transparent, m_def_Transparent
      .WriteProperty "ShowBorders", m_ShowBorders, True
   End With

   Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
End Sub

Public Property Get AutoResize() As Boolean
    AutoResize = m_AutoResize
End Property

Public Property Let AutoResize(ByVal New_AutoResize As Boolean)
    m_AutoResize = New_AutoResize
    PropertyChanged "AutoResize"
End Property

Public Property Get AutoResizeOn() As AutoResizeEvent
    AutoResizeOn = m_AutoResizeEvent
End Property

Public Property Let AutoResizeOn(ByVal vNewValue As AutoResizeEvent)
   m_AutoResizeEvent = vNewValue
   PropertyChanged "AutoResizeOn"
End Property

Public Property Get BackAngle() As Single
   BackAngle = m_BackAngle
End Property

Public Property Let BackAngle(ByVal New_BackAngle As Single)
'  do some bounds checking.
   If New_BackAngle > 360 Then
      New_BackAngle = 360
   ElseIf New_BackAngle < 0 Then
      New_BackAngle = 0
   End If
   m_BackAngle = New_BackAngle
   PropertyChanged "BackAngle"
   RedrawControl
End Property

Public Property Get BackColor1() As OLE_COLOR
   BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor1 As OLE_COLOR)
   m_BackColor1 = New_BackColor1
   PropertyChanged "BackColor1"
   RedrawControl
End Property

Public Property Get BackColor2() As OLE_COLOR
   BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
   m_BackColor2 = New_BackColor2
   PropertyChanged "BackColor2"
   RedrawControl
End Property

Public Property Get BackMiddleOut() As Boolean
   BackMiddleOut = m_BackMiddleOut
End Property

Public Property Let BackMiddleOut(ByVal New_BackMiddleOut As Boolean)
   m_BackMiddleOut = New_BackMiddleOut
   PropertyChanged "BackMiddleOut"
   RedrawControl
End Property

Public Property Get BorderColor() As OLE_COLOR
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
   m_BorderColor = New_BorderColor
   PropertyChanged "BorderColor"
   RedrawControl
End Property

Public Property Get BorderWidth() As Integer
   BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
   m_BorderWidth = New_BorderWidth
   PropertyChanged "BorderWidth"
   RedrawControl
End Property

Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   RedrawControl
End Property

Public Property Get CaptionAlignment() As CaptionAlignment
   CaptionAlignment = m_Alignment
End Property

Public Property Let CaptionAlignment(ByVal vNewAlignment As CaptionAlignment)
   m_Alignment = vNewAlignment
   PropertyChanged "CaptionAlignment"
   RedrawControl
End Property

Public Property Get CaptionColor() As OLE_COLOR
   CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
   m_CaptionColor = New_CaptionColor
   PropertyChanged "CaptionColor"
   RedrawControl
End Property

Public Property Get CaptionFont() As Font
   Set CaptionFont = m_CaptionFont
End Property

Public Property Set CaptionFont(ByVal vNewCaptionFont As Font)
   Set m_CaptionFont = vNewCaptionFont
   PropertyChanged "CaptionFont"
   RedrawControl
End Property

Public Property Get CurveBottomLeft() As Long
   CurveBottomLeft = m_CurveBottomLeft
End Property

Public Property Let CurveBottomLeft(ByVal New_CurveBottomLeft As Long)
   m_CurveBottomLeft = New_CurveBottomLeft
   PropertyChanged "CurveBottomLeft"
   RedrawControl
End Property

Public Property Get CurveBottomRight() As Long
   CurveBottomRight = m_CurveBottomRight
End Property

Public Property Let CurveBottomRight(ByVal New_CurveBottomRight As Long)
   m_CurveBottomRight = New_CurveBottomRight
   PropertyChanged "CurveBottomRight"
   RedrawControl
End Property

Public Property Get CurveTopLeft() As Long
   CurveTopLeft = m_CurveTopLeft
End Property

Public Property Let CurveTopLeft(ByVal New_CurveTopLeft As Long)
   m_CurveTopLeft = New_CurveTopLeft
   PropertyChanged "CurveTopLeft"
   RedrawControl
End Property

Public Property Get CurveTopRight() As Long
   CurveTopRight = m_CurveTopRight
End Property

Public Property Let CurveTopRight(ByVal New_CurveTopRight As Long)
   m_CurveTopRight = New_CurveTopRight
   PropertyChanged "CurveTopRight"
   RedrawControl
End Property

Public Property Get Enabled() As Boolean
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   m_Enabled = New_Enabled
   PropertyChanged "Enabled"
End Property

Public Property Get Expanded() As Boolean
   Expanded = m_Expanded
End Property

Public Property Get HeaderAngle() As Single
   HeaderAngle = m_HeaderAngle
End Property

Public Property Let HeaderAngle(ByVal New_HeaderAngle As Single)
'  do some bounds checking.
   If New_HeaderAngle > 360 Then
      New_HeaderAngle = 360
   ElseIf New_HeaderAngle < 0 Then
      New_HeaderAngle = 0
   End If
   m_HeaderAngle = New_HeaderAngle
   PropertyChanged "HeaderAngle"
   RedrawControl
End Property

Public Property Get HeaderColor1() As OLE_COLOR
   HeaderColor1 = m_HeaderColor1
End Property

Public Property Let HeaderColor1(ByVal New_HeaderColor1 As OLE_COLOR)
   m_HeaderColor1 = New_HeaderColor1
   PropertyChanged "HeaderColor1"
   RedrawControl
End Property

Public Property Get HeaderColor2() As OLE_COLOR
   HeaderColor2 = m_HeaderColor2
End Property

Public Property Let HeaderColor2(ByVal New_HeaderColor2 As OLE_COLOR)
   m_HeaderColor2 = New_HeaderColor2
   PropertyChanged "HeaderColor2"
   RedrawControl
End Property

Public Property Get HeaderHeight() As Long
   HeaderHeight = m_HeaderHeight
End Property

Public Property Let HeaderHeight(ByVal vNewHeight As Long)
   m_HeaderHeight = vNewHeight
   PropertyChanged "HeaderHeight"
   RedrawControl
End Property

Public Property Get Icon() As Picture
   Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal vNewIcon As Picture)
   Set m_Icon = vNewIcon
   PropertyChanged "HeaderIcon"
   RedrawControl
End Property

Public Property Get HeaderMiddleOut() As Boolean
   HeaderMiddleOut = m_HeaderMiddleOut
End Property

Public Property Let HeaderMiddleOut(ByVal New_HeaderMiddleOut As Boolean)
   m_HeaderMiddleOut = New_HeaderMiddleOut
   PropertyChanged "HeaderMiddleOut"
   RedrawControl
End Property

Public Property Get HeaderVisible() As Boolean
   HeaderVisible = m_HeaderVisible
End Property

Public Property Let HeaderVisible(ByVal New_HeaderVisible As Boolean)
   m_HeaderVisible = New_HeaderVisible
   PropertyChanged "HeaderVisible"
   RedrawControl
End Property

Public Property Get IconSize() As IconSizeEnum
   IconSize = m_Iconsize
End Property

Public Property Let IconSize(ByVal New_IconSize As IconSizeEnum)
   m_Iconsize = New_IconSize
   PropertyChanged "IconSize"
   RedrawControl
End Property

Public Property Get Theme() As XPThemes
   Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As XPThemes)
   m_Theme = New_Theme
   PropertyChanged "Theme"
   ApplyTheme
End Property

Public Property Get Movable() As Boolean
   Movable = m_Movable
End Property

Public Property Let Movable(ByVal New_Movable As Boolean)
   m_Movable = New_Movable
   PropertyChanged "Movable"
End Property

Public Property Get IconAlignment() As IconAlignmentOptions
Attribute IconAlignment.VB_Description = "Allows the icon to be displayed on either the left or right side of the header."
   IconAlignment = m_IconAlignment
End Property

Public Property Let IconAlignment(ByVal New_IconAlignment As IconAlignmentOptions)
   m_IconAlignment = New_IconAlignment
   PropertyChanged "IconAlignment"
   RedrawControl
End Property

Public Property Get ShowBorders() As Boolean
   ShowBorders = m_ShowBorders
End Property

Public Property Let ShowBorders(ByVal New_ShowBorders As Boolean)
   m_ShowBorders = New_ShowBorders
   PropertyChanged "ShowBorders"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get Transparent() As Boolean
Attribute Transparent.VB_Description = "If True, sets the body of the MorphContainer to the graphics of the underlying container (A form, picturebox, another MorphContainer, etc.)"
   Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal New_Transparent As Boolean)
   m_Transparent = New_Transparent
   PropertyChanged "Transparent"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
   hWnd = UserControl.hWnd
End Property

Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
   hDC = UserControl.hDC
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
   Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
   Set m_Picture = New_Picture
   PropertyChanged "Picture"
   RedrawControl
End Property
