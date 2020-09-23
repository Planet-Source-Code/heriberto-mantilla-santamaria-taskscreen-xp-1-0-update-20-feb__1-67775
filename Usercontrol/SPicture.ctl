VERSION 5.00
Begin VB.UserControl SPicture 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   ControlContainer=   -1  'True
   PropertyPages   =   "SPicture.ctx":0000
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   ToolboxBitmap   =   "SPicture.ctx":0016
   Begin VB.PictureBox picSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   2490
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "SPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CombineRgn Lib "gdi32" ( _
        ByVal hDestRgn As Long, _
        ByVal hSrcRgn1 As Long, _
        ByVal hSrcRgn2 As Long, _
        ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" ( _
        ByVal x1 As Long, _
        ByVal y1 As Long, _
        ByVal x2 As Long, _
        ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal _
        OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) _
        As Long
Private Declare Function SetWindowRgn Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hRgn As Long, _
        ByVal bRedraw As Long) As Long

Private Const RGN_OR = 2

Private glHeight                As Long
Private glWidth                 As Long
Private myPicture               As StdPicture
Private myMouseIcon             As StdPicture
Private myMousePointer          As MousePointerConstants
Private mySizeHeight            As Long
Private mySizeWidth             As Long

Public Event Click()

Private Sub CreateSkin()

  Dim lRegion As Long

    picSkin.Cls
    UserControl.Cls
    Set picSkin.Picture = myPicture
    If (picSkin.Picture = 0) Then
        UserControl.Cls
        Exit Sub
    End If

    With UserControl
        .Picture = .picSkin.Picture
        .ScaleWidth = .picSkin.ScaleWidth
        .ScaleHeight = .picSkin.ScaleHeight
        lRegion = RegionFromBitmap(.picSkin, TranslateColor(vbButtonFace))
        SetWindowRgn .hWnd, lRegion, True
        DeleteObject lRegion
        Set .picSkin = Nothing
        .Refresh
    End With
    
End Sub

Public Property Get MouseIcon() As StdPicture

    Set MouseIcon = myMouseIcon

End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)

    Set myMouseIcon = New_MouseIcon
    Set UserControl.MouseIcon = myMouseIcon
    PropertyChanged "MouseIcon"

End Property

Public Property Get MousePointer() As MousePointerConstants

    MousePointer = myMousePointer

End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)

    myMousePointer = New_MousePointer
    UserControl.MousePointer = myMousePointer
    PropertyChanged "MousePointer"

End Property

Public Property Get Picture() As StdPicture

    Set Picture = myPicture

End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)

    Set myPicture = New_Picture
    CreateSkin
    PropertyChanged "Picture"
    Refresh

End Property

Private Function RegionFromBitmap(ByRef picSource As PictureBox, _
                                  Optional ByVal lBackColor As Long = -1) As Long

  Dim lRgnTmp   As Long
  Dim lStart    As Long
  Dim lSkinRgn  As Long
  Dim lCol      As Long
  Dim lRow      As Long

    lSkinRgn = CreateRectRgn(0, 0, 0, 0)

    With picSource
        glHeight = .ScaleHeight
        glWidth = .ScaleWidth
        If (lBackColor < 1) Then
            lBackColor = GetPixel(.hDC, 0, 0)
        End If

        For lRow = 0 To glHeight - 1
            lCol = 0

            Do While (lCol < glWidth)
                Do While (lCol < glWidth) And (GetPixel(.hDC, lCol, lRow) = lBackColor)
                    lCol = lCol + 1
                Loop

                If (lCol < glWidth) Then
                    lStart = lCol

                    Do While (lCol < glWidth) And (GetPixel(.hDC, lCol, lRow) <> lBackColor)
                        lCol = lCol + 1
                    Loop

                    If (lCol > glWidth) Then
                        lCol = glWidth
                    End If
                    
                    lRgnTmp = CreateRectRgn(lStart, lRow, lCol, lRow + 1)
                    CombineRgn lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR
                    DeleteObject lRgnTmp
                End If

            Loop
            
        Next lRow

    End With
    RegionFromBitmap = lSkinRgn

End Function

Public Property Get SizeHeight() As Long

    SizeHeight = mySizeHeight

End Property

Public Property Let SizeHeight(ByVal vNewValue As Long)

    mySizeHeight = vNewValue
    PropertyChanged "SizeHeight"

End Property

Public Property Get SizeWidth() As Long

    SizeWidth = mySizeWidth

End Property

Public Property Let SizeWidth(ByVal vNewValue As Long)

    mySizeWidth = vNewValue
    PropertyChanged "SizeWidth"

End Property

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

'*************************************************************************
'* converts color long COLORREF for api coloring purposes.               *
'*************************************************************************

   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If

End Function

Private Sub UserControl_InitProperties()

    Set myPicture = Nothing
    mySizeHeight = 0
    mySizeWidth = 0

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    SizeHeight = PropBag.ReadProperty("SizeHeight", mySizeHeight)
    SizeWidth = PropBag.ReadProperty("SizeWidth", mySizeWidth)

End Sub

Private Sub UserControl_Resize()

    If (SizeWidth > 0) And (SizeHeight > 0) Then
        UserControl.Height = SizeHeight
        UserControl.Width = SizeWidth
    End If

End Sub

Private Sub UserControl_Show()

    CreateSkin

End Sub

Private Sub UserControl_Terminate()

    Set picSkin = Nothing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "MouseIcon", myMouseIcon, Nothing
    PropBag.WriteProperty "MousePointer", myMousePointer, 0
    PropBag.WriteProperty "Picture", myPicture, Nothing
    PropBag.WriteProperty "SizeHeight", mySizeHeight, 0
    PropBag.WriteProperty "SizeWidth", mySizeWidth, 0

End Sub

