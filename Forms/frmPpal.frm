VERSION 5.00
Begin VB.Form frmPpal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "TaskScreen 1.0"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10215
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   681
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3990
      Left            =   420
      MouseIcon       =   "frmPpal.frx":1D82
      MousePointer    =   99  'Custom
      ScaleHeight     =   264
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   341
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   975
      Width           =   5145
      Begin TaskScreen.SPicture imgIconMin 
         Height          =   285
         Left            =   2580
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2130
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         Picture         =   "frmPpal.frx":1ED4
      End
      Begin TaskScreen.SPicture imgMinIcon 
         Height          =   540
         Left            =   2220
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1740
         Visible         =   0   'False
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   953
      End
   End
   Begin TaskScreen.SPicture imgIcon 
      Height          =   540
      Left            =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   390
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   953
      Picture         =   "frmPpal.frx":226E
   End
   Begin TaskScreen.ucGradContainer ucGradC2 
      Height          =   375
      Left            =   5655
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5010
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   661
      IconSize        =   0
      CaptionColor    =   4595759
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoResize      =   0   'False
      Transparent     =   -1  'True
      ShowBorders     =   0   'False
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1155
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   -930
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   570
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   -915
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picDraw1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   120
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   -930
      Visible         =   0   'False
      Width           =   435
   End
   Begin TaskScreen.TrayArea TrayArea 
      Left            =   3945
      Top             =   -780
      _ExtentX        =   900
      _ExtentY        =   900
      Icon            =   "frmPpal.frx":4000
      ToolTip         =   "TaskScreen 1.0 by HACKPRO TM"
   End
   Begin VB.Timer tmrKey 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6915
      Top             =   6645
   End
   Begin TaskScreen.MorphBorder MBorder 
      Height          =   600
      Left            =   6495
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3630
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1058
      BorderWidth     =   15
      Color1          =   33023
      Color2          =   4210816
      MiddleOut       =   0   'False
   End
   Begin TaskScreen.MorphListBox MList 
      Height          =   4035
      Left            =   5745
      TabIndex        =   3
      Top             =   975
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   7117
      BackAngle       =   70
      BackColor2      =   8438015
      BackColor1      =   16512
      BorderColor     =   4210816
      BackMiddleOut   =   0   'False
      CurveTopLeft    =   5
      CurveTopRight   =   5
      CurveBottomLeft =   10
      CurveBottomRight=   10
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   0   'False
      SelColor1       =   16512
      SelColor2       =   16512
      SelTextColor    =   12640511
      TrackBarColor1  =   16512
      TrackBarColor2  =   12640511
      ButtonColor1    =   4210816
      ButtonColor2    =   33023
      ThumbColor1     =   4210816
      ThumbColor2     =   16576
      ThumbBorderColor=   33023
      ArrowUpColor    =   12640511
      ArrowDownColor  =   16512
      Theme           =   0
      FocusRectColor  =   12640511
      TrackClickColor1=   4210816
      TrackClickColor2=   8438015
      ItemImageSize   =   32
   End
   Begin TaskScreen.ucGradContainer ucGradC 
      Height          =   375
      Left            =   435
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   661
      HeaderAngle     =   180
      BackAngle       =   120
      IconSize        =   0
      HeaderColor2    =   15390687
      HeaderColor1    =   14004415
      BackColor2      =   16249331
      BackColor1      =   15720934
      BorderColor     =   13740473
      CaptionColor    =   4595759
      Caption         =   "           MorphContainer"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CurveTopLeft    =   5
      CurveTopRight   =   5
      CurveBottomLeft =   5
      CurveBottomRight=   5
      Theme           =   8
      AutoResize      =   0   'False
      AutoResizeOn    =   0
   End
   Begin TaskScreen.StartPanel StartPanel 
      Height          =   4425
      Left            =   420
      Top             =   960
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   7805
      CaptionColor    =   4595759
      Caption         =   "TaskScreen 1.0"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgNoIcon 
      Height          =   480
      Left            =   45
      Picture         =   "frmPpal.frx":5D92
      Top             =   -795
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###########################################################################'
' Title:     TaskScreen XP                                                  '
' Author:    Heriberto Mantilla Santamaría                                  '
' Company:   HACKPRO TM                                                     '
' Created:   29/12/06                                                       '
' Version:   1.0 (09 January 2007)                                          '
'                                                                           '
'             Copyright © 2006 HACKPRO TM. All rights reserved              '
'###########################################################################'
Option Explicit

Private Const RDW_INVALIDATE As Long = &H1

Private Const ABM_GETSTATE      As Long = &H4
Private Const ABM_GETTASKBARPOS As Long = &H5

Private Type APPBARDATA
    cbSize           As Long
    hWnd             As Long
    uCallBackmessage As Long
    uEdge            As Long
    rc               As RECT
    lParam           As Long   ' Message specific.
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
        ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As _
        Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) _
        As Long
Private Declare Function GetClassLong Lib "user32" _
        Alias "GetClassLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) _
        As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) _
        As Long
Private Declare Function PrintWindow Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hDCBlt As Long, _
        ByVal nFlags As Long) As Long
Private Declare Function RedrawWindow Lib "user32" ( _
        ByVal hWnd As Long, _
        lprcUpdate As Any, _
        ByVal hrgnUpdate As Long, _
        ByVal fuRedraw As Long) As Long
Private Declare Function SetClassLong Lib "user32" _
        Alias "SetClassLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nStretchMode As Long) As Long
Private Declare Function SHAppBarMessage Lib "shell32.dll" _
        (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Private Declare Function SwitchToThisWindow Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hWindowState As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" ( _
        ByVal lOleColor As Long, _
        ByVal lHPalette As Long, _
        lColorRef As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd _
        As Long) As Long

Private Const STRETCH_HALFTONE As Long = &H4&

Private Declare Function StretchBlt Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal mDestWidth As Long, _
        ByVal mDestHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal mSrcWidth As Long, _
        ByVal mSrcHeight As Long, _
        ByVal dwRop As Long) As Long

Private sWallpaper As String

Private Enum ApiConstants
    CS_DROPSHADOW = &H20000
    GCL_STYLE = -26
End Enum

#If False Then
    Private CS_DROPSHADOW
    Private GCL_STYLE
#End If

Private LastIndex As Long

Private Type POINTAPI
    x       As Long
    y       As Long
End Type

Private Type Msg
    hWnd    As Long
    Message As Long
    wParam  As Long
    lParam  As Long
    time    As Long
    pt      As POINTAPI
End Type

Private Declare Function RegisterHotKey Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal id As Long, _
        ByVal fsModifiers As Long, _
        ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal id As Long) As Long
Private Declare Function PeekMessage Lib "user32" _
        Alias "PeekMessageA" ( _
        lpMsg As Msg, _
        ByVal hWnd As Long, _
        ByVal wMsgFilterMin As Long, _
        ByVal wMsgFilterMax As Long, _
        ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long

Private Const MOD_ALT   As Long = &H1
Private Const PM_REMOVE As Long = &H1
Private Const WM_HOTKEY As Long = &H312

Private FirstTime As Boolean
Private bKey      As Boolean
Private isHide    As Boolean

Private Function ConvertSystemColor(ByVal theColor As Long) As Long

    '* Convierte un long en un color del sistema.
    OleTranslateColor theColor, 0, ConvertSystemColor

End Function

Private Sub Form_Click()

On Error Resume Next
    MList.SetFocus
On Error GoTo 0

End Sub

Private Sub Form_Load()

  Dim i     As Long
  Dim l     As Long
    
    GetOsSystemVersion
    
    If (StartPanel.GetFaultDraw = True) Then
        picDraw.Move 28, 97, 343, 222
        imgMinIcon.Move 148, 93
        imgIconMin.Move 172, 119
        StartPanel.PutForm frmPpal
    Else
        picDraw.Move 28, 65, 343, 266
        imgMinIcon.Move 148, 116
        imgIconMin.Move 172, 142
    End If
    
    bKey = False
    isHide = True
    TrayArea.Visible = True
    Hide
    ' Call the EnumWindows function.
    ReDim tmpWin(0)
    EnumWindows
    Let SaveVal = tmpWin

    ' Register the Alt+Tab hotkey.
    l = RegisterHotKey(frmPpal.hWnd, &HBFFF&, MOD_ALT, vbKeyTab)

    MBorder.Color1 = vb3DHighlight
    MBorder.Color2 = ShiftColorOXP(vbActiveTitleBar)
    
    With ucGradC
        .BorderColor = vbButtonShadow
        .CaptionColor = vbButtonText
        .HeaderColor1 = MBorder.Color1
        .HeaderColor2 = MBorder.Color2
    End With
    
    StartPanel.CaptionColor = vbButtonText
    ucGradC2.CaptionColor = vbButtonText
    
    BackColor = vbButtonFace
    picIcon.BackColor = vbButtonFace

    'GoTo NoTheme

    With MList
        .RedrawFlag = False
        .Theme = 0
        .ArrowDownColor = MBorder.Color2
        .ArrowUpColor = MBorder.Color1
        .BackColor1 = MBorder.Color2
        .BackColor2 = MBorder.Color1
        .BorderColor = vbWindowFrame
        .ButtonColor1 = MBorder.Color2
        .ButtonColor2 = MBorder.Color1
        .FocusRectColor = MBorder.Color2
        .SelColor1 = MBorder.Color2
        .SelColor2 = MBorder.Color1
        .TextColor = vbButtonText
        .SelTextColor = ShiftColorOXP(.TextColor, &H30)
        .ThumbBorderColor = vbWindowFrame
        .ThumbColor1 = MBorder.Color2
        .ThumbColor2 = MBorder.Color1
        .TrackBarColor1 = MBorder.Color2
        .TrackBarColor2 = MBorder.Color1
        .TrackClickColor1 = MBorder.Color2
        .TrackClickColor2 = MBorder.Color1
        .Enabled = True
    End With

    'NoTheme:

    sWallpaper = GetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper")
    Set picBack.Picture = LoadPicture(sWallpaper)
    StretchPic picBack, picBack.ScaleWidth - 10, picBack.ScaleHeight - 40, picBack, _
        picBack.ScaleWidth, picBack.ScaleHeight

    ' Ulli's VB Recent Code.
    l = SetClassLong(hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW)

    MList.Clear
    MList.RedrawFlag = False

    l = 0

    For i = 0 To UBound(SaveVal) - 1
        MList.AddItem SaveVal(i).TextWindow, i

        'If Not (SaveVal(i).Picture Is Nothing) Then
        '    SavePicture SaveVal(i).Picture, App.Path & "\tmp.bmp"
        '    MList.AddImage App.Path & "\tmp.bmp"
        '    Kill App.Path & "\tmp.bmp"
        '    MList.ImageIndex(MList.NewIndex) = l
        '    l = l + 1
        'End If

    Next i

    MList.RedrawFlag = True

    If (MList.ListCount > 0) Then
        MList.ListIndex = 0
        ucGradC.Caption = SaveVal(MList.ListIndex).TextWindow
        ucGradC2.Caption = MList.ListIndex + 1 & "/" & MList.ListCount
        Set imgMinIcon.Picture = SaveVal(MList.ListIndex).Picture
        Set imgIcon.Picture = SaveVal(MList.ListIndex).Picture
        StartPanel.Caption = SaveVal(MList.ListIndex).Description
    Else
        Set imgIcon.Picture = Nothing
        Set imgMinIcon.Picture = Nothing
        ucGradC.Caption = ""
        ucGradC2.Caption = ""
        StartPanel.Caption = "TaskScreen 1.0"
    End If

    Wait 0.6
    tmrKey.Enabled = True

End Sub

Private Sub Form_Resize()

    If (FirstTime = False) Then
        MBorder.DisplayBorder ' Redraw the form's border.
        FirstTime = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    TrayArea.Visible = False
    bKey = True
    ' Unregister hotkey.
    Call UnregisterHotKey(frmPpal.hWnd, &HBFFF&)
    End

End Sub

Private Sub HideMe()

  Dim lhWnd As Long

On Error GoTo HideMeErr

    If (MList.ListIndex >= 0) Then
        Hide
        ReOrder
        Wait 0.02
        bKey = False
        lhWnd = SaveVal(MList.ListIndex).lhWnd
        LastIndex = -1
        ' Put the Focus.
        SwitchToThisWindow lhWnd, vbNormalFocus
        RedrawWindow lhWnd, ByVal 0&, ByVal 0&, RDW_INVALIDATE
        isHide = True
        tmrKey.Enabled = True
    End If

    Exit Sub

HideMeErr:

End Sub

Private Sub MList_Changed(Index As Long)

    If (Index <> MList.ListIndex) Then
        MList.ListIndex = Index
    ElseIf (MList.ListIndex <> LastIndex) Then
        LastIndex = -1
        PaintMe
    End If
    
End Sub

Private Sub MList_Click()

    HideMe

End Sub

Private Sub MList_KeyUp(KeyCode As Integer, Shift As Integer)

On Error Resume Next

    If (KeyCode = vbKeyTab) And (Shift = 4) Then
        LastIndex = -1
        If (MList.ListIndex < (MList.ListCount - 1)) Then
            MList.ListIndex = MList.ListIndex + 1
            PaintMe
        ElseIf (MList.ListIndex = (MList.ListCount - 1)) Then
            MList.ListIndex = 0
            PaintMe
        End If

        bKey = True
    Else
        HideMe
    End If

On Error GoTo 0

End Sub

' Draw the window screenshot.
Private Sub PaintMe()

  Dim mWnd    As Long
  Dim lIndex  As Long
  Dim lResult As Long
  Dim ABD     As APPBARDATA

On Error Resume Next

    lIndex = MList.ListIndex
    mWnd = -1
    If (lIndex >= 0) And (LastIndex <> lIndex) Then
        LastIndex = lIndex
        mWnd = SaveVal(lIndex).lhWnd
        ucGradC.Caption = "           " & SaveVal(lIndex).TextWindow
        ucGradC.Refresh
        StartPanel.Caption = SaveVal(lIndex).Description
        ucGradC2.Caption = lIndex + 1 & "/" & MList.ListCount
        ucGradC2.Refresh
        Set imgIcon.Picture = SaveVal(lIndex).Picture
        Wait 0.01
        If (LastIndex <> lIndex) Then
            Exit Sub
        End If
        
        DeleteDC picDraw1.hDC
        picDraw.Cls
        Set picDraw.Picture = Nothing
        picDraw.AutoRedraw = True
        picDraw1.Cls
        picDraw1.Width = (Screen.Width \ Screen.TwipsPerPixelX)
        picDraw1.Height = (Screen.Height \ Screen.TwipsPerPixelY) + 10
        StretchPic picDraw1, picDraw1.ScaleWidth, picDraw1.ScaleHeight, picBack, _
            picBack.ScaleWidth, picBack.ScaleHeight - 20
        imgMinIcon.Visible = False
        imgIconMin.Visible = False

        If (CBool(IsIconic(mWnd)) = False) Then
            UpdateWindow mWnd
            lResult = PrintWindow(mWnd, picDraw1.hDC, 0)
        Else
            Set imgMinIcon.Picture = SaveVal(lIndex).Picture
            imgMinIcon.Visible = True
            imgIconMin.Visible = True
        End If

        ' Draw active windows.
        If (SaveVal(lIndex).FullSize = False) Then
            StretchPic picDraw, picDraw.ScaleWidth - 1, picDraw.ScaleHeight, picDraw1, picDraw1.ScaleWidth, _
                picDraw1.ScaleHeight + 5
        Else
            StretchPic picDraw, picDraw.ScaleWidth - 1, picDraw.ScaleHeight + 1, picDraw1, picDraw1.ScaleWidth, _
                picDraw1.ScaleHeight
        End If
        
        If (SaveVal(lIndex).FullSize = False) Then
            ' Get the taskbar's position.
            SHAppBarMessage ABM_GETTASKBARPOS, ABD
            ' Get the taskbar's state.
            SHAppBarMessage ABM_GETSTATE, ABD
            DeleteDC picDraw1.hDC
            picDraw1.Cls
        
            ' Draw Taskbar.
            BitBlt picDraw1.hDC, 0, 0, ABD.rc.rRight, ABD.rc.rBottom, _
                GetDC(0), 0, ABD.rc.rTop, vbSrcCopy
            StretchPic picDraw, picDraw.ScaleWidth - 1, picDraw.ScaleHeight + 120, picDraw1, picDraw1.ScaleWidth, _
                picDraw1.ScaleHeight, 0, picDraw.ScaleHeight - 10
            DeleteDC picDraw1.hDC
            picDraw1.Cls
        End If
        
        Set picIcon.Picture = Nothing
        picIcon.Cls
        picDraw.Refresh
    End If

On Error GoTo 0

End Sub

Private Sub picDraw_Click()

    HideMe

End Sub

Private Sub ReOrder()

  Dim i      As Long
  Dim bFound As Boolean
  Dim l      As Long
  Dim DPath  As String

    If (bKey = True) Then
        Exit Sub
    End If
    
    DPath = GetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper")

    If (sWallpaper <> DPath) Then
        sWallpaper = DPath
        Set picBack.Picture = LoadPicture(sWallpaper)
        StretchPic picBack, picBack.ScaleWidth - 10, picBack.ScaleHeight - 40, picBack, _
            picBack.ScaleWidth, picBack.ScaleHeight
    End If

    ReDim xTemp(0)
    ReDim tmpWin(0)
    ' Call the Enumwindows-function.
    EnumWindows

    Let xTemp = tmpWin

    If (UBound(xTemp) <> UBound(SaveVal)) Then

        ReDim SaveVal(0)
        Let SaveVal = xTemp

        With MList
            .RedrawFlag = False
            .Clear
            l = 0

            For i = 0 To UBound(SaveVal) - 1
                .AddItem SaveVal(i).TextWindow, i

                'If Not (SaveVal(i).Picture Is Nothing) Then
                '    SavePicture SaveVal(i).Picture, App.Path & "\tmp.bmp"
                '    .AddImage App.Path & "\tmp.bmp"
                '    Kill App.Path & "\tmp.bmp"
                '    .ImageIndex(MList.NewIndex) = l
                '    l = l + 1
                'End If

            Next i
            .RedrawFlag = True
            .ListIndex = 0
        End With

    Else

        bFound = False

        For i = 0 To UBound(xTemp) - 1

            If (SaveVal(i).TextWindow <> xTemp(i).TextWindow) Then
                bFound = True
                Exit For
            End If

        Next i

        If (bFound = True) Then
            ReDim SaveVal(0)
            Let SaveVal = xTemp

            With MList
                .RedrawFlag = False
                .Clear
                l = 0

                For i = 0 To UBound(SaveVal) - 1
                    .AddItem SaveVal(i).TextWindow, i

                    'If Not (SaveVal(i).Picture Is Nothing) Then
                    '    SavePicture SaveVal(i).Picture, App.Path & "\tmp.bmp"
                    '    .AddImage App.Path & "\tmp.bmp"
                    '    Kill App.Path & "\tmp.bmp"
                    '    .ImageIndex(MList.NewIndex) = l
                    '    l = l + 1
                    'End If

                Next i
                .RedrawFlag = True
                .ListIndex = 0
            End With

        End If

    End If

End Sub

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
  
  Dim cRed   As Long, cBlue  As Long
  Dim Delta  As Long, cGreen As Long

    ' Shift a color.
    theColor = ConvertSystemColor(theColor)
    cBlue = ((theColor \ &H10000) Mod &H100)
    cGreen = ((theColor \ &H100) Mod &H100)
    cRed = (theColor And &HFF)
    Delta = &HFF - Base
    cBlue = Base + cBlue * Delta \ &HFF
    cGreen = Base + cGreen * Delta \ &HFF
    cRed = Base + cRed * Delta \ &HFF
    If (cRed > 255) Then
       cRed = 255
    End If
    
    If (cGreen > 255) Then
        cGreen = 255
    End If
    
    If (cBlue > 255) Then
        cBlue = 255
    End If
    
    ShiftColorOXP = cRed + 256& * cGreen + 65536 * cBlue
    
End Function

Private Sub ShowMe()

On Error GoTo ShowMeErr
    isHide = False
    LastIndex = -1
    tmrKey.Enabled = False
    If (MList.ListCount >= 1) Then
        If (MList.ListCount = 1) Then
            MList.ListIndex = 0
            LastIndex = -1
        Else
            MList.ListIndex = 1
        End If

        PaintMe
        Wait 0.001
        Show
        Refresh
        bKey = True
        MList.SetFocus
    End If

    Exit Sub

ShowMeErr:

End Sub

Private Sub StretchPic(ByRef inDestPic As PictureBox, _
                       ByVal newW As Long, _
                       ByVal newH As Long, _
                       ByRef inSrcPic As PictureBox, _
                       ByVal w As Long, _
                       ByVal H As Long, _
                       Optional ByVal l As Integer = 0, _
                       Optional ByVal t As Integer = 0)

  Dim mOrigTone As Long
  Dim mResult   As Long

On Error GoTo StretchPicErr
    mOrigTone = SetStretchBltMode(inDestPic.hDC, STRETCH_HALFTONE)
    StretchBlt inDestPic.hDC, l, t, newW, newH, inSrcPic.hDC, 0, 0, w, H, vbSrcCopy
    inDestPic.Picture = inDestPic.Image
    mResult = SetStretchBltMode(inDestPic.hDC, mOrigTone)

    Exit Sub

StretchPicErr:

End Sub

Private Sub tmrKey_Timer()

  Dim Message As Msg

On Error GoTo tmrKeyErr

    If (bKey = False) Then
        ' Wait for a message.
        WaitMessage
        ' Check if it's a HOTKEY-message.
        If (PeekMessage(Message, frmPpal.hWnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) > 0) Then
            ShowMe
        ElseIf (isHide = True) Then
            ReOrder
        End If
        ' Let the operating system TaskScreen other events.
        DoEvents
    End If

    Exit Sub

tmrKeyErr:

End Sub

Private Sub Wait(ByVal Segundos As Single)

  Dim ComienzoSeg As Single
  Dim FinSeg      As Single

    ' Wait a certain time.
    ComienzoSeg = Timer
    FinSeg = ComienzoSeg + Segundos

    Do While (FinSeg > Timer)
        DoEvents
        If (ComienzoSeg > Timer) Then
            FinSeg = FinSeg - 24 * 60 * 60
        End If

    Loop

End Sub

Private Sub TrayArea_MouseDown(Button As Integer)

    If (Button = vbLeftButton) Then
        ShowMe
    End If
    
End Sub
