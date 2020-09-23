VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Paint me!!!"
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TaskScreen.MorphBorder MBorder 
      Height          =   600
      Left            =   1425
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1380
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1058
      BorderWidth     =   15
      Color1          =   33023
      Color2          =   4210816
      MiddleOut       =   0   'False
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   360
      Left            =   3855
      TabIndex        =   0
      Top             =   2475
      Width           =   795
   End
   Begin TaskScreen.ucGradContainer ucGradContainer 
      Height          =   2850
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5027
      IconSize        =   0
      HeaderColor2    =   15390687
      HeaderColor1    =   14004415
      BackColor2      =   16249331
      BackColor1      =   15720934
      BorderColor     =   13740473
      CaptionColor    =   4595759
      Caption         =   "   TaskScreen 1.0 Developed By HACKPRO TM "
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderIcon      =   "frmAbout.frx":1D82
      CurveTopLeft    =   10
      CurveTopRight   =   10
      CurveBottomLeft =   10
      CurveBottomRight=   10
      Theme           =   8
      Begin VB.Label lblCredits 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   150
         TabIndex        =   3
         Top             =   630
         Width           =   4410
      End
   End
End
Attribute VB_Name = "frmAbout"
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

Private FirstTime As Boolean

Private Declare Function GetClassLong Lib "user32" Alias _
        "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) _
        As Long

Private Declare Function SetClassLong Lib "user32" Alias _
        "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

Private Enum ApiConstants
    CS_DROPSHADOW = &H20000
    GCL_STYLE = -26
End Enum
#If False Then
    Private CS_DROPSHADOW, GCL_STYLE
#End If

Private Sub cmdClose_Click()

    frmPpal.tmrKey.Enabled = True
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Refresh
    Unload frmAbout
    Set frmAbout = Nothing

End Sub

Private Sub Form_Activate()

    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub Form_Load()

    BackColor = vbButtonFace
    With ucGradContainer
        .BackColor1 = vb3DHighlight
        .BackColor2 = vbButtonFace
        .BorderColor = vbWindowFrame
        .CaptionColor = vbTitleBarText
        .HeaderColor1 = vbActiveBorder
        .HeaderColor2 = vbInactiveBorder
    End With
    
    lblCredits.Caption = "Credits And Thanks" & vbCrLf & vbCrLf & _
            "  John Underhill (Steppenwolfe)" & vbCrLf & "  Matthew R. Usner" & _
            vbCrLf & "  Luprix" & vbCrLf & "  Andrea Tincani" & vbCrLf & "  Khalid Pervaz" & _
            vbCrLf & "  KPD-Team: http://www.allapi.net/" & vbCrLf & "  Ulli's" & _
            vbCrLf & "  Richard Mewett"
    
    MBorder.Color1 = vbActiveTitleBar
    MBorder.Color2 = vbInactiveTitleBar
    
    ' Ulli's VB Recent Code.
    SetClassLong hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW

End Sub

Private Sub Form_Resize()

    If (FirstTime = False) Then
        Me.Cls
        MBorder.DisplayBorder
        FirstTime = True
    End If

End Sub
