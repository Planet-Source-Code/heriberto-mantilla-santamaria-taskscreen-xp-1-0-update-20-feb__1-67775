Attribute VB_Name = "modVars"
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

Public Type RECT
    rLeft   As Long
    rTop    As Long
    rRight  As Long
    rBottom As Long
End Type

Private LasthWnd As Long

Private Declare Function GetClassName Lib "user32" _
        Alias "GetClassNameA" ( _
        ByVal hWnd As Long, _
        ByVal lpClassName As String, _
        ByVal nMaxCount As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () _
        As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As _
        Long, lpRect As RECT) As Long

Private Declare Function FindWindow Lib "user32" _
        Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
Private Declare Function GetParent Lib "user32" _
        (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd _
        As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias _
        "GetWindowLongA" (ByVal hWnd As Long, ByVal wIndx As _
        Long) As Long
Private Declare Function GetWindowText Lib "user32" _
        Alias "GetWindowTextA" ( _
        ByVal hWnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" _
        Alias "GetWindowTextLengthA" ( _
        ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" _
        (ByVal hWnd As Long) As Long
       
Private Const GW_HWNDFIRST     As Long = 0
Private Const GW_HWNDNEXT      As Long = 2
Private Const GW_OWNER         As Long = 4
'Private Const GWL_HINSTANCE    As Long = (-6)
'Private Const GWL_HWNDPARENT   As Long = (-8)
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const GWL_EXSTYLE      As Long = (-20)
Private Const WS_EX_APPWINDOW  As Long = &H40000
'Private Const GWL_STYLE        As Long = (-16)
'Private Const WS_BORDER        As Long = &H800000
'Private Const WS_VISIBLE       As Long = &H10000000

Private Type WINDOWSVALUE
    ClassName    As String
    Description  As String
    FullSize     As Boolean
    lhWnd        As Long
    Picture      As StdPicture
    ProcessId    As Long
    TextWindow   As String
End Type

Public xTemp()   As WINDOWSVALUE
Public SaveVal() As WINDOWSVALUE
Public tmpWin()  As WINDOWSVALUE

Private Const REG_SZ = 1 ' Unicode nul terminated string.
Private Const REG_BINARY = 3 ' Free form binary.

Public Const HKEY_CURRENT_USER = &H80000001

Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" _
        Alias "RegOpenKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" ( _
        ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Any, _
        lpcbData As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hWnd As Long, lpdwProcessId As Long) As Long
        
Public Const HWND_TOPMOST    As Long = -1
Public Const HWND_NOTOPMOST  As Long = -2
Public Const SWP_NOSIZE      As Long = &H1
Public Const SWP_NOMOVE      As Long = &H2
Public Const SWP_NOACTIVATE  As Long = &H10
Public Const SWP_SHOWWINDOW  As Long = &H40

Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
        ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long)

'----------------------------------------------------------------------------------------->
' By Richard Mewett
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=61438&lngWId=1
'----------------------------------------------------------------------------------------->

' Declares for Unicode support.
Private Const VER_PLATFORM_WIN32_NT = 2

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

Public mWindowsNT       As Boolean
'----------------------------------------------------------------------------------------->

' Enum all process in the OS.
Public Sub EnumWindows()

  Dim hWnd As Long

    LasthWnd = -1
    ReDim Process(0)
    ProcessList
    hWnd = GetDesktopWindow
    GetWindowInfo hWnd
    hWnd = GetWindow(frmPpal.hWnd, GW_HWNDFIRST)
    
    Do
      GetWindowInfo hWnd
      hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop Until (hWnd = 0)

End Sub

' Get the operating system version.
Public Sub GetOsSystemVersion()

  Dim OS As OSVERSIONINFO
    
    OS.dwOSVersionInfoSize = Len(OS)
    GetVersionEx OS
    mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Sub

Public Function GetString(hKey As Long, strPath As String, strValue As String)

  Dim Ret As Long

    ' Open the key.
    RegOpenKey hKey, strPath, Ret
    ' Get the key's content.
    GetString = RegQueryStringValue(Ret, strValue)
    ' Close the key.
    RegCloseKey Ret

End Function

' Info about the specific Window handle.
Private Sub GetWindowInfo(ByVal hWnd&)
 
  Dim Result   As Long, Title       As String, FoundMe  As Boolean
  Dim RetVal   As Long, lpClassName As String, lExStyle As Long
  Dim WinWnd   As Long, Task        As Long, i          As Long
  Dim LastTask As Long, Rec         As RECT
  
    Result = GetWindowTextLength(hWnd) + 1
    Title = Space$(Result)
    lpClassName = Space$(256)
    Result = GetWindowText(hWnd, Title, Result)
    Title = Left$(Title, Len(Title) - 1)
    WinWnd = FindWindow(vbNullString, Title)
    RetVal = GetClassName(WinWnd, lpClassName, 256)
    lpClassName = Left$(lpClassName, RetVal)
    
    If (IsWindowVisible(hWnd)) And (Title <> "") Then
        If (GetParent(hWnd) = 0) Then
            Result = (GetWindow(hWnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
                                    
            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And Result) Or _
                ((lExStyle And WS_EX_APPWINDOW) And Not Result) _
            Then
                GetWindowThreadProcessId hWnd, Task
                GetWindowThreadProcessId LasthWnd, LastTask
                FoundMe = False
                For i = 0 To UBound(Process)
                    If (Process(i).ProcessId = Task) Then
                        If (InStr(1, UCase$(Process(i).EXEName), "EXPLORER.EXE", vbTextCompare) > 0) Or (InStr(1, UCase$(Process(i).EXEName), "IEXPLORE.EXE") > 0) Then
                            GetIconFromFile Process(i).EXEName, 1, True
                        Else
                            GetIconFromFile Process(i).EXEName, 0, True
                        End If
                        FoundMe = True
                        Exit For
                        
                    End If
                    
                Next i
                    
                GetWindowRect WinWnd, Rec
                    
                If (Rec.rBottom >= (Screen.Height \ Screen.TwipsPerPixelY)) And (Rec.rRight >= (Screen.Width \ Screen.TwipsPerPixelX)) Then
                    tmpWin(UBound(tmpWin)).FullSize = True
                Else
                    tmpWin(UBound(tmpWin)).FullSize = False
                End If
                    
                tmpWin(UBound(tmpWin)).ClassName = lpClassName
                If (LastTask = Task) Then
                    tmpWin(UBound(tmpWin)).lhWnd = CLng(LasthWnd)
                    LasthWnd = -1
                Else
                    tmpWin(UBound(tmpWin)).lhWnd = CLng(hWnd)
                End If
                
                tmpWin(UBound(tmpWin)).TextWindow = Trim$(Title)
                If (FoundMe = True) Then
                    Set tmpWin(UBound(tmpWin)).Picture = frmPpal.picIcon.Image
                    Set frmPpal.picIcon.Picture = Nothing
                Else
                    Set tmpWin(UBound(tmpWin)).Picture = frmPpal.imgNoIcon.Picture
                End If
                    
                If (Process(i).Description <> "") Then
                    tmpWin(UBound(tmpWin)).Description = Process(i).Description
                Else
                    tmpWin(UBound(tmpWin)).Description = Process(i).EXEName
                End If
                
                tmpWin(UBound(tmpWin)).ProcessId = CLng(Result)
                ReDim Preserve tmpWin(UBound(tmpWin) + 1)
                
            Else
                LasthWnd = hWnd
            End If
            
        End If
            
    End If
    
End Sub

Private Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String

  Dim lResult      As Long
  Dim lValueType   As Long
  Dim strBuf       As String
  Dim lDataBufSize As Long

    ' Retrieve information about the key.
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)

    If (lResult = 0) Then
        If (lValueType = REG_SZ) Then
            ' Create a buffer.
            strBuf = String(lDataBufSize, Chr$(0))
            ' Retrieve the key's content.
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)

            If (lResult = 0) Then
                ' Remove the unnecessary chr$(0)'s.
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If

        ElseIf (lValueType = REG_BINARY) Then
            Dim strData As Integer

            ' Retrieve the key's value.
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)

            If (lResult = 0) Then
                RegQueryStringValue = strData
            End If

        End If

    End If

End Function
