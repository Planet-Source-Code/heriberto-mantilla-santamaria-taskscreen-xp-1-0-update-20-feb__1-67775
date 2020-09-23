Attribute VB_Name = "modEnumWindows"
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

Private Type VARPROCESS
    Description  As String
    ProcessId    As Long
    EXEName      As String
End Type

Public Process() As VARPROCESS

Private Const PROCESS_QUERY_INFORMATION As Long = 1024
Private Const PROCESS_VM_READ           As Long = &H10

Private Const MAX_PATH  As Long = 260

Private Type PROCESSENTRY32
    dwSize              As Long
    cntUsage            As Long
    th32ProcessID       As Long
    th32DefaultHeapID   As Long
    th32ModuleID        As Long
    cntThreads          As Long
    th32ParentProcessID As Long
    pcPriClassBase      As Long
    dwFlags             As Long
    szExeFile           As String * MAX_PATH
End Type

Private Const TH32CS_SNAPHEAPLIST As Long = &H1
Private Const TH32CS_SNAPMODULE   As Long = &H8
Private Const TH32CS_SNAPPROCESS  As Long = &H2
Private Const TH32CS_SNAPTHREAD   As Long = &H4
Private Const TH32CS_SNAPALL      As Long = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)

Private Declare Function EnumProcessModules Lib "psapi.dll" ( _
        ByVal hProcess As Long, _
        ByRef lphModule As Long, _
        ByVal cb As Long, _
        ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" ( _
        ByVal hProcess As Long, _
        ByVal hModule As Long, _
        ByVal ModuleName As String, _
        ByVal nSize As Long) As Long
'Private Declare Function GetModuleFileName Lib "kernel32" Alias _
        "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName _
        As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
        "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize _
        As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal _
        dwDesiredAccessas As Long, ByVal bInheritHandle As Long, _
        ByVal dwProcId As Long) As Long
        
Private Declare Function CloseHandle Lib "kernel32.dll" ( _
        ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" _
        (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" _
        (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" _
        (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

'  Draw and get Icon.
Private Const DI_MASK   As Long = &H1
Private Const DI_IMAGE  As Long = &H2
Private Const DI_NORMAL As Long = DI_MASK Or DI_IMAGE

Private Declare Function DestroyIcon Lib "user32" _
        (ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
        ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, _
        ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal _
        istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal _
        diFlags As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias _
        "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex _
        As Long, phiconLarge As Long, phiconSmall As Long, ByVal _
        nIcons As Long) As Long

Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
        Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As _
        String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias _
        "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal _
        dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias _
        "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
        lplpBuffer As Any, puLen As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
        (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, _
        ByVal Source As Any, ByVal Length As Long)
        
' John Underhill (Steppenwolfe).
'/* file info structure
Private Type VS_FIXEDFILEINFO
    dwSignature          As Long
    dwStrucVersionl      As Integer
    dwStrucVersionh      As Integer
    dwFileVersionMSl     As Integer
    dwFileVersionMSh     As Integer
    dwFileVersionLSl     As Integer
    dwFileVersionLSh     As Integer
    dwProductVersionMSl  As Integer
    dwProductVersionMSh  As Integer
    dwProductVersionLSl  As Integer
    dwProductVersionLSh  As Integer
    dwFileFlagsMask      As Long
    dwFileFlags          As Long
    dwFileOS             As Long
    dwFileType           As Long
    dwFileSubtype        As Long
    dwFileDateMS         As Long
    dwFileDateLS         As Long
End Type

Private mabytBuffer()    As Byte
Private mffi             As VS_FIXEDFILEINFO

' Code of Hyperion! _
  by John Underhill (Steppenwolfe) _
  CodeId = 63763
Private Function adhTrimNull(strVal As String) As String

  Dim intPos As Integer

    intPos = InStr(1, strVal, vbNullChar)
    Select Case intPos
    Case Is > 1
        adhTrimNull = Left$(strVal, intPos - 1)
    Case 0
        adhTrimNull = strVal
    Case 1
        adhTrimNull = vbNullString
    End Select

End Function

' Based in the code of NT System Security Master Class 1.3 _
  by John Underhill (Steppenwolfe) and Luprix code. _
  CodeId = 65030 _
  CodeId = 54119
Private Function GetFilePathName(ByVal hProcess As Long) As String

  Dim cbNeeded          As Long
  Dim Modules(1 To 200) As Long
  Dim Ret               As Long
  Dim ModuleName        As String
  Dim nSize             As Long
  Dim lProcess          As Long
  
On Error GoTo GetFilePathNameError
    lProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
        Or PROCESS_VM_READ, False, hProcess)

    If (lProcess <> 0) Then
        Ret = EnumProcessModules(lProcess, Modules(1), _
            200, cbNeeded)
        If (Ret <> 0) Then
            ModuleName = Space(MAX_PATH)
            nSize = 500
            Ret = GetModuleFileNameExA(lProcess, _
                Modules(1), ModuleName, nSize)
            GetFilePathName = Left$(ModuleName, Ret)
        End If
        
        If (Mid$(GetFilePathName, 1, Len("\??\")) = "\??\") Then
            GetFilePathName = Replace$(GetFilePathName, "\??\", "")
        ElseIf (Mid$(GetFilePathName, 1, Len("\SystemRoot")) = "\SystemRoot") Then
            GetFilePathName = Replace$(GetFilePathName, "\SystemRoot", GetSystemPath)
        End If
       
        Ret = CloseHandle(lProcess)
       
    End If
    
    If (GetFilePathName = "") Then
        GetFilePathName = "SYSTEM"
    End If
    
Exit Function

GetFilePathNameError:
    
    GetFilePathName = ""

End Function

Public Sub GetIconFromFile(ByVal FileName As String, ByVal IconIndex As Long, ByVal UseLargeIcon As Boolean)

  Dim hLargeIcon As Long
  Dim hSmallIcon As Long
  Dim selHandle  As Long

    frmPpal.picIcon.Cls
    If (ExtractIconEx(FileName, IconIndex, hLargeIcon, hSmallIcon, 1) > 0) Then
        If (UseLargeIcon = True) Then
            selHandle = hLargeIcon
        Else
            selHandle = hSmallIcon
        End If
        
        DrawIconEx frmPpal.picIcon.hDC, 1, 1, selHandle, 32, 32, 0, 0, DI_NORMAL

        DestroyIcon hSmallIcon
        DestroyIcon hLargeIcon
    End If
    
End Sub

' Code of _
  by John Underhill (Steppenwolfe) _
  CodeId =
Private Function GetLanguageInfo(ByVal mstrFileName As String) As String

  Dim lngValue      As Long
  Dim lngBufferLen  As Long
  Dim lngVerPointer As Long
  Dim abytTemp()    As Byte
  Dim mstrLangCP    As String

    If (Len(mstrFileName) = 0) Then
        GetLanguageInfo = ""
        Exit Function
    End If

    lngBufferLen = GetFileVersionInfoSize(mstrFileName, lngValue)
    If (lngBufferLen = 0) Then
        GetLanguageInfo = ""
        Exit Function
    End If

    ReDim mabytBuffer(0 To lngBufferLen - 1)
    If (GetFileVersionInfo(mstrFileName, 0, lngBufferLen, mabytBuffer(0)) = 0) Then
        GetLanguageInfo = ""
        Exit Function
    End If

    If (VerQueryValue(mabytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferLen) = 0) Then
        GetLanguageInfo = ""
        Exit Function
    End If

    ReDim abytTemp(0 To lngBufferLen - 1) As Byte
    RtlMoveMemory abytTemp(0), lngVerPointer, lngBufferLen
    mstrLangCP = ZeroPad(Hex$(abytTemp(1)), 2) & ZeroPad(Hex$(abytTemp(0)), 2) & ZeroPad(Hex$(abytTemp(3)), 2) & ZeroPad(Hex$(abytTemp(2)), 2)

    If (mstrLangCP = "00000000") Then
        mstrLangCP = "040904E4"
    End If
    
    GetLanguageInfo = mstrLangCP

    If (VerQueryValue(mabytBuffer(0), "\", lngVerPointer, lngBufferLen) = 0) Then
        GetLanguageInfo = ""
        Exit Function
    End If
    
    RtlMoveMemory mffi, lngVerPointer, lngBufferLen

End Function

Private Function GetSystemPath()
   
  Dim sSave As String
  Dim Ret   As Long
   
    ' Create a buffer.
    sSave = Space$(255)
    ' Get the system directory.
    Ret = GetSystemDirectory(sSave, 255)
    ' Remove all unnecessary chr$(0)'s.
    sSave = Left$(sSave, Ret)
    
    Ret = InStrRev(sSave, "\")
    
    GetSystemPath = Mid$(sSave, 1, Ret - 1)
   
End Function

' Code of Hyperion! _
  by John Underhill (Steppenwolfe) _
  CodeId = 63763
Private Function GetValue(ByVal strItem As String, ByVal mstrFileName As String) As String

  Dim strTemp       As String
  Dim lngVerPointer As Long
  Dim lngBufferLen  As Long
  Dim strResult     As String

On Error GoTo Handler

    strTemp = "\StringFileInfo\" & GetLanguageInfo(mstrFileName) & "\" & strItem
    If (VerQueryValue(mabytBuffer(0), strTemp, lngVerPointer, lngBufferLen) <> 0) Then
        strResult = Space$(lngBufferLen - 1)
        lstrcpy strResult, lngVerPointer
        strResult = adhTrimNull(strResult)
    End If

ExitHere:
       
    If (Trim$(strResult) = "") Then
        lngBufferLen = InStrRev(mstrFileName, "\", , vbTextCompare) + 1
        GetValue = Mid$(mstrFileName, lngBufferLen, Len(mstrFileName))
    Else
        GetValue = strResult
    End If

Exit Function

Handler:
    lngBufferLen = InStrRev(mstrFileName, "\", , vbTextCompare) + 1
    GetValue = Mid$(mstrFileName, lngBufferLen, Len(mstrFileName))
    Resume ExitHere

End Function

Public Sub ProcessList()
    
  Dim hSnapShot As Long
  Dim uProcess  As PROCESSENTRY32
  Dim r         As Long
  Dim l         As Long
  Dim sFile     As String
        
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
    l = 0
    Do While (r)
        sFile = Left$(uProcess.szExeFile, IIf(InStr(1, _
            uProcess.szExeFile, Chr$(0)) > 0, InStr(1, _
            uProcess.szExeFile, Chr$(0)) - 1, 0))
        
        If (InStrRev(sFile, ".", , vbTextCompare) > 0) Then
            ReDim Preserve Process(l)
            Process(l).ProcessId = CLng(uProcess.th32ProcessID)
            If (mWindowsNT = True) Then
                Process(l).EXEName = GetFilePathName(uProcess.th32ProcessID)
                Process(l).Description = GetValue("FileDescription", Process(l).EXEName)
            Else
                Process(l).Description = GetValue("FileDescription", sFile)
                Process(l).EXEName = sFile
            End If
            l = l + 1
        End If
        
        r = Process32Next(hSnapShot, uProcess)
    Loop
    
    CloseHandle hSnapShot
    
End Sub

' Code of Hyperion! _
  by John Underhill (Steppenwolfe) _
  CodeId = 63763
Private Function ZeroPad(strValue As String, _
                         intLen As String) As String

    ZeroPad = Right$(String$(intLen, "0") & strValue, intLen)

End Function
