Attribute VB_Name = "mdlForm"
'---------------------------------------------------------------------------------------'
'                                                                                       '
' SIMPLE MACHINE PROTECT                                                                '
' Copyright (C) 2008 Bagus Judistirah                                                   '
'                                                                                       '
' This program is free software; you can redistribute it and/or modify                  '
' it under the terms of the GNU General Public License as published by                  '
' the Free Software Foundation; either version 2 of the License, or                     '
' (at your option) any later version.                                                   '
'                                                                                       '
' This program is distributed in the hope that it will be useful,                       '
' but WITHOUT ANY WARRANTY; without even the implied warranty of                        '
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                         '
' GNU General Public License for more details.                                          '
'                                                                                       '
' You should have received a copy of the GNU General Public License along               '
' with this program; if not, write to the Free Software Foundation, Inc.,               '
' 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.                           '
'                                                                                       '
'---------------------------------------------------------------------------------------'
'                                                                                       '
' Module     : Simple Machine Protect - Portable Edition                                '
' Author     : Bagus Judistirah (bagus_badboy)                                          '
' E-mail     : bagus.judistirah@hotmail.com or bagus_badboy@users.sourceforge.net       '
' Homepage   : http://wwww.e-freshware.com                                              '
'            : http://smp.e-freshware.com                                               '
' Project    : http://sourceforge.net/projects/smpav/                                   '
' License    : GNU General Public License                                               '
' History    : Minor bugs fixed.                                                        '
'                                                                                       '
'---------------------------------------------------------------------------------------'
'                                                                                       '
' Note       : I try to keep my software as bug-free as possible.                       '
'              But it's a general rule that no software ever is error free,             '
'              and the number of errors increases with the complexity of the program.   '
'                                                                                       '
'---------------------------------------------------------------------------------------'
'                                                                                       '
' Control    : Simple Machine Protect has been written and developed using Microsoft    '
'              Visual Basic 6. Portions of the source code of this program have been    '
'              taken from or inspired by the source of the following products. Please   '
'              do not remove these copyright notices. The following code & control was  '
'              used during the development of Simple Machine Protect.                   '
'              * Calculate CRC32 Checksum Precompiled Assembler Code, Get Icon          '
'                Coded by: Noel A Dacara                                                '
'                Downloaded from: http://www.planetsourcecode.com                       '
'              * XP Theme                                                               '
'                Coded by: Steve McMahon                                                '
'                Downloaded from: http://www.vbaccelerator.com                          '
'              * Chameleon Button                                                       '
'                Coded by: Gonchuki                                                     '
'                Downloaded from: http://www.planetsourcecode.com                       '
'              * Cool XP ProgressBar                                                    '
'                Coded by: Mario Flores                                                 '
'                Downloaded from: http://www.planetsourcecode.com                       '
'              * OnSystray                                                              '
'                Coded by: Bagus Judistirah                                             '
'                                                                                       '
'---------------------------------------------------------------------------------------'
'                                                                                       '
' Disclaimer : Modifying the registry can cause serious problems that may require you   '
'              to reinstall your operating system. I cannot guarantee that problems     '
'              resulting from modifications to the registry can be solved.              '
'              Use the information provided at your own risk.                           '
'                                                                                       '
'---------------------------------------------------------------------------------------'
' Thanks     : * SOURCEFORGE.NET [http://www.sourceforge.net]                           '
'              * OGNIZER [http://www.ognizer.net or http://virus.ognizer.net]           '
'              * VIROLOGI [http://www.virologi.info]                                    '
'              * ANSAV [http://www.ansav.com]                                           '
'              * VBACCELERATOR [http://www.vbaccelerator.com]                           '
'              * VBBEGO [http://www.vb-bego.com]                                        '
'              * MIGHTHOST [http://www.mighthost.com]                                   '
'              * UDARAMAYA [http://www.udaramaya.com]                                   '
'              * PSC - The home millions of lines of source code.                       '
'                [http://www.planetsourcecode.com]                                      '
'              * DONIXSOFTWARE - Dony Wahyu Isp [http://donixsoftware.web.id]           '
'              * Aat Shadewa, Jan Kristanto, Boby Ertanto, Irwan Halim, Dony Wahyu Isp, '
'                Yusuf Teretsa Patiku, Erwin, MI People, Nita, Husni, I Gede, Fadil,    '
'                Harry, Jimmy Wijaya, Sumanto Adi, Gafur, Selwin, Deny Kurniawan,       '
'                Paul, Marx, Gonchuki, Noel A Dacara, Steve McMahon, Mario Flores,      '
'                VM, Wardana, Achmad Darmal, Andi, Septian, all my friends,             '
'                Dream Theater, Evanescence, & Umild.                                   '
'              * Free software developer around the world.                              '
'              * Thanks to all for the suggestions and comments.                        '
'                                                                                       '
'---------------------------------------------------------------------------------------'
'                                                                                       '
' Contact    : If you have any questions, suggestions, bug reports or anything else,    '
'              feel free to contact me at bagus.judistirah@hotmail.com or               '
'              bagus_badboy@users.sourceforge.net.                                      '
'                                                                                       '
'---------------------------------------------------------------------------------------'

Private Declare Function ReleaseCapture Lib _
    "user32" () As Long
Private Declare Function SendMessage Lib _
    "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Function SetWindowPos Lib _
    "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hwnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib _
    "user32" (ByVal hwnd As Long, _
    ByVal hdcDst As Long, _
    pptDst As Any, _
    psize As Any, _
    ByVal hdcSrc As Long, _
    pptSrc As Any, _
    crKey As Long, _
    ByVal pblend As Long, _
    ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib _
    "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib _
    "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Private Declare Function BeepAPI Lib _
    "kernel32" Alias "Beep" (ByVal dwFreq As Long, _
    ByVal dwDuration As Long) As Long
Private Declare Function GetSaveFileName Lib _
    "COMDLG32.DLL" Alias "GetSaveFileNameA" ( _
    lpofn As OPENFILENAME) As Long
Private Declare Function OpenProcessToken Lib _
    "advapi32" (ByVal ProcessHandle As Long, _
    ByVal DesiredAccess As Long, _
    TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib _
    "advapi32" Alias "LookupPrivilegeValueA" ( _
    ByVal lpSystemName As String, _
    ByVal lpName As String, _
    lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib _
    "advapi32" (ByVal TokenHandle As Long, _
    ByVal DisableAllPrivileges As Long, _
    NewState As TOKEN_PRIVILEGES, _
    ByVal BufferLength As Long, _
    PreviousState As TOKEN_PRIVILEGES, _
    ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib _
    "kernel32" () As Long
Private Declare Function ExitWindowsEx Lib _
    "user32" (ByVal uFlags As Long, _
    ByVal dwReserved As Long) As Long

Private Const EWX_FORCE As Long = 4
Private Const EWX_LOGOFF = 0
Private Const EWX_REBOOT = 2
Private Const EWX_SHUTDOWN = 1

Private Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_EXPLORER = &H80000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_HIDEREADONLY = &H4
Private Const LVM_FIRST = &H1000

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustomFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Const SMP_SITE As String = "smp.e-freshware.com"

Public Sub MoveForm(hwnd As Long)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Public Sub AlwaysOnTop(hwnd As Long, SetOnTop As Boolean)
    If SetOnTop Then
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPFLAGS
    Else
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPFLAGS
    End If
End Sub

Public Function IsTransparent(hwnd As Long) As Boolean
    On Error Resume Next
    Dim Msg As Long
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
        IsTransparent = True
    Else
        IsTransparent = False
    End If
    If Err Then
        IsTransparent = False
    End If
End Function

Public Function MakeTransparent(hwnd As Long, Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 255 Then
        MakeTransparent = 1
    Else
        Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, Msg
        SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
        MakeTransparent = 0
    End If
    If Err Then
        MakeTransparent = 2
    End If
End Function

Public Function MakeOpaque(hwnd As Long) As Long
    Dim Msg As Long
    On Error Resume Next
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    Msg = Msg And Not WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
    MakeOpaque = 0
    If Err Then
        MakeOpaque = 2
    End If
End Function

Public Function CheckValueData(lValue As Long, _
    Optional CheckItemValue As String) As String
    Dim sValueNow As String
    Select Case lValue
        Case Is = 0
            Select Case LCase$(CheckItemValue)
                Case "scanned"
                    sValueNow = "Scanned!"
                Case "infected"
                    sValueNow = "Infected!"
                Case "repaired"
                    sValueNow = "Repaired!"
                Case "detected"
                    sValueNow = "Detected!"
            End Select
            CheckValueData = ": No File " & sValueNow
        Case Is = 1
            CheckValueData = ": " & CStr(lValue) & " File"
        Case Else
            CheckValueData = ": " & CStr(lValue) & " Files"
    End Select
End Function

Public Function CheckBoxesValues(lValue As CheckBox) As String
    If lValue.Value = vbChecked Then
        CheckBoxesValues = ": Enable"
    Else
        CheckBoxesValues = ": Disable"
    End If
End Function

Public Function CheckFileScanValue(lValue As OptionButton, _
    sExtForm As ComboBox) As String
    If lValue.Value = True Then
        CheckFileScanValue = ": All Files"
    Else
        CheckFileScanValue = ": Filtered [" & sExtForm & "]"
    End If
End Function

Public Sub FinishAlert()
    If frmMain.chkSound.Value = 1 Then
        BeepAPI 1800, 50
        Sleep 20
        BeepAPI 1800, 100
    End If
End Sub

Public Sub CreateLogFile(sLocation As String, sInputData As String)
    On Error Resume Next
    Dim lFree As Integer
    lFree = FreeFile
    Open sLocation For Output As #lFree
        Print #lFree, sInputData
    Close #lFree
End Sub

Public Function GetSaveName(Optional WindowTitle As String = "Save Report As", _
    Optional FilterStr As String = "Text Documents (*.log)" + vbNullChar + "*.log") _
    As String
    On Error Resume Next
    Dim DlgInfo As OPENFILENAME
    Dim ret As Long
    Dim FileName As String
    With DlgInfo
        .lStructSize = Len(DlgInfo)
        .hWndOwner = Screen.ActiveForm.hwnd
        .lpstrFilter = FilterStr
        .nFilterIndex = 1
        .lpstrFile = FileName & String(255 - Len(FileName), Chr(0))
        .nMaxFile = 256
        .lpstrFileTitle = String(255, Chr(0))
        .nMaxFileTitle = 256
        .lpstrInitialDir = CurDir & vbNullChar
        .lpstrTitle = WindowTitle & vbNullChar
        .flags = OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or _
            OFN_OVERWRITEPROMPT Or OFN_ENABLEHOOK
        .nMaxCustomFilter = 0
        .nFileOffset = 0
        .nFileExtension = 0
        .lCustData = 0
        .lpfnHook = 0
        .hInstance = 0
    End With
    ret = GetSaveFileName(DlgInfo)
    If Not ret = 0 Then
        GetSaveName = Left(DlgInfo.lpstrFile, InStr(DlgInfo.lpstrFile, vbNullChar) - 1)
    Else
        GetSaveName = ""
    End If
End Function

Public Sub AnimateText(lAnim As Label)
    On Error Resume Next
    With lAnim
        If .Caption = "[-]" Then
            .Caption = "[\]"
        ElseIf .Caption = "[\]" Then
            .Caption = "[|]"
        ElseIf .Caption = "[|]" Then
            .Caption = "[/]"
        ElseIf .Caption = "[/]" Then
            .Caption = "[-]"
        End If
    End With
End Sub

Public Sub LV_AutoSizeColumn(ByVal LV As ListView, _
    Optional ByVal Column As ColumnHeader = Nothing)
    On Error Resume Next
    Dim c As ColumnHeader
    If Column Is Nothing Then
        For Each c In LV.ColumnHeaders
            SendMessage LV.hwnd, LVM_FIRST + 30, c.Index - 1, -1
        Next
    Else
        SendMessage LV.hwnd, LVM_FIRST + 30, Column.Index - 1, -1
    End If
    LV.Refresh
End Sub

Sub ExitNow()
    On Error Resume Next
    frmKarantina.tmrOnTop.Enabled = False
    App.TaskVisible = False
    frmMain.Hide
    frmMain.OnSystray.Visible = False
    frmInfo.Caption = "Closing Application"
    frmInfo.prgInfo.Color = &H4080&
    frmInfo.Show vbModal
    MsgBox "Thanks for using Simple Machine Protect", _
        vbInformation + vbSystemModal, "Thanks"
        Unload frmRTP
        
    End
End Sub

Public Function GenerateMainTitle() As String
    GenerateMainTitle = "$ÏMPLÈ MÅÇHÌNË PRÔTÊÇT"
End Function

Public Function GenerateRandomTitle(sGenNow As Boolean) As String
    Dim sTitle() As Variant
    sTitle = Array("a", "b", "c", "d", "e", "f", "g", _
        "h", "i", "j", "k", "l", "m", "n", "o", "p", _
        "q", "r", "s", "t", "u", "v", "w", "x", "y", _
        "z", "A", "B", "C", "D", "E", "F", "G", "I", _
        "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
        "S", "T", "U", "V", "W", "X", "Y", "Z")
    Randomize
    If sGenNow Then
        GenerateRandomTitle = sTitle(Rnd * UBound(sTitle)) & _
            sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * _
            UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & _
            sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * _
            UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & _
            sTitle(Rnd * UBound(sTitle))
    Else
        GenerateRandomTitle = EncryptText(GenerateRandomTitle)
    End If
End Function

Public Function Decrypt(TextInput As String) As String
    Dim NewLen As Integer
    Dim NewTextInput As String
    Dim NewChar As String
    Dim i As Integer
    NewChar = ""
    NewLen = Len(TextInput)
    For i = 1 To NewLen
        NewChar = Mid(TextInput, i, 1)
        Select Case Asc(NewChar)
            Case 192 To 217
                NewChar = Chr(Asc(NewChar) - 127)
            Case 218 To 243
                NewChar = Chr(Asc(NewChar) - 121)
            Case 244 To 253
                NewChar = Chr(Asc(NewChar) - 196)
            Case 32
                NewChar = Chr(32)
        End Select
        NewTextInput = NewTextInput + NewChar
    Next
    Decrypt = NewTextInput
End Function

Private Function EncryptText(sText As String) _
    As String
    Dim intLen As Integer
    Dim sNewText As String
    Dim sChar As String
    Dim i As Integer
    sChar = ""
    intLen = Len(sText)
    For i = 1 To intLen
        sChar = Mid(sText, i, 1)
        Select Case Asc(sChar)
            Case 65 To 90: sChar = Chr$(Asc(sChar) + 127)
            Case 97 To 122: sChar = Chr$(Asc(sChar) + 121)
            Case 48 To 57: sChar = Chr$(Asc(sChar) + 196)
            Case 32: sChar = Chr$(32)
        End Select
        sNewText = sNewText + sChar
    Next i
    EncryptText = sNewText
End Function

Private Sub AdjustToken()
    Dim lProcHandle As Long
    Dim lTokenHandle As Long
    Dim tmpLUID As LUID
    Dim TKP As TOKEN_PRIVILEGES
    Dim TKPNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    lProcHandle = GetCurrentProcess()
    OpenProcessToken lProcHandle, (&H20 Or &H8), lTokenHandle
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLUID
    With TKP
        .PrivilegeCount = 1
        .TheLuid = tmpLUID
        .Attributes = &H2
    End With
    AdjustTokenPrivileges lTokenHandle, False, TKP, _
        Len(TKPNewButIgnored), TKPNewButIgnored, lBufferNeeded
End Sub

Public Sub ExitWindowsNow(Optional ExitOption As String)
    AdjustToken
    Select Case LCase$(ExitOption)
      Case "logoff"
        ExitWindowsEx (EWX_LOGOFF Or EWX_FORCE), &HFFFF
      Case "reboot"
        ExitWindowsEx (EWX_REBOOT Or EWX_FORCE), &HFFFF
      Case "shutdown"
        ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE), &HFFFF
      'Case "poweroff"
      '  ExitWindowsEx (EWX_POWEROFF Or EWX_FORCE), &HFFFF
    End Select
    End
End Sub
