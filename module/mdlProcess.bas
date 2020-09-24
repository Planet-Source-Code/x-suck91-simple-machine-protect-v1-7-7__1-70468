Attribute VB_Name = "mdlProcess"
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

Option Explicit

Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or _
    TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const MAX_PATH = 260
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const THREAD_SUSPEND_RESUME = &H2
Private Const REGISTER_SERVICE = 1
Private Const UNREGISTER_SERVICE = 0

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 260
End Type

Private Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(256) As Byte
End Type

Public Type VERHEADER
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    Comments As String
    LegalTradeMarks As String
    PrivateBuild As String
    SpecialBuild As String
End Type

Private Declare Function RegisterServiceProcess Lib _
    "kernel32" (ByVal dwProcessId As Long, _
    ByVal dwType As Long) As Long
Public Declare Function GetCurrentProcessId Lib _
    "kernel32" () As Long
Private Declare Function CreateToolhelp32Snapshot Lib _
    "kernel32" (ByVal lFlags As Long, _
    ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib _
    "kernel32" (ByVal hSnapShot As Long, _
    uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib _
    "kernel32" (ByVal hSnapShot As Long, _
    uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib _
    "kernel32" (ByVal hObject As Long) As Long
Private Declare Function Module32First Lib _
    "kernel32" (ByVal hSnapShot As Long, _
    uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib _
    "kernel32" (ByVal hSnapShot As Long, _
    uProcess As MODULEENTRY32) As Long
Private Declare Function OpenProcess Lib _
    "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib _
    "kernel32" (ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long
Private Declare Function GetPriorityClass Lib _
    "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function SetPriorityClass Lib _
    "kernel32" (ByVal hProcess As Long, _
    ByVal dwPriorityClass As Long) As Long
Private Declare Function OpenThread Lib _
    "kernel32.dll" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Boolean, _
    ByVal dwThreadId As Long) As Long
Private Declare Function ResumeThread Lib _
    "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib _
    "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function Thread32First Lib _
    "kernel32.dll" (ByVal hSnapShot As Long, _
    ByRef lpte As THREADENTRY32) As Boolean
Private Declare Function Thread32Next Lib _
    "kernel32.dll" (ByVal hSnapShot As Long, _
    ByRef lpte As THREADENTRY32) As Boolean
Private Declare Function lstrlen Lib _
    "kernel32" Alias "lstrlenA" ( _
    ByVal lpString As String) As Long
Public Declare Function GetFileAttributes Lib _
    "kernel32" Alias "GetFileAttributesA" ( _
    ByVal lpFileName As String) As Long
Private Declare Function GetFileTitle Lib _
    "comdlg32.dll" Alias "GetFileTitleA" ( _
    ByVal lpszFile As String, _
    ByVal lpszTitle As String, _
    ByVal cbBuf As Integer) As Integer
Private Declare Function OpenFile Lib _
    "kernel32.dll" (ByVal lpFileName As String, _
    ByRef lpReOpenBuff As OFSTRUCT, _
    ByVal wStyle As Long) As Long
Private Declare Function GetFileSize Lib _
    "kernel32" (ByVal hFile As Long, _
    lpFileSizeHigh As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib _
    "psapi.dll" (ByVal Process As Long, _
    ByRef ppsmemCounters As PROCESS_MEMORY_COUNTERS, _
    ByVal cb As Long) As Long
Private Declare Function GetLongPathName Lib _
    "kernel32.dll" Alias "GetLongPathNameA" ( _
    ByVal lpszShortPath As String, _
    ByVal lpszLongPath As String, _
    ByVal cchBuffer As Long) As Long
Private Declare Function GetShortPathNameA Lib _
    "kernel32" (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
Private Declare Function GetFileVersionInfo Lib _
    "Version.dll" Alias "GetFileVersionInfoA" ( _
    ByVal lptstrFilename As String, _
    ByVal dwhandle As Long, _
    ByVal dwlen As Long, _
    lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib _
    "Version.dll" Alias "GetFileVersionInfoSizeA" ( _
    ByVal lptstrFilename As String, _
    lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib _
    "Version.dll" Alias "VerQueryValueA" ( _
    pBlock As Any, _
    ByVal lpSubBlock As String, _
    lplpBuffer As Any, _
    puLen As Long) As Long
Private Declare Sub MoveMemory Lib _
    "kernel32" Alias "RtlMoveMemory" ( _
    dest As Any, _
    ByVal Source As Long, _
    ByVal Length As Long)
Private Declare Function lstrcpy Lib _
    "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, _
    ByVal lpString2 As Long) As Long

Public Enum PriorityClass
   REALTIME_PRIORITY_CLASS = &H100
   HIGH_PRIORITY_CLASS = &H80
   NORMAL_PRIORITY_CLASS = &H20
   IDLE_PRIORITY_CLASS = &H40
End Enum

Dim GetIco As New clsGetIconFile

Function StripNulls(ByVal sStr As String) As String
    StripNulls = Left$(sStr, lstrlen(sStr))
End Function

Public Function NTProcessList(lvwProc As ListView, _
    ilsProc As ImageList) As Long
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Dim Filename As String, ExePath As String
    Dim hProcSnap As Long, hModuleSnap As Long, _
        lProc As Long
    Dim uProcess As PROCESSENTRY32, _
        uModule As MODULEENTRY32
    Dim lvwProcItem As ListItem
    Dim intLVW As Integer
    Dim hVer As VERHEADER
    ExePath = String$(128, Chr$(0))
    hProcSnap = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    lProc = Process32First(hProcSnap, uProcess)
    ilsProc.ListImages.Clear
    lvwProc.ListItems.Clear
    lvwProc.Visible = False
    Do While lProc
        If uProcess.th32ProcessID <> 0 Then
            hModuleSnap = CreateToolhelp32Snapshot(TH32CS_SNAPALL, _
                uProcess.th32ProcessID)
            uModule.dwSize = Len(uModule)
            Module32First hModuleSnap, uModule
            If hModuleSnap > 0 Then
                ExePath = StripNulls(uModule.szExePath)
                Filename = GetFileName(ExePath)
                GetVerHeader ExePath, hVer
                ilsProc.ListImages.Add , "PID" & uProcess.th32ProcessID, _
                    GetIco.Icon(ExePath, SmallIcon)
                Set lvwProcItem = lvwProc.ListItems.Add(, , Filename, , _
                    "PID" & uProcess.th32ProcessID)
                With lvwProcItem
                    .SubItems(1) = GetLongPath(ExePath)
                    .SubItems(2) = Format(GetSizeOfFile(ExePath) / 1024, _
                        "###,###") & " KB"
                    .SubItems(3) = GetAttribute(ExePath)
                    .SubItems(4) = hVer.FileDescription
                    .SubItems(5) = uProcess.th32ProcessID
                    .SubItems(6) = uProcess.cntThreads
                    .SubItems(7) = Format(GetMemory(uProcess.th32ProcessID) / 1024, _
                        "###,####") & " KB"
                    .SubItems(8) = GetBasePriority(uProcess.th32ProcessID)
                End With
            End If
        End If
        lProc = Process32Next(hProcSnap, uProcess)
    Loop
    Call CloseHandle(hProcSnap)
    For intLVW = 1 To lvwProc.ColumnHeaders.Count
        LV_AutoSizeColumn lvwProc, lvwProc.ColumnHeaders.Item(intLVW)
    Next intLVW
    With lvwProc
        With .ColumnHeaders
            .Item(4).Width = 900
            .Item(6).Width = 950
            .Item(7).Width = 800
            .Item(8).Width = 1250
            .Item(9).Width = 800
        End With
        .Refresh
        .Visible = True
        .SetFocus
    End With
    Screen.MousePointer = vbNormal
End Function

Public Function GetBasePriority(ReadPID As Long) As String
    Dim hPID As Long
    hPID = OpenProcess(PROCESS_QUERY_INFORMATION, 0, ReadPID)
    Select Case GetPriorityClass(hPID)
        Case 32: GetBasePriority = "Normal"
        Case 64: GetBasePriority = "Idle"
        Case 128: GetBasePriority = "High"
        Case 256: GetBasePriority = "Realtime"
        Case Else: GetBasePriority = "N/A"
    End Select
    Call CloseHandle(hPID)
End Function

Public Function SetBasePriority(lvwProc As ListView, _
    ItemProcessID As Integer, BasePriority As PriorityClass) As Long
    Dim hPID As Long
    hPID = OpenProcess(PROCESS_ALL_ACCESS, 0, lvwProc.SelectedItem.SubItems( _
        ItemProcessID))
    SetBasePriority = SetPriorityClass(hPID, BasePriority)
    Call CloseHandle(hPID)
End Function

Private Function Thread32Enum(ByRef Thread() As THREADENTRY32, _
    ByVal lProcessID As Long) As Long
    On Error Resume Next
    ReDim Thread(0)
    Dim THREADENTRY32 As THREADENTRY32
    Dim hThreadSnap As Long
    Dim lThread As Long
    hThreadSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lProcessID)
    THREADENTRY32.dwSize = Len(THREADENTRY32)
    If Thread32First(hThreadSnap, THREADENTRY32) = False Then
        Thread32Enum = -1
        Exit Function
    Else
        ReDim Thread(lThread)
        Thread(lThread) = THREADENTRY32
    End If
    Do
        If Thread32Next(hThreadSnap, THREADENTRY32) = False Then
            Exit Do
        Else
            lThread = lThread + 1
            ReDim Preserve Thread(lThread)
            Thread(lThread) = THREADENTRY32
        End If
    Loop
    Thread32Enum = lThread
    Call CloseHandle(hThreadSnap)
End Function

Public Function SetSuspendResumeThread(lvwProc As ListView, _
    ItemProcessID As Integer, SuspendNow As Boolean) As Long
    Dim Thread() As THREADENTRY32, hPID As Long, hThread As Long, i As Long
    hPID = lvwProc.SelectedItem.SubItems(ItemProcessID)
    Thread32Enum Thread(), hPID
    For i = 0 To UBound(Thread)
        If Thread(i).th32OwnerProcessID = hPID Then
            hThread = OpenThread(THREAD_SUSPEND_RESUME, False, (Thread(i).th32ThreadID))
            If SuspendNow Then
                SetSuspendResumeThread = SuspendThread(hThread)
            Else
                SetSuspendResumeThread = ResumeThread(hThread)
            End If
            Call CloseHandle(hThread)
        End If
    Next i
End Function

Public Function TerminateProcessID(lvwProc As ListView, _
    ItemProcessID As Integer) As Long
    Dim hPID As Long
    hPID = OpenProcess(PROCESS_ALL_ACCESS, 0, lvwProc.SelectedItem.SubItems( _
        ItemProcessID))
    TerminateProcessID = TerminateProcess(hPID, 0)
    Call CloseHandle(hPID)
End Function

Public Function GetAttribute(ByVal sFilePath As String) As String
    Select Case GetFileAttributes(sFilePath)
        Case 1: GetAttribute = "R": Case 2: GetAttribute _
            = "H": Case 3: GetAttribute = "RH": Case 4: _
            GetAttribute = "S": Case 5: GetAttribute = _
            "RS": Case 6: GetAttribute = "HS": Case 7: _
            GetAttribute = "RHS"
        Case 32: GetAttribute = "A": Case 33: GetAttribute _
            = "RA": Case 34: GetAttribute = "HA": Case 35: _
            GetAttribute = "RHA": Case 36: GetAttribute = _
            "SA": Case 37: GetAttribute = "RSA": Case 38: _
            GetAttribute = "HSA": Case 39: GetAttribute = _
            "RHSA"
        Case 128: GetAttribute = "Normal"
        Case 2048: GetAttribute = "C": Case 2049: _
            GetAttribute = "RC": Case 2050: GetAttribute = _
            "HC": Case 2051: GetAttribute = "RHC": Case _
            2052: GetAttribute = "SC": Case 2053: _
            GetAttribute = "RSC": Case 2054: GetAttribute _
            = "HSC": Case 2055: GetAttribute = "RHSC": Case _
            2080: GetAttribute = "AC": Case 2081: _
            GetAttribute = "RAC": Case 2082: GetAttribute _
            = "HAC": Case 2083: GetAttribute = "RHAC": Case _
            2084: GetAttribute = "SAC": Case 2085: _
            GetAttribute = "RSAC": Case 2086: GetAttribute _
            = "HSAC": Case 2087: GetAttribute = "RHSAC"
        Case Else: GetAttribute = "N/A"
    End Select
End Function

Public Function GetFileName(ByVal sFileName As String) As String
    Dim buffer As String
    buffer = String(255, 0)
    GetFileTitle sFileName, buffer, Len(buffer)
    buffer = StripNulls(buffer)
    GetFileName = buffer
End Function

Public Function GetSizeOfFile(ByVal PathFile As String) As Long
    Dim hFile As Long, OFS As OFSTRUCT
    hFile = OpenFile(PathFile, OFS, 0)
    GetSizeOfFile = GetFileSize(hFile, 0)
    Call CloseHandle(hFile)
End Function

Public Function GetMemory(ProcessID As Long) As String
    On Error Resume Next
    Dim byteSize As Double, hProcess As Long, ProcMem As PROCESS_MEMORY_COUNTERS
    ProcMem.cb = LenB(ProcMem)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessID)
    If hProcess <= 0 Then GetMemory = "N/A": Exit Function
    GetProcessMemoryInfo hProcess, ProcMem, ProcMem.cb
    byteSize = ProcMem.WorkingSetSize
    GetMemory = byteSize
    Call CloseHandle(hProcess)
End Function

Private Function GetLongPath(ByVal ShortPath As String) As String
    Dim lngRet As Long
    GetLongPath = String$(MAX_PATH, vbNullChar)
    lngRet = GetLongPathName(ShortPath, GetLongPath, Len(GetLongPath))
    If lngRet > Len(GetLongPath) Then
        GetLongPath = String$(lngRet, vbNullChar)
        lngRet = GetLongPathName(ShortPath, GetLongPath, lngRet)
    End If
    If Not lngRet = 0 Then GetLongPath = Left$(GetLongPath, lngRet)
End Function

Public Function GetVerHeader(ByVal fPN$, ByRef oFP As VERHEADER)
    Dim lngBufferlen&, lngDummy&, lngRc&, lngVerPointer&, lngHexNumber&, i%
    Dim bytBuffer() As Byte, bytBuff(255) As Byte, strBuffer$, strLangCharset$, _
        strVersionInfo(11) As String, strTemp$
    If Dir(fPN$, vbHidden + vbArchive + vbNormal + vbReadOnly + vbSystem) = "" Then
        oFP.CompanyName = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.FileDescription = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.FileVersion = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.InternalName = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.LegalCopyright = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.OrigionalFileName = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.ProductName = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.ProductVersion = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.Comments = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.LegalTradeMarks = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.PrivateBuild = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.SpecialBuild = "The file """ & GetShortPath(fPN) & """ N/A"
        Exit Function
    End If
    lngBufferlen = GetFileVersionInfoSize(fPN$, 0)
    If lngBufferlen > 0 Then
        ReDim bytBuffer(lngBufferlen)
        lngRc = GetFileVersionInfo(fPN$, 0&, lngBufferlen, bytBuffer(0))
        If lngRc <> 0 Then
            lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", _
                lngVerPointer, lngBufferlen)
            If lngRc <> 0 Then
                MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
                lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * _
                    &H10000 + bytBuff(1) * &H1000000
                strLangCharset = Hex(lngHexNumber)
                Do While Len(strLangCharset) < 8
                    strLangCharset = "0" & strLangCharset
                Loop
                strVersionInfo(0) = "CompanyName"
                strVersionInfo(1) = "FileDescription"
                strVersionInfo(2) = "FileVersion"
                strVersionInfo(3) = "InternalName"
                strVersionInfo(4) = "LegalCopyright"
                strVersionInfo(5) = "OriginalFileName"
                strVersionInfo(6) = "ProductName"
                strVersionInfo(7) = "ProductVersion"
                strVersionInfo(8) = "Comments"
                strVersionInfo(9) = "LegalTrademarks"
                strVersionInfo(10) = "PrivateBuild"
                strVersionInfo(11) = "SpecialBuild"
                For i = 0 To 11
                    strBuffer = String$(255, 0)
                    strTemp = "\StringFileInfo\" & strLangCharset & "\" & _
                        strVersionInfo(i)
                    lngRc = VerQueryValue(bytBuffer(0), strTemp, lngVerPointer, _
                        lngBufferlen)
                    If lngRc <> 0 Then
                        lstrcpy strBuffer, lngVerPointer
                        strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
                        strVersionInfo(i) = strBuffer
                    Else
                        strVersionInfo(i) = ""
                    End If
                Next i
            End If
        End If
    End If
    For i = 0 To 11
        If Trim(strVersionInfo(i)) = "" Then strVersionInfo(i) = ""
    Next i
    oFP.CompanyName = strVersionInfo(0)
    oFP.FileDescription = strVersionInfo(1)
    oFP.FileVersion = strVersionInfo(2)
    oFP.InternalName = strVersionInfo(3)
    oFP.LegalCopyright = strVersionInfo(4)
    oFP.OrigionalFileName = strVersionInfo(5)
    oFP.ProductName = strVersionInfo(6)
    oFP.ProductVersion = strVersionInfo(7)
    oFP.Comments = strVersionInfo(8)
    oFP.LegalTradeMarks = strVersionInfo(9)
    oFP.PrivateBuild = strVersionInfo(10)
    oFP.SpecialBuild = strVersionInfo(11)
End Function

Private Function GetShortPath(ByVal strFileName As String) As String
    Dim lngRet As Long
    GetShortPath = String$(MAX_PATH, vbNullChar)
    lngRet = GetShortPathNameA(strFileName, GetShortPath, MAX_PATH)
    If Not lngRet = 0 Then GetShortPath = Left$(GetShortPath, lngRet)
End Function

Public Function GetModuleProcessID(lvwProc As ListView, _
    ItemProcID As Integer, lvwModule As ListView, ilsModule As ImageList) As Long
    On Error Resume Next
    Dim ExePath As String
    Dim uProcess As MODULEENTRY32
    Dim hSnapShot As Long
    Dim hPID As Long
    Dim lMod As Long
    Dim intLVW As Integer
    Dim i As Integer
    Dim lvwItem As ListItem
    Dim hVer As VERHEADER
    hPID = lvwProc.SelectedItem.SubItems(ItemProcID)
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, hPID)
    uProcess.dwSize = Len(uProcess)
    lMod = Module32First(hSnapShot, uProcess)
    lvwModule.ListItems.Clear
    ilsModule.ListImages.Clear
    i = 0
    Do While lMod
        i = i + 1
        ExePath = StripNulls(uProcess.szExePath)
        GetVerHeader ExePath, hVer
        ilsModule.ListImages.Add i, , GetIco.Icon(ExePath, SmallIcon)
        Set lvwItem = lvwModule.ListItems.Add(, , ExePath, , i)
        With lvwItem
            .SubItems(1) = hVer.FileDescription
            .SubItems(2) = GetPathType(ExePath)
            .SubItems(3) = hVer.FileVersion
        End With
        lMod = Module32Next(hSnapShot, uProcess)
    Loop
    Call CloseHandle(hSnapShot)
    For intLVW = 1 To lvwModule.ColumnHeaders.Count
        LV_AutoSizeColumn lvwModule, lvwModule.ColumnHeaders.Item(intLVW)
    Next intLVW
End Function

Sub ScanProcess(showMode As Boolean)
    On Error Resume Next
    Dim ExePath As String
    Dim hProcSnap As Long, hModuleSnap As Long, _
        lProc As Long
    Dim uProcess As PROCESSENTRY32, _
        uModule As MODULEENTRY32
    Dim hPID As Long, hExitCode As Long
    ExePath = String$(128, Chr$(0))
    hProcSnap = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    lProc = Process32First(hProcSnap, uProcess)
    Do While lProc
        If uProcess.th32ProcessID <> 0 Then
            hModuleSnap = CreateToolhelp32Snapshot(TH32CS_SNAPALL, _
                uProcess.th32ProcessID)
            uModule.dwSize = Len(uModule)
            Module32First hModuleSnap, uModule
            If hModuleSnap > 0 Then
                DoEvents
                Sleep 10
                ExePath = StripNulls(uModule.szExePath)
                If showMode = True Then
                    frmMain.lblScan.Caption = GetLongPath(ExePath)
                    nMemory = nMemory + 1
                End If
                If IsVirus(ExePath) Then
                    hPID = OpenProcess(1&, -1&, uProcess.th32ProcessID)
                    hExitCode = TerminateProcess(hPID, 0&)
                    Call CloseHandle(hPID)
                End If
            End If
        End If
        lProc = Process32Next(hProcSnap, uProcess)
    Loop
    Call CloseHandle(hProcSnap)
End Sub

Public Function GetAppID() As Long
    GetAppID = GetCurrentProcessId
End Function
