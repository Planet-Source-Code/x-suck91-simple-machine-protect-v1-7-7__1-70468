Attribute VB_Name = "mdlScan"
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

Private Const MaxLen = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const vbStar = "*"
Private Const vbAllFiles = "*.*"
Private Const vbBackslash = "\"
Private Const vbKeyDot = 46

Private Declare Function FindFirstFile Lib _
    "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib _
    "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib _
    "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetFileAttributes Lib _
    "kernel32" Alias "GetFileAttributesA" _
    (ByVal lpFileName As String) As Long
Private Declare Function PathIsDirectory Lib _
    "shlwapi.dll" Alias "PathIsDirectoryA" _
    (ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib _
    "shlwapi.dll" Alias "PathFileExistsA" _
    (ByVal pszPath As String) As Long
Private Declare Function SetFileAttributes Lib _
    "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileName As String, _
    ByVal dwFileAttributes As Long) As Long
Private Declare Function DeleteFile Lib _
    "kernel32" Alias "DeleteFileA" ( _
    ByVal lpFileName As String) As Long
Private Declare Function Beep Lib _
    "kernel32" (ByVal dwFreq As Long, _
    ByVal dwDuration As Long) As Long
    
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MaxLen
    cShortFileName As String * 14
End Type

Dim FileSpec As String, UseFileSpec As Boolean
Dim WFD As WIN32_FIND_DATA, hFindFile As Long
Dim nCalcFiles As Long
Dim PathExt As String
Dim strSavePath As String
Dim nameDetect As String
Public nFile As Long
Public nInfect As Long
Public nRepair As Long
Public nMemory As Long
Public VirLog As String
Public StopScan As Boolean

Dim CRC32 As New clsGetCRC32
Public VirusName As New Collection
Dim VirusSign As New Collection
Dim Heuristic As New clsHeuristic
Dim Addressbar As String
Public Waktu As Timer
Const sDocRepair As String = "ÐÏà¡±á"

'Function GetChecksum: Created by Dony Wahyu Isp, modified by Bagus Judistirah
Function GetChecksum(sFile As String) As String
    On Error GoTo ErrHandle
    Dim cb0 As Byte
    Dim cb1 As Byte
    Dim cb2 As Byte
    Dim cb3 As Byte
    Dim cb4 As Byte
    Dim cb5 As Byte
    Dim cb6 As Byte
    Dim cb7 As Byte
    Dim cb8 As Byte
    Dim cb9 As Byte
    Dim cb10 As Byte
    Dim cb11 As Byte
    Dim cb12 As Byte
    Dim cb13 As Byte
    Dim cb14 As Byte
    Dim cb15 As Byte
    Dim cb16 As Byte
    Dim cb17 As Byte
    Dim cb18 As Byte
    Dim cb19 As Byte
    Dim cb20 As Byte
    Dim cb21 As Byte
    Dim cb22 As Byte
    Dim cb23 As Byte
    Dim buff As String
    Open sFile For Binary Access Read As #1
        buff = Space$(1)
        Get #1, , buff
    Close #1
    'If buff <> "MZ" Then Exit Function
    'If Not (LCase(Right(sFile, 4)) = ".exe") Or _
    '    (LCase(Right(sFile, 4)) = ".scr") Or _
    '    (LCase(Right(sFile, 4)) = ".bat") Or _
    '    (LCase(Right(sFile, 4)) = ".pif") Then
    '    Exit Function
    'End If
    Open sFile For Binary Access Read As #2
        Get #2, 512, cb0
        Get #2, 1024, cb1
        Get #2, 2048, cb2
        Get #2, 3000, cb3
        Get #2, 4096, cb4
        Get #2, 5000, cb5
        Get #2, 6000, cb6
        Get #2, 7000, cb7
        Get #2, 8192, cb8
        Get #2, 9000, cb9
        Get #2, 10000, cb10
        Get #2, 11000, cb11
        Get #2, 12288, cb12
        Get #2, 13000, cb13
        Get #2, 14000, cb14
        Get #2, 15000, cb15
        Get #2, 16384, cb16
        Get #2, 17000, cb17
        Get #2, 18000, cb18
        Get #2, 19000, cb19
        Get #2, 20480, cb20
        Get #2, 21000, cb21
        Get #2, 22000, cb22
        Get #2, 23000, cb23
    Close #2
    buff = cb0
    buff = buff & cb1
    buff = buff & cb2
    buff = buff & cb3
    buff = buff & cb4
    buff = buff & cb5
    buff = buff & cb6
    buff = buff & cb7
    buff = buff & cb8
    buff = buff & cb9
    buff = buff & cb10
    buff = buff & cb11
    buff = buff & cb12
    buff = buff & cb13
    buff = buff & cb14
    buff = buff & cb15
    buff = buff & cb16
    buff = buff & cb17
    buff = buff & cb18
    buff = buff & cb19
    buff = buff & cb20
    buff = buff & cb21
    buff = buff & cb22
    buff = buff & cb23
    GetChecksum = CRC32.StringChecksum(buff)
    Set CRC32 = Nothing
    Exit Function
ErrHandle:
    Close #2
End Function

Private Function FixPath(ByVal path As String) As String
    If Right(path, 1) = vbBackslash Then
        FixPath = path
    Else
        FixPath = path & vbBackslash
    End If
End Function

Private Sub SearchFile(PathSearch As String)
    Dim dirs As Integer, dirbuff() As String, i As Integer
    hFindFile = FindFirstFile(PathSearch & vbAllFiles, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            DoEvents
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    If (dirs Mod 10) = 0 Then ReDim Preserve dirbuff(dirs + 10)
                    dirs = dirs + 1
                    dirbuff(dirs) = StripNulls(WFD.cFileName)
                End If
            ElseIf Not UseFileSpec Then
                nCalcFiles = nCalcFiles + 1
            End If
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    If UseFileSpec Then
        Call SearchFileSpec(PathSearch)
    End If
    For i = 1 To dirs: SearchFile PathSearch & dirbuff(i) & vbBackslash: Next i
End Sub
Private Sub SearchFile1(PathSearch As String)
    Dim dirs As Integer, dirbuff() As String, i As Integer
    hFindFile = FindFirstFile(PathSearch & vbAllFiles, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            DoEvents
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                   
                    dirbuff(dirs) = StripNulls(WFD.cFileName)
               
            ElseIf Not UseFileSpec Then
                nCalcFiles = nCalcFiles + 1
            End If
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    If UseFileSpec Then
        Call SearchFileSpec(PathSearch)
    End If
    For i = 1 To dirs: SearchFile1 PathSearch & dirbuff(i) & vbBackslash: Next i
End Sub

Private Sub SearchFileSpec(PathSearch As String)
    hFindFile = FindFirstFile(PathSearch & FileSpec, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            nCalcFiles = nCalcFiles + 1
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
End Sub

Sub CalcFileNow()
    nFile = 0
    nInfect = 0
    nRepair = 0
    nCalcFiles = 0
    Dim strPath As String
    Dim extNow As String
    strPath = FixPath(frmMain.txtLocation.Text)
    strSavePath = strPath & "...\"
    extNow = Left$(frmMain.cboExt, 5)
    If frmMain.optAllFiles.Value = True Then
        PathExt = vbAllFiles
        SearchFile strPath
    Else
        UseFileSpec = True
        FileSpec = extNow
        PathExt = FileSpec
        SearchFile strPath
        UseFileSpec = False
    End If
End Sub
Sub CalcFDNow()
    nFile = 0
    nInfect = 0
    nRepair = 0
    nCalcFiles = 0
    Dim strPath As String
    Dim extNow As String
    strPath = FixPath(frmScanFD.Text1.Text)
    strSavePath = strPath & "...\"
    extNow = Left$(frmMain.cboExt, 5)
    If frmMain.optAllFiles.Value = True Then
        PathExt = vbAllFiles
        SearchFile strPath
    Else
        UseFileSpec = True
        FileSpec = extNow
        PathExt = FileSpec
        SearchFile strPath
        UseFileSpec = False
    End If
End Sub
Sub Hitung()
nFile = 0
    nInfect = 0
    nRepair = 0
    nCalcFiles = 0
    Dim strPath As String
    Dim extNow As String
    strPath = FixPath(frmRTP.Text1.Text)
    strSavePath = strPath & "...\"
    extNow = Left$(frmMain.cboExt, 5)
    If frmMain.optAllFiles.Value = True Then
        PathExt = vbAllFiles
        SearchFile1 strPath
    Else
        UseFileSpec = True
        FileSpec = extNow
        PathExt = FileSpec
        SearchFile1 strPath
        UseFileSpec = False
    End If
End Sub
Sub RealTimeProtector(path As String, FormMe As Form)
On Error Resume Next
    path = FixPath(path)
    Dim dirs As Integer, dirbuff() As String, i As Integer
    Dim sFilename As String
    hFindFile = FindFirstFile(path & vbStar, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    
                   
                    dirbuff(dirs) = StripNulls(WFD.cFileName)
                End If
            End If
         Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    hFindFile = FindFirstFile(path & PathExt, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
       Do
            If StopScan = True Then Exit Sub
            sFilename = StripNulls(WFD.cFileName)
            If IsFileExist(path & sFilename) = True Then
                
                If Heuristic.CekHeuristic(path & sFilename, frmMain) = True Then
                frmRTP.Label1.Caption = "scan with Heuristic mode"
                frmRTP.lblLokasi.Caption = path & sFilename
                VirusAlert
                frmRTP.tmrGetAddExplo.Enabled = False
                frmRTP.tmrScan.Enabled = False
                frmRTP.Visible = True
                frmRTP.Timer1.Enabled = False
                
                
                End If
                If IsVirus(path & sFilename) Then
                frmRTP.Label1.Caption = nameDetect
                frmRTP.lblLokasi.Caption = path & sFilename
                VirusAlert
                frmRTP.tmrGetAddExplo.Enabled = False
                frmRTP.tmrScan.Enabled = False
                frmRTP.Visible = True
                frmRTP.Timer1.Enabled = False
                End If
            End If
       Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    
        Call RealTimeProtector(path & dirbuff(dirs) & vbBackslash, frmMain)
   
End Sub

Sub ScanVirus(path As String, FormMe As Form)
    On Error Resume Next
    path = FixPath(path)
    Dim dirs As Integer, dirbuff() As String, i As Integer
    Dim sFilename As String
    hFindFile = FindFirstFile(path & vbStar, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    If (dirs Mod 10) = 0 Then ReDim Preserve dirbuff(dirs + 10)
                    dirs = dirs + 1
                    dirbuff(dirs) = StripNulls(WFD.cFileName)
                End If
            End If
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    hFindFile = FindFirstFile(path & PathExt, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            If StopScan = True Then Exit Sub
            sFilename = StripNulls(WFD.cFileName)
            If IsFileExist(path & sFilename) = True Then
                frmMain.lblScan.Caption = strSavePath & sFilename
                If Heuristic.CekHeuristic(path & sFilename, frmMain) = True Then
                    frmMain.lvwVirFound.ListItems.Add(, , "Suspected, Have A Characteristic Virus", , 1).SubItems(1) = path & sFilename
                    VirLog = VirLog & "    Virus Found" & vbTab & vbTab & _
                        vbTab & ": " & (path & sFilename) & vbCrLf
                    If frmMain.chkSound.Value = 1 Then
                        VirusAlert
                         End If
                    End If
                If IsVirus(path & sFilename) Then
                    frmMain.lvwVirFound.ListItems.Add(, , nameDetect, , 1) _
                        .SubItems(1) = path & sFilename
                    VirLog = VirLog & "    Virus Found" & vbTab & vbTab & _
                        vbTab & ": " & (path & sFilename) & vbCrLf
                    If frmMain.chkSound.Value = 1 Then
                        VirusAlert
                    End If
                   
                    If UCase(nameDetect) = "TH.DROP.LOOPS.A.1" Then
                        nInfect = nInfect + 1
                        If frmMain.chkRep.Value = 1 Then
                            RecoverData (path & sFilename), "kspoold"
                            nRepair = nRepair + 1
                            VirLog = VirLog & "    Repaired" & vbTab & vbTab & _
                                vbTab & ": " & (path & sFilename) & vbCrLf
                        End If
                    ElseIf nameDetect = "TH.DROP.VB.DU.1" Then
                        nInfect = nInfect + 1
                        If frmMain.chkRep.Value = 1 Then
                            RecoverData (path & sFilename), "fluburung"
                            nRepair = nRepair + 1
                            VirLog = VirLog & "    Repaired" & vbTab & vbTab & _
                                vbTab & ": " & (path & sFilename) & vbCrLf
                        End If
                    End If
                End If
                nFile = nFile + 1
                frmMain.prgScan.Value = Abs(Round((nFile * 100) / nCalcFiles, 2))
                DoEvents
            End If
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    For i = 1 To dirs
        Call ScanVirus(path & dirbuff(i) & vbBackslash, frmMain)
    Next i
End Sub
Sub ScanFD(path As String, FormMe As Form)
    On Error Resume Next
    path = FixPath(path)
    Dim dirs As Integer, dirbuff() As String, i As Integer
    Dim sFilename As String
    hFindFile = FindFirstFile(path & vbStar, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    If (dirs Mod 10) = 0 Then ReDim Preserve dirbuff(dirs + 10)
                    dirs = dirs + 1
                    dirbuff(dirs) = StripNulls(WFD.cFileName)
                End If
            End If
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    hFindFile = FindFirstFile(path & PathExt, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            If StopScan = True Then Exit Sub
            sFilename = StripNulls(WFD.cFileName)
            If IsFileExist(path & sFilename) = True Then
                frmScanFD.lblScan.Caption = strSavePath & sFilename
                If Heuristic.CekHeuristic(path & sFilename, frmScanFD) = True Then
                    frmScanFD.lvwVirFound.ListItems.Add(, , "Suspected, Have A Characteristic Virus", , 1).SubItems(1) = path & sFilename
                    VirLog = VirLog & "    Virus Found" & vbTab & vbTab & _
                        vbTab & ": " & (path & sFilename) & vbCrLf
                    If frmMain.chkSound.Value = 1 Then
                        VirusAlert
                         End If
                    End If
                If IsVirus(path & sFilename) Then
                    frmScanFD.lvwVirFound.ListItems.Add(, , nameDetect, , 1) _
                        .SubItems(1) = path & sFilename
                    VirLog = VirLog & "    Virus Found" & vbTab & vbTab & _
                        vbTab & ": " & (path & sFilename) & vbCrLf
                    If frmMain.chkSound.Value = 1 Then
                        VirusAlert
                    End If
                   
                    If UCase(nameDetect) = "TH.DROP.LOOPS.A.1" Then
                        nInfect = nInfect + 1
                        If frmMain.chkRep.Value = 1 Then
                            RecoverData (path & sFilename), "kspoold"
                            nRepair = nRepair + 1
                            VirLog = VirLog & "    Repaired" & vbTab & vbTab & _
                                vbTab & ": " & (path & sFilename) & vbCrLf
                        End If
                    ElseIf nameDetect = "TH.DROP.VB.DU.1" Then
                        nInfect = nInfect + 1
                        If frmMain.chkRep.Value = 1 Then
                            RecoverData (path & sFilename), "fluburung"
                            nRepair = nRepair + 1
                            VirLog = VirLog & "    Repaired" & vbTab & vbTab & _
                                vbTab & ": " & (path & sFilename) & vbCrLf
                        End If
                    End If
                End If
                nFile = nFile + 1
                'frmMain.prgScan.Value = Abs(Round((nFile * 100) / nCalcFiles, 2))
                DoEvents
            End If
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    
        Call ScanFD(path & dirbuff(dirs) & vbBackslash, frmScanFD)
  
End Sub
Private Function IsFileExist(path As String) As Boolean
    If PathFileExists(path) = 1 And PathIsDirectory(path) = 0 Then
        IsFileExist = True
    Else
        IsFileExist = False
    End If
End Function

Public Function IsVirus(strFile As String) As Boolean
    On Error Resume Next
    Dim sCrc As String
    sCrc = GetChecksum(strFile)
    Dim i As Integer
    For i = 1 To VirusSign.Count
        If sCrc = VirusSign.Item(i) Then
            IsVirus = True
            nameDetect = VirusName.Item(i)
            Exit Function
        End If
    Next i
End Function

Function IsVirusContain(strFile As String, strCheck As String) As Boolean
    On Error Resume Next
    Dim sFileData As String
    Open strFile For Binary Access Read As #1
        sFileData = Space$(LOF(1))
        Get #1, , sFileData
        If InStr(sFileData, strCheck) > 0 Then
            IsVirusContain = True
            Exit Function
        End If
    Close #1
End Function

Sub LoadVirusDatabase()
    On Error Resume Next
    Dim dbFile As String
    dbFile = App.path & "\SMP.EVD"
    If Not IsFileExist(dbFile) Then
        frmLoading.Hide
        MsgBox "Error opening virus database.", vbCritical + vbSystemModal, _
            "Error Opening Application"
        MsgBox "Error preparing virus list.", vbCritical + vbSystemModal, _
            "Error Opening Application"
        End
    End If
    Dim sData As String
    Open dbFile For Binary Access Read As #1
        sData = String(LOF(1), Chr(0))
        Get #1, , sData
    Close #1
    Dim strArray() As String
    strArray = Split(sData, vbCrLf)
    Dim i As Long
    For i = 1 To UBound(strArray)
        Dim cVirus() As String
        cVirus = Split(strArray(i), ";")
        VirusName.Add cVirus(0)
        VirusSign.Add cVirus(1)
    Next i
End Sub

Public Sub VirusAlert()
    Dim i As Integer
    For i = 1000 To 2000 Step 100
        Beep i, 1
    Next i
End Sub

Private Function MatchFile(FName As String, Mark As String, _
    Optional PosFile As Long = -1) As Boolean
    On Error GoTo ErrHandle
    Dim i As Integer
    Dim hHex() As String
    Dim tmp As String
    hHex() = Split(Mark, " ")
    Dim data() As Byte
    ReDim data(UBound(hHex)) As Byte
    If PosFile > 0 Then
        Open FName For Binary Access Read As #1
           Get #1, PosFile, data
        Close #1
        For i = 0 To UBound(data)
            tmp = tmp & String(2 - Len(Hex(data(i))), "0") & Hex(data(i)) & " "
        Next i
        tmp = IIf(Right(tmp, 1) = " ", Left(tmp, Len(tmp) - 1), tmp)
        If tmp = Mark Then
            MatchFile = True
        End If
    End If
    Exit Function
ErrHandle:
    Close #1
End Function

Private Sub RecoverData(sSourcePath As String, Optional sVirToRepair As String)
    On Error GoTo ErrHandle
    Dim filedata As String
    Dim sResult As String
    Dim lStart As Long
    Dim sNewExt As String
    Open sSourcePath For Binary Access Read As #1
        filedata = Space$(LOF(1))
        Get #1, , filedata
    Close #1
    Select Case LCase$(sVirToRepair)
        Case "kspoold"
            lStart = InStr(filedata, sDocRepair)
            If MatchFile(sSourcePath, "64 6F 63", 330774) = True Then
                sNewExt = ".doc"
            Else
                sNewExt = ".xls"
            End If
        Case "fluburung"
            lStart = InStr(filedata, sDocRepair)
            sNewExt = ".doc"
        'Case "sality"
        '    lStart = InStr(filedata, sSalRepair)
        '    sNewExt = ".exe"
        'case "bacalid"
        '    lStart = InStr(filedata, sBacRepair)
        '    sNewExt = ".exe"
    End Select
    If lStart > 0 Then
        sResult = Mid(filedata, lStart)
        sSourcePath = Replace(sSourcePath, Right$(sSourcePath, 4), sNewExt)
        Open sSourcePath For Binary Access Write As #2
            Put #2, , sResult
        Close #2
    End If
ErrHandle:
End Sub

Sub HiddenRecovery(PathFileHidden As String, AnimateLabel As Label)
    On Error Resume Next
    PathFileHidden = FixPath(PathFileHidden)
    Dim dirs As Integer, dirbuff() As String, i As Integer
    Dim FileName As String
    hFindFile = FindFirstFile(PathFileHidden & vbStar, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            If StopScan = True Then Exit Sub
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    If (dirs Mod 10) = 0 Then ReDim Preserve dirbuff(dirs + 10)
                    dirs = dirs + 1
                    dirbuff(dirs) = StripNulls(WFD.cFileName)
                    End If
            End If
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    hFindFile = FindFirstFile(PathFileHidden & vbAllFiles, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            DoEvents
            If StopScan = True Then Exit Sub
            FileName = StripNulls(WFD.cFileName)
            If GetFileAttributes(PathFileHidden & FileName) = _
                FILE_ATTRIBUTE_DIRECTORY Or FILE_ATTRIBUTE_HIDDEN Or _
                FILE_ATTRIBUTE_SYSTEM Then
                SetFileAttributes PathFileHidden & FileName, _
                FILE_ATTRIBUTE_ARCHIVE + FILE_ATTRIBUTE_NORMAL
            End If
            AnimateText AnimateLabel
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    For i = 1 To dirs: HiddenRecovery PathFileHidden & dirbuff(i) & _
        vbBackslash, AnimateLabel: Next i
End Sub
Public Function ShredVirusNow(ByVal sPathDel As String) As Long
'Dim Strength As Integer
'Strength = CInt(frmMain.cboStrength.Text)

'If Strength < 4 Then
'    lblStrength.Caption = "Very quick"
'ElseIf Strength < 6 Then
'    lblStrength.Caption = "Quick"
'ElseIf Strength < 9 Then
'    lblStrength.Caption = "Normal"
'ElseIf Strength < 12 Then
'    lblStrength.Caption = "Strong"
'ElseIf Strength < 19 Then
'    lblStrength.Caption = "Paranoid"
'ElseIf Strength < 28 Then
'    lblStrength.Caption = "Very paranoid"
'Else
'    lblStrength.Caption = "Maximum destruction"
'End If

    On Error Resume Next
    SetFileAttributes sPathDel, FILE_ATTRIBUTE_NORMAL
    ShredFile (sPathDel), CInt(frmMain.cboStrength.Text)
    Exit Function
End Function

Public Function KillVirusNow(ByVal sPathDel As String) As Long
    On Error Resume Next
    SetFileAttributes sPathDel, FILE_ATTRIBUTE_NORMAL
    DeleteFile (sPathDel)
    Exit Function
End Function

