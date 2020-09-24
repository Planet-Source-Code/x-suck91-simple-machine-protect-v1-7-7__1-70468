Attribute VB_Name = "mdlPerformance"
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

Private Declare Sub GlobalMemoryStatus Lib _
    "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
    dwLength        As Long
    dwMemoryLoad    As Long
    dwTotalPhys     As Long
    dwAvailPhys     As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual  As Long
    dwAvailVirtual  As Long
End Type

Private Enum PDH_STATUS
    PDH_CSTATUS_VALID_DATA = &H0
    PDH_CSTATUS_NEW_DATA = &H1
End Enum

Private Declare Function PdhOpenQuery Lib _
    "PDH.DLL" (ByVal Reserved As Long, _
    ByVal dwUserData As Long, _
    ByRef hQuery As Long) As PDH_STATUS
Private Declare Function PdhVbAddCounter Lib _
    "PDH.DLL" (ByVal QueryHandle As Long, _
    ByVal CounterPath As String, _
    ByRef CounterHandle As Long) As PDH_STATUS
Private Declare Function PdhCollectQueryData Lib _
    "PDH.DLL" (ByVal QueryHandle As Long) As PDH_STATUS
Private Declare Function PdhVbGetDoubleCounterValue Lib _
    "PDH.DLL" (ByVal CounterHandle As Long, _
    ByRef CounterStatus As Long) As Double

Private Type CounterInfo
    hCounter As Long
    strName As String
End Type

Dim pdhStatus As PDH_STATUS
Dim Counters(0 To 99) As CounterInfo
Dim hQuery As Long

Public Sub MemoryInfo(lAPageFile As Label, lAPhys As Label, _
    lAVirtual As Label, lTPageFile As Label, lTPhys As Label, _
    lTVirtual As Label, lMemUsage As Label)
    Dim MemStat As MEMORYSTATUS
    MemStat.dwLength = Len(MemStat)
    GlobalMemoryStatus MemStat
    lAPageFile.Caption = Format(MemStat.dwAvailPageFile _
        / 1024, "###,###,###") & " KB"
    lAPhys.Caption = Format(MemStat.dwAvailPhys / 1024, _
        "###,###,###") & " KB"
    lAVirtual.Caption = Format(MemStat.dwAvailVirtual / _
        1024, "###,###,###") & " KB"
    lTPageFile.Caption = Format(MemStat.dwTotalPageFile _
        / 1024, "###,###,###") & " KB"
    lTPhys.Caption = Format(MemStat.dwTotalPhys / 1024, _
        "###,###,###") & " KB"
    lTVirtual.Caption = Format(MemStat.dwTotalVirtual / _
        1024, "###,###,###") & " KB"
    lMemUsage.Caption = MemStat.dwMemoryLoad & " %"
End Sub

Public Sub UpdateValues(lblCPU As Label)
    Dim dblCounterValue As Double
    Dim pdhStatus As Long
    Dim strInfo As String
    Dim i As Long
    PdhCollectQueryData (hQuery)
    i = 0
    dblCounterValue = PdhVbGetDoubleCounterValue( _
        Counters(i).hCounter, pdhStatus)
    If (pdhStatus = PDH_CSTATUS_VALID_DATA) Or (pdhStatus _
        = PDH_CSTATUS_NEW_DATA) Then
        lblCPU.Caption = Abs(Fix(dblCounterValue)) & " %"
    End If
End Sub

Private Sub AddCounter(strCounterName As String, _
    hQuery As Long)
    Dim pdhStatus As PDH_STATUS
    Dim hCounter As Long, currentCounterIdx As Long
    pdhStatus = PdhVbAddCounter(hQuery, strCounterName, _
        hCounter)
    Counters(currentCounterIdx).hCounter = hCounter
    Counters(currentCounterIdx).strName = strCounterName
    currentCounterIdx = currentCounterIdx + 1
End Sub

Public Sub GetCPUInfo(lblCPU As Label)
    pdhStatus = PdhOpenQuery(0, 1, hQuery)
    AddCounter "\Processor(0)\% Processor Time", hQuery
    UpdateValues lblCPU
End Sub

