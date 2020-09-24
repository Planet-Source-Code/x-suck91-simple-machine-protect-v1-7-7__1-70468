Attribute VB_Name = "mdlSystemOptimizer"
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

Private Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Private Declare Function SHEmptyRecycleBin Lib _
    "shell32.dll" Alias "SHEmptyRecycleBinA" ( _
    ByVal hwnd As Long, _
    ByVal pszRootPath As String, _
    ByVal dwFlags As Long) As Long
    
Private Const SHERB_NOCONFIRMATION = &H1
Private Const SHERB_NOPROGRESSUI = &H2
Private Const SHERB_NOSOUND = &H4
Private Const WITHOUT_ANY = SHERB_NOCONFIRMATION Or _
    SHERB_NOPROGRESSUI Or SHERB_NOSOUND

Public Sub FillSystemOptimizer(lvw As ListView)
    On Error GoTo ErrHandle
    Dim SysItem(22) As String
    SysItem(1) = "Always Unload Unused Dynamic Libraries (DLLs) from memory."
    SysItem(2) = "Auto End Task Program when Not Responding."
    SysItem(3) = "Clean Shutdown."
    SysItem(4) = "Clear Paging File at shutdown."
    SysItem(5) = "Clear Recents Docs on exit Windows."
    SysItem(6) = "Delete Virtual Memory at shutdown."
    SysItem(7) = "Disable Desktop Cleanup Wizard."
    SysItem(8) = "Disable Low Disk Space Warning."
    SysItem(9) = "Disable Recent History."
    SysItem(10) = "Disable Windows Animation."
    SysItem(11) = "Display BSOD (Blue Screen Of Death) when system crashes."
    SysItem(12) = "Do not move deleted files to the Recycle Bin."
    SysItem(13) = "Do not display the Getting Started welcome screen at logon."
    SysItem(14) = "Optimize Desktop Process."
    SysItem(15) = "Optimize for Fast Operating System Boot."
    SysItem(16) = "Optimize Hard Drive."
    SysItem(17) = "Optimize Start Menu."
    SysItem(18) = "Quick Rebooting Operating System."
    SysItem(19) = "Remove Undock Option from Start Menu."
    SysItem(20) = "Removes Temporary Internet Files."
    SysItem(21) = "Turn Off Autoplay."
    SysItem(22) = "Use Large System Cache."
    Dim i As Integer
    For i = 1 To UBound(SysItem)
        lvw.ListItems.Add , , SysItem(i), , 4
    Next i
ErrHandle:
End Sub

Public Sub CheckOptimizer(lvw As ListView)
    On Error Resume Next
    Dim sRegStrOpt As String
    Dim lDword As Integer
    With lvw
        sRegStrOpt = GetSTRINGValue(HKEY_LOCAL_MACHINE, _
            "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AlwaysUnloadDLL", "")
        If Val(sRegStrOpt) <> 1 Then
            .ListItems.Item(1).Checked = False
        Else
            .ListItems.Item(1).Checked = True
        End If
        sRegStrOpt = GetSTRINGValue(HKEY_CURRENT_USER, "Control Panel\Desktop", _
            "AutoEndTasks")
        If Val(sRegStrOpt) <> 1 Then
            .ListItems.Item(2).Checked = False
        Else
            .ListItems.Item(2).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_CURRENT_USER, _
            "Software\Microsoft\Windows\CurrentVersion\Explorer", "CleanShutdown")
        If Val(lDword) <> 1 Then
            .ListItems.Item(3).Checked = False
        Else
            .ListItems.Item(3).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, _
            "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management", _
            "ClearPageFileAtShutdown")
        If Val(lDword) <> 1 Then
            .ListItems.Item(4).Checked = False
        Else
            .ListItems.Item(4).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_CURRENT_USER, _
            "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
            "ClearRecentDocsOnExit")
        If Val(lDword) <> 1 Then
            .ListItems.Item(5).Checked = False
        Else
            .ListItems.Item(5).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, _
            "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management", _
            "DisablePagingExecutive")
        If Val(lDword) <> 1 Then
            .ListItems.Item(6).Checked = False
        Else
            .ListItems.Item(6).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_CURRENT_USER, _
            "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
            "NoDesktopCleanupWizard")
        If Val(lDword) <> 1 Then
            .ListItems.Item(7).Checked = False
        Else
            .ListItems.Item(7).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_CURRENT_USER, _
            "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
            "NoLowDiskSpaceChecks")
        If Val(lDword) <> 1 Then
            .ListItems.Item(8).Checked = False
        Else
            .ListItems.Item(8).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_CURRENT_USER, _
            "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
            "NoRecentDocsHistory")
        If Val(lDword) <> 1 Then
            .ListItems.Item(9).Checked = False
        Else
            .ListItems.Item(9).Checked = True
        End If
        sRegStrOpt = GetSTRINGValue(HKEY_CURRENT_USER, _
            "Control Panel\Desktop\WindowMetrics", "MinAnimate")
        If Val(sRegStrOpt) <> 1 Then
            .ListItems.Item(10).Checked = True
        Else
            .ListItems.Item(10).Checked = False
        End If
        lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, _
        "SYSTEM\ControlCurrentSet\Control\CrashControl", "AutoReboot")
        If Val(lDword) <> 1 Then
            .ListItems.Item(11).Checked = False
        Else
            .ListItems.Item(11).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_CURRENT_USER, _
            "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
            "NoRecycleFiles")
        If Val(lDword) <> 1 Then
            .ListItems.Item(12).Checked = False
        Else
            .ListItems.Item(12).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_CURRENT_USER, _
            "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
            "NoWelcomeScreen")
        If Val(lDword) <> 1 Then
            .ListItems.Item(13).Checked = False
        Else
            .ListItems.Item(13).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_CURRENT_USER, _
            "Software\Microsoft\Windows\CurrentVersion\Explorer", _
            "DesktopProcess")
        If Val(lDword) <> 1 Then
            .ListItems.Item(14).Checked = False
        Else
            .ListItems.Item(14).Checked = True
        End If
        sRegStrOpt = GetSTRINGValue(HKEY_LOCAL_MACHINE, _
            "SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction", "Enable")
        If Left(sRegStrOpt, 1) <> "Y" Then
            .ListItems.Item(15).Checked = False
        Else
            .ListItems.Item(15).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, _
            "SOFTWARE\Microsoft\Windows\CurrentVersion\OptimalLayout", _
            "EnableAutoLayout")
        If Val(lDword) <> 1 Then
            .ListItems.Item(16).Checked = False
        Else
            .ListItems.Item(16).Checked = True
        End If
        sRegStrOpt = GetSTRINGValue(HKEY_CURRENT_USER, "Control Panel\Desktop", _
            "MenuShowDelay")
        If Val(sRegStrOpt) <> 1 Then
            .ListItems.Item(17).Checked = False
        Else
            .ListItems.Item(17).Checked = True
        End If
        sRegStrOpt = GetSTRINGValue(HKEY_LOCAL_MACHINE, _
            "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", _
            "EnableQuickReboot")
        If Left(sRegStrOpt, 1) <> "" Then
            .ListItems.Item(18).Checked = False
        Else
            .ListItems.Item(18).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, _
            "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
            "NoStartMenuEjectPC")
        If Val(lDword) <> 1 Then
            .ListItems.Item(19).Checked = False
        Else
            .ListItems.Item(19).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, _
            "Software\Microsoft\Windows\CurrentVersion\Internet Setting\Cache", _
            "Persistent")
        If Val(lDword) <> 1 Then
            .ListItems.Item(20).Checked = True
        Else
            .ListItems.Item(20).Checked = False
        End If
        lDword = GetDWORDValue(HKEY_CURRENT_USER, _
            "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
            "NoDriveTypeAutoRun")
        If Val(lDword) <> 99 Then
            .ListItems.Item(21).Checked = False
        Else
            .ListItems.Item(21).Checked = True
        End If
        lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, _
            "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management", _
            "LargeSystemCache")
        If Val(lDword) <> 1 Then
            .ListItems.Item(22).Checked = False
        Else
            .ListItems.Item(22).Checked = True
        End If
    End With
End Sub

Public Sub ExecuteOptimizer(lvw As ListView)
    With lvw.ListItems
        If .Item(1).Checked = True Then
            CreateStringValue HKEY_LOCAL_MACHINE, _
                "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AlwaysUnloadDLL", _
                "", "1"
        Else
            CreateStringValue HKEY_LOCAL_MACHINE, _
                "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AlwaysUnloadDLL", _
                 "", "0"
        End If
        If .Item(2).Checked = True Then
            CreateStringValue HKEY_USERS, _
                 ".DEFAULT\Control Panel\Desktop\", "AutoEndTasks", "1"
            CreateStringValue HKEY_CURRENT_USER, "Control Panel\Desktop\", _
                 "AutoEndTasks", "1"
        Else
            CreateStringValue HKEY_USERS, ".DEFAULT\Control Panel\Desktop\", _
                 "AutoEndTasks", "0"
            CreateStringValue HKEY_CURRENT_USER, "Control Panel\Desktop\", _
                 "AutoEndTasks", "0"
        End If
        If .Item(3).Checked = True Then
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Explorer\", _
                 "CleanShutdown", 1
        Else
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Explorer\", _
                 "CleanShutdown", 0
        End If
        If .Item(4).Checked = True Then
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", _
                 "ClearPageFileAtShutdown", 1
        Else
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", _
                 "ClearPageFileAtShutdown", 0
        End If
        If .Item(5).Checked = True Then
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "ClearRecentDocsOnExit", 1
        Else
            CreateDwordValue HKEY_CURRENT_USER _
                , "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "ClearRecentDocsOnExit", 0
        End If
        If .Item(6).Checked = True Then
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", _
                 "DisablePagingExecutive", 1
        Else
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", _
                 "DisablePagingExecutive", 0
        End If
        If .Item(7).Checked = True Then
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoDesktopCleanupWizard", 1
        Else
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoDesktopCleanupWizard", 0
        End If
        If .Item(8).Checked = True Then
            CreateDwordValue HKEY_CURRENT_USER, _
                "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoLowDiskSpaceChecks", 1
        Else
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoLowDiskSpaceChecks", 0
        End If
        If .Item(9).Checked = True Then
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoRecentDocsHistory", 1
        Else
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoRecentDocsHistory", 0
        End If
        If .Item(10).Checked = True Then
            CreateStringValue HKEY_CURRENT_USER, _
                 "Control Panel\Desktop\WindowMetrics\", "MinAnimate", "0"
        Else
            CreateStringValue HKEY_CURRENT_USER, _
                 "Control Panel\Desktop\WindowMetrics\", "MinAnimate", "1"
        End If
        If .Item(11).Checked = True Then
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SYSTEM\ControlCurrentSet\Control\CrashControl\", "AutoReboot", 1
        Else
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SYSTEM\ControlCurrentSet\Control\CrashControl\", "AutoReboot", 0
        End If
        If .Item(12).Checked = True Then
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoRecycleFiles", 1
        Else
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoRecycleFiles", 0
        End If
        If .Item(13).Checked = True Then
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoWelcomeScreen", 1
        Else
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoWelcomeScreen", 0
        End If
        If .Item(14).Checked = True Then
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Explorer\", _
                 "DesktopProcess", 1
        Else
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Explorer\", _
                 "DesktopProcess", 0
        End If
        If .Item(15).Checked = True Then
            CreateStringValue HKEY_LOCAL_MACHINE, _
                 "SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\", "Enable", "Y"
        Else
            CreateStringValue HKEY_LOCAL_MACHINE, _
                 "SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\", "Enable", "N"
        End If
        If .Item(16).Checked = True Then
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SOFTWARE\Microsoft\Windows\CurrentVersion\OptimalLayout\", _
                 "EnableAutoLayout", 1
        Else
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SOFTWARE\Microsoft\Windows\CurrentVersion\OptimalLayout\", _
                 "EnableAutoLayout", 0
        End If
        If .Item(17).Checked = True Then
            CreateStringValue HKEY_CURRENT_USER, "Control Panel\Desktop\", _
                 "MenuShowDelay", "1"
            CreateStringValue HKEY_USERS, ".DEFAULT\Control Panel\Desktop\", _
                 "MenuShowDelay", "1"
        Else
            CreateStringValue HKEY_CURRENT_USER, "Control Panel\Desktop\", _
                 "MenuShowDelay", "400"
            CreateStringValue HKEY_USERS, ".DEFAULT\Control Panel\Desktop\", _
                 "MenuShowDelay", "400"
        End If
        If .Item(18).Checked = True Then
            CreateStringValue HKEY_LOCAL_MACHINE, _
                 "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", _
                "EnableQuickReboot", ""
        Else
            DeleteValue HKEY_LOCAL_MACHINE, _
                 "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", _
                 "EnableQuickReboot"
        End If
        If .Item(19).Checked = True Then
            CreateDwordValue HKEY_CURRENT_USER _
                , "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoStartMenuEjectPC", 1
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoStartMenuEjectPC", 1
        Else
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoStartMenuEjectPC", 0
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoStartMenuEjectPC", 0
        End If
        If .Item(20).Checked = True Then
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Internet Setting\Cache\", _
                 "Persistent", 0
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Setting\Cache\", _
                 "Persistent", 0
        Else
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Internet Setting\Cache\", _
                 "Persistent", 1
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Setting\Cache\", _
                 "Persistent", 1
        End If
        If .Item(21).Checked = True Then
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoDriveTypeAutoRun", 99
        Else
            CreateDwordValue HKEY_CURRENT_USER, _
                 "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", _
                 "NoDriveTypeAutoRun", 0
        End If
        If .Item(22).Checked = True Then
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", _
                 "LargeSystemCache", 1
        Else
            CreateDwordValue HKEY_LOCAL_MACHINE, _
                 "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", _
                 "LargeSystemCache", 0
        End If
    End With
End Sub

Public Sub Clean_Registry()
    On Error Resume Next
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\System", _
        "DisableRegistryTools", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", _
        "DisableRegistryTools", 0
    CreateDwordValue HKEY_CURRENT_USER, _
         "Software\Microsoft\Windows\CurrentVersion\Policies\System\", _
         "DisableTaskMgr", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, _
         "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\", _
        "DisableTaskMgr", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Policies\Microsoft\Windows\System", _
        "DisableCMD", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
        "NoFolderOptions", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
        "NoFolderOptions", 0
    DeleteValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", _
        "Shell"
    CreateStringValue HKEY_CLASSES_ROOT, _
        "exefile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, _
        "lnkfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, _
        "piffile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, _
        "batfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, _
        "comfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, _
        "cmdfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, _
        "scrfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, _
        "regfile\shell\open\command", "", "regedit.exe %1"
    CreateStringValue HKEY_LOCAL_MACHINE, _
        "SYSTEM\CurrentControlSet\Control\SafeBoot\", "AlternateShell", "cmd.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Auto", "0"
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\", _
        "Disabled", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\", _
        "Disabled", 0
    CreateStringValue HKEY_CLASSES_ROOT, _
        "exefile", "", "Application"
    CreateDwordValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Policies\Microsoft\Windows\Installer", _
        "LimitSystemRestoreCheckpointing", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
        "NoDriveTypeAutoRun", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
        "NoTrayContextMenu", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
        "NoViewContextMenu", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\System", _
        "NoDispSettingsPage", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\System", _
        "NoDispBackgroundPage", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoScrSavPage", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\System", _
        "NoDispApprearancePage", 0
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCpl", 0
    CreateStringValue HKEY_CURRENT_USER, "Control Panel\Desktop\", "SCRNSAVE.EXE", ""
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", _
        "HideFileExt", 1
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "Hidden", 1
    CreateDwordValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", _
        "ShowSuperHidden", 1
    DeleteValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrive"
    DeleteValue HKEY_LOCAL_MACHINE, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrive"
    DeleteValue HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
        "DisableRegistryTools"
    DeleteValue HKEY_LOCAL_MACHINE, _
         "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
         "DisableRegistryTools"
    DeleteValue HKEY_LOCAL_MACHINE, _
         "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
         "shutdownwithoutlogon"
    DeleteValue HKEY_LOCAL_MACHINE, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
        "undockwithoutlogon"
End Sub

Public Function EmptyTheRecycle(hwnd As Long) As Long
    On Error Resume Next
    EmptyTheRecycle = SHEmptyRecycleBin(hwnd, vbNullString, WITHOUT_ANY)
End Function

Public Sub ClearJunkFile()
    On Error Resume Next
    Kill GetWindowsPath & "Prefetch\*.*"
    Kill GetWindowsPath & "Temp\*.*"
    Kill GetSpecialFolder(CSIDL_RECENT) & "\*.*"
    Kill GetSpecialFolder(CSIDL_HISTORY) & "\*.*"
    Kill GetSpecialFolder(CSIDL_PROFILER) & "\Cookies\*.*"
    Kill GetSpecialFolder(CSIDL_PROFILER) & "\Local Settings\Temp\*.*"
    Kill GetSpecialFolder(CSIDL_PROFILER) & _
        "\Local Settings\Temporary Internet Files\*.*"
        
End Sub

Public Function OpenXPTool(hwnd As Long, lpOperation As String) As Long
    On Error Resume Next
    OpenXPTool = ShellExecute(hwnd, vbNullString, lpOperation, _
        vbNullString, Left(GetWindowsPath, 3), 1)
End Function

Public Function OnlineHelp(hwnd As Long, strSite As String) As Long
    On Error Resume Next
    OnlineHelp = ShellExecute(hwnd, vbNullString, _
        "http://" & strSite, vbNullString, Left(GetWindowsPath, 3), 1)
End Function
