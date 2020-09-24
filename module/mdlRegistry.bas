Attribute VB_Name = "mdlRegistry"
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

Private Declare Function RegCreateKey Lib _
    "advapi32.dll" Alias "RegCreateKeyA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib _
    "advapi32.dll" Alias "RegSetValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    lpData As Any, _
    ByVal cbData As Long) As Long
Private Declare Function RegOpenKey Lib _
    "advapi32.dll" Alias "RegOpenKeyA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib _
    "advapi32.dll" Alias "RegDeleteValueA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib _
    "advapi32.dll" Alias "RegOpenKeyExA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long
Private Declare Function RegCloseKey Lib _
    "advapi32.dll" ( _
    ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib _
    "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpData As String, _
    lpcbData As Long) As Long
Private Declare Function RegQueryValueExA Lib _
    "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByRef lpData As Long, _
    lpcbData As Long) As Long

Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1&
Private Const REG_DWORD = 4&
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const KEY_READ = STANDARD_RIGHTS_READ Or _
    KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY

Dim MainKeyHandle As REG
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim lKey As Long

Public Enum REG
    HKEY_CURRENT_USER = &H80000001
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Const SMP_KEY As String = "Software\Simple Machine Protect"

Public Function CreateRegistryKey(hKey As REG, sSubKey As String)
    Dim lReg As Long
    RegCreateKey hKey, sSubKey, lReg
    RegCloseKey lReg
End Function

Public Function GetDWORDValue(hKey As REG, SubKey As String, Entry As String)
    Dim ret As Long
    rtn = RegOpenKeyEx(hKey, SubKey, 0, KEY_READ, ret)
    If rtn = ERROR_SUCCESS Then
        rtn = RegQueryValueExA(ret, Entry, 0, REG_DWORD, lBuffer, 4)
        If rtn = ERROR_SUCCESS Then
            rtn = RegCloseKey(ret)
            GetDWORDValue = lBuffer
        Else
            GetDWORDValue = "Error"
        End If
    Else
        GetDWORDValue = "Error"
    End If
End Function

Public Function GetSTRINGValue(hKey As REG, SubKey As String, Entry As String)
    Dim ret As Long
    rtn = RegOpenKeyEx(hKey, SubKey, 0, KEY_READ, ret)
    If rtn = ERROR_SUCCESS Then
        sBuffer = Space(255)
        lBufferSize = Len(sBuffer)
        rtn = RegQueryValueEx(ret, Entry, 0, REG_SZ, sBuffer, lBufferSize)
        If rtn = ERROR_SUCCESS Then
            rtn = RegCloseKey(ret)
            sBuffer = Trim(sBuffer)
            GetSTRINGValue = Left(sBuffer, Len(sBuffer) - 1)
        Else
            GetSTRINGValue = "Error"
        End If
    Else
        GetSTRINGValue = "Error"
    End If
End Function

Public Function CreateDwordValue(hKey As REG, SubKey As String, _
    strValueName As String, dwordData As Long) As Long
    Dim ret As Long
    RegCreateKey hKey, SubKey, ret
    CreateDwordValue = RegSetValueEx(ret, strValueName, 0, REG_DWORD, dwordData, 4)
    RegCloseKey ret
End Function

Public Function CreateStringValue(hKey As REG, SubKey As String, _
    strValueName As String, strdata As String) As Long
    Dim ret As Long
    RegCreateKey hKey, SubKey, ret
    CreateStringValue = RegSetValueEx(ret, strValueName, 0, _
        REG_SZ, ByVal strdata, Len(strdata))
    RegCloseKey ret
End Function

Public Function DeleteValue(hKey As REG, SubKey As String, lpValName As String) As Long
    Dim ret As Long
    RegOpenKey hKey, SubKey, ret
    DeleteValue = RegDeleteValue(ret, lpValName)
    RegCloseKey ret
End Function



