VERSION 5.00
Begin VB.UserControl OnSystray 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   270
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   270
   ScaleWidth      =   270
   Begin VB.Image imgIconSystray 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   0
      Picture         =   "OnSystray.ctx":0000
      Top             =   0
      Width           =   270
   End
End
Attribute VB_Name = "OnSystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

' OnSysTray User Control - A simple systray menu.
' Created by Bagus Judistirah (C) 2008
' GNU General Public License

Option Explicit

Private Declare Function Shell_NotifyIcon Lib _
    "shell32" (ByVal dwMessage As Long, _
    pnid As NOTIFYICONDATA) As Boolean

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    UID As Long
    uFlags As Long
    uCallBackmessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private NID As NOTIFYICONDATA
Private var_visible As Boolean
Private var_tooltiptext As String
Private var_icon As Picture

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONUP = &H205
Private Const WM_MOUSEOVER = &H200
Private Const smp_var_visible = False
Private Const smp_var_tooltiptext = ""
Private Const smp_var_uid = 65535

Public Event DoubleClick()
Public Event MouseUp(Button As Integer)
Public Event MouseMove()

Public Property Get Icon() As Picture
    Set Icon = var_icon
End Property

Public Property Set Icon(ByVal stdNewIcon As Picture)
    Set var_icon = stdNewIcon
    If stdNewIcon Is Nothing Then
        Visible = False
    Else
        If var_visible Then
            NID.uFlags = NIF_ICON
            NID.hIcon = var_icon
            Shell_NotifyIcon NIM_MODIFY, NID
        End If
    End If
    PropertyChanged "Icon"
End Property

Public Property Get ToolTipText() As String
    ToolTipText = var_tooltiptext
End Property

Public Property Let ToolTipText(ByVal stdNewToolTip As String)
    var_tooltiptext = Trim(stdNewToolTip)
    NID.uFlags = NIF_TIP
    NID.szTip = var_tooltiptext & vbNullChar
    Shell_NotifyIcon NIM_MODIFY, NID
    PropertyChanged "ToolTipText"
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_MemberFlags = "400"
    Visible = var_visible
End Property

Public Property Let Visible(ByVal stdNewVisible As Boolean)
    If var_visible = stdNewVisible Then Exit Property
    var_visible = stdNewVisible
    If var_visible Then
        If Ambient.UserMode Then
            NID.cbSize = Len(NID)
            NID.hwnd = UserControl.hwnd
            NID.UID = Int((Rnd * smp_var_uid) + 1)
            NID.uFlags = NIF_MESSAGE
            If Not var_icon Is Nothing Then
                NID.uFlags = NID.uFlags + NIF_ICON
                NID.hIcon = var_icon
            End If
            If var_tooltiptext <> "" Then
                NID.uFlags = NID.uFlags + NIF_TIP
                NID.szTip = var_tooltiptext & vbNullChar
            End If
            NID.uCallBackmessage = WM_MOUSEMOVE
            Shell_NotifyIcon NIM_ADD, NID
        End If
    Else
        Shell_NotifyIcon NIM_DELETE, NID
    End If
    PropertyChanged "Visible"
End Property

Private Sub UserControl_InitProperties()
    Set var_icon = LoadPicture("")
    var_tooltiptext = smp_var_tooltiptext
    var_visible = smp_var_visible
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set var_icon = PropBag.ReadProperty("Icon", Nothing)
    var_tooltiptext = PropBag.ReadProperty("ToolTipText", smp_var_tooltiptext)
    var_visible = PropBag.ReadProperty("Visible", smp_var_visible)
End Sub

Private Sub UserControl_Resize()
    Static inloop As Boolean
    If inloop Then Exit Sub
    inloop = True
    Height = imgIconSystray.Height
    Width = imgIconSystray.Width
    inloop = False
End Sub

Private Sub UserControl_Terminate()
    Shell_NotifyIcon NIM_DELETE, NID
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Icon", var_icon, Nothing)
    Call PropBag.WriteProperty("ToolTipText", var_tooltiptext, smp_var_tooltiptext)
    Call PropBag.WriteProperty("Visible", var_visible, smp_var_visible)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, _
    x As Single, y As Single)
    Select Case x / Screen.TwipsPerPixelX
        Case WM_LBUTTONDBLCLK
            RaiseEvent DoubleClick
        Case WM_RBUTTONUP
            RaiseEvent MouseUp(vbRightButton)
        Case WM_MOUSEOVER
            RaiseEvent MouseMove
    End Select
End Sub
