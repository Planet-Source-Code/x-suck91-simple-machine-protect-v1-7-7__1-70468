VERSION 5.00
Begin VB.Form frmLoading 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Loading"
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   Picture         =   "frmLoading.frx":0000
   ScaleHeight     =   2730
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SimpleMachineProtect.XP_ProgressBar prgLoad 
      Height          =   300
      Left            =   720
      TabIndex        =   5
      Top             =   2025
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   12937777
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Timer tmrFadeout 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4800
      Top             =   120
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5160
      Top             =   120
   End
   Begin VB.Timer tmrMain 
      Interval        =   2000
      Left            =   5520
      Top             =   120
   End
   Begin VB.Label lblVersionHeader 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "VERSION 1.7.7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image imgLoad 
      Height          =   855
      Index           =   1
      Left            =   1275
      Picture         =   "frmLoading.frx":360C2
      Top             =   480
      Width           =   3525
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2007 - 2008 BJ`s Software Studios && X-Suck91 Community"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Label lblLoad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Application"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   150
      TabIndex        =   3
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait Simple Machine Protect is configuring environment."
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   1740
      Width           =   5775
   End
   Begin VB.Image imgLoad 
      Height          =   180
      Index           =   0
      Left            =   120
      Picture         =   "frmLoading.frx":3CCE4
      ToolTipText     =   "Made In Indonesia"
      Top             =   135
      Width           =   270
   End
   Begin VB.Label lblMalang 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SENG NGGAWE WONG MALANG && URANG SUNDA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Dim lAlpha As Integer

Private Sub Form_Load()
    On Error Resume Next
    If App.PrevInstance Then End
    AlwaysOnTop Me.hwnd, True
    lAlpha = 255
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    AlwaysOnTop Me.hwnd, True
End Sub

Private Sub tmrFadeout_Timer()
    If lAlpha > 0 Then
        DoEvents
        lAlpha = lAlpha - 5
        MakeTransparent Me.hwnd, lAlpha
    Else
        lAlpha = 0
        frmMain.Show
    End If
End Sub

Private Sub tmrMain_Timer()
    tmrStart.Enabled = True
    prgLoad.Orientation = ccOrientationHorizontal
    tmrMain.Enabled = False
End Sub

Private Sub tmrStart_Timer()
    With prgLoad
        If .Value < 100 Then
            DoEvents
            .Value = .Value + 1
            If .Value = 20 Then
                lblLoad.Caption = "Building Database": DoEvents
                LoadVirusDatabase
                Sleep 1000
            End If
            If .Value = 60 Then
                lblLoad.Caption = "Scanning Memory": DoEvents
                ScanProcess False
                Sleep 1000
            End If
            If .Value = 90 Then
                tmrStart.Enabled = False
                lblLoad.Caption = "Creating Application Environtment": DoEvents
                FillSystemOptimizer frmMain.lvwSystemOptimizer
                CheckOptimizer frmMain.lvwSystemOptimizer
                'Dim lReg As Long
                'lReg = GetDWORDValue(HKEY_CURRENT_USER, SMP_KEY, "Register")
                'If lReg = True Then
                '   LoadAppSettings
                'Else
                '    DefaultAppSettings
                'End If
                Dim i As Integer
                For i = 1 To VirusName.Count
                    frmMain.lstVirList.AddItem VirusName(i)
                Next i
                Sleep 1000
                tmrStart.Enabled = True
            End If
        Else
            .Value = 100
            lblLoad.Caption = "Load Complete"
            tmrStart.Enabled = False
            tmrFadeout.Enabled = True
        End If
    End With
End Sub
