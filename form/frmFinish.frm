VERSION 5.00
Begin VB.Form frmFinish 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Finished"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinish.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFinish.frx":08CA
   ScaleHeight     =   2895
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin SimpleMachineProtect.chameleonButton cmdClose 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
      _extentx        =   2566
      _extenty        =   661
      btype           =   3
      tx              =   "Close"
      enab            =   -1  'True
      font            =   "frmFinish.frx":1E60C
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmFinish.frx":1E634
      picn            =   "frmFinish.frx":1E652
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Image imgSMP 
      Height          =   480
      Index           =   0
      Left            =   1320
      Picture         =   "frmFinish.frx":1EBEE
      Top             =   120
      Width           =   1980
   End
   Begin VB.Label lblScanned 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblScanned 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblScanned 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblScanned 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblEmpty 
      BackStyle       =   0  'Transparent
      Caption         =   "File Detected"
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblEmpty 
      BackStyle       =   0  'Transparent
      Caption         =   "File Repaired"
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblEmpty 
      BackStyle       =   0  'Transparent
      Caption         =   "File Infected"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblEmpty 
      BackStyle       =   0  'Transparent
      Caption         =   "File Scanned"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblEmpty 
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Result"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   260
      TabIndex        =   1
      Top             =   680
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu popVirAct 
         Caption         =   "Show File Properties"
         Index           =   0
      End
      Begin VB.Menu popVirAct 
         Caption         =   "Open Containing Folder"
         Index           =   1
      End
      Begin VB.Menu popVirAct 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu popVirAct 
         Caption         =   "Karantina Virus"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmFinish"
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AlwaysOnTop Me.hwnd, True
    lblScanned(0).Caption = CheckValueData(nFile, "scanned")
    lblScanned(1).Caption = CheckValueData(nInfect, "infected")
    lblScanned(2).Caption = CheckValueData(nRepair, "repaired")
    lblScanned(3).Caption = CheckValueData(frmMain.lvwVirFound.ListItems.Count, _
        "detected")
    FinishAlert
End Sub

Private Sub popVirAct_Click(Index As Integer)
Select Case Index
Case 0: ShowFileProperties1 Me.hwnd, frmMain.lvwVirFound, 1
Case 1: LocateFile frmMain.lvwVirFound, 1
'Case 4: frmMain.DeleteVirus
Case 5: frmMain.Karantina

End Select
End Sub
