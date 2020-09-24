VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information"
   ClientHeight    =   900
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   3960
   ControlBox      =   0   'False
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin SimpleMachineProtect.XP_ProgressBar prgInfo 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
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
      Color           =   32896
      Scrolling       =   1
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait"
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   640
      Width           =   855
   End
   Begin VB.Menu mnuProcess 
      Caption         =   "Process List"
      Visible         =   0   'False
      Begin VB.Menu popProcess 
         Caption         =   "Process Information"
         Index           =   0
      End
      Begin VB.Menu popProcess 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu popProcess 
         Caption         =   "Refresh Process List"
         Index           =   2
      End
      Begin VB.Menu popProcess 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu popProcess 
         Caption         =   "Open Containing Folder"
         Index           =   4
      End
      Begin VB.Menu popProcess 
         Caption         =   "Dos Prompt"
         Index           =   5
      End
      Begin VB.Menu popProcess 
         Caption         =   "Show File Properties"
         Index           =   6
      End
      Begin VB.Menu popProcess 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu popProcess 
         Caption         =   "Threads"
         Index           =   8
         Begin VB.Menu popThreads 
            Caption         =   "Resume"
            Index           =   0
         End
         Begin VB.Menu popThreads 
            Caption         =   "Suspend"
            Index           =   1
         End
      End
      Begin VB.Menu popProcess 
         Caption         =   "Priority"
         Index           =   9
         Begin VB.Menu popBase 
            Caption         =   "Realtime"
            Index           =   0
         End
         Begin VB.Menu popBase 
            Caption         =   "High"
            Index           =   1
         End
         Begin VB.Menu popBase 
            Caption         =   "Normal"
            Index           =   2
         End
         Begin VB.Menu popBase 
            Caption         =   "Idle"
            Index           =   3
         End
      End
      Begin VB.Menu popProcess 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu popProcess 
         Caption         =   "Terminate"
         Index           =   11
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Online Help"
      Visible         =   0   'False
      Begin VB.Menu popHelp 
         Caption         =   "Developer Homepage"
         Index           =   0
      End
      Begin VB.Menu popHelp 
         Caption         =   "Project Homepage"
         Index           =   1
      End
      Begin VB.Menu popHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu popHelp 
         Caption         =   "Product News"
         Index           =   3
      End
      Begin VB.Menu popHelp 
         Caption         =   "Frequently Asked Questions, Support"
         Index           =   4
      End
      Begin VB.Menu popHelp 
         Caption         =   "Feedback (Bug, Suggestion)..."
         Index           =   5
      End
      Begin VB.Menu popHelp 
         Caption         =   "Check For Updates"
         Index           =   6
      End
      Begin VB.Menu popHelp 
         Caption         =   "Submit Virus"
         Index           =   7
      End
      Begin VB.Menu popHelp 
         Caption         =   "X-Suck91 Community"
         Index           =   8
      End
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "Systray Menu"
      Visible         =   0   'False
      Begin VB.Menu popSystray 
         Caption         =   "Restore"
         Index           =   0
      End
      Begin VB.Menu popSystray 
         Caption         =   "Quarantine"
         Index           =   1
      End
      Begin VB.Menu popSystray 
         Caption         =   "Check For Updates"
         Index           =   2
      End
      Begin VB.Menu popSystray 
         Caption         =   "Online Help"
         Index           =   3
      End
      Begin VB.Menu popSystray 
         Caption         =   "Submit Virus"
         Index           =   4
      End
      Begin VB.Menu popSystray 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu popSystray 
         Caption         =   "Exit Windows"
         Index           =   6
         Begin VB.Menu popExit 
            Caption         =   "Log Off"
            Index           =   0
         End
         Begin VB.Menu popExit 
            Caption         =   "Restart"
            Index           =   1
         End
         Begin VB.Menu popExit 
            Caption         =   "Shutdown"
            Index           =   2
         End
      End
      Begin VB.Menu popSystray 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu popSystray 
         Caption         =   "Exit"
         Index           =   8
      End
      Begin VB.Menu popSystray 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu popSystray 
         Caption         =   "Real Time Protector"
         Index           =   10
      End
   End
End
Attribute VB_Name = "frmInfo"
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

Private Sub Form_Activate()
    lblInfo.ForeColor = prgInfo.Color
    AlwaysOnTop Me.hwnd, True
    Screen.MousePointer = 13
    Dim i As Long
    i = 0
    Do
        DoEvents
        i = i + 2
        prgInfo.Value = i
        Sleep 1
    Loop Until i = 100
    Sleep 100
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub mnuProcess_Click()
    Dim i As Integer
    Dim lPID As Long
    Dim sBase As String
    sBase = LCase(frmMain.lvwProcessExplorer.SelectedItem.SubItems(8))
    lPID = frmMain.lvwProcessExplorer.SelectedItem.SubItems(5)
    If GetAppID = lPID Then
        popProcess(8).Enabled = False
        popProcess(11).Enabled = False
    Else
        popProcess(8).Enabled = True
        popProcess(11).Enabled = True
    End If
    For i = popBase.LBound To popBase.UBound
        popBase.Item(i).Checked = False
        Select Case sBase
            Case Is = "realtime"
                popBase.Item(0).Checked = True
            Case Is = "high"
                popBase.Item(1).Checked = True
            Case Is = "normal"
                popBase.Item(2).Checked = True
            Case Is = "idle"
                popBase.Item(3).Checked = True
            Case Else
                popBase.Item(i).Checked = False
        End Select
    Next i
End Sub

Private Sub popBase_Click(Index As Integer)
    Dim lBase As Long
    Select Case Index
        Case 0
            lBase = SetBasePriority(frmMain.lvwProcessExplorer, _
                5, REALTIME_PRIORITY_CLASS)
        Case 1
            lBase = SetBasePriority(frmMain.lvwProcessExplorer, _
                5, HIGH_PRIORITY_CLASS)
        Case 2
            lBase = SetBasePriority(frmMain.lvwProcessExplorer, _
                5, NORMAL_PRIORITY_CLASS)
        Case 3
            lBase = SetBasePriority(frmMain.lvwProcessExplorer, _
                5, IDLE_PRIORITY_CLASS)
    End Select
    popProcess_Click 2
End Sub

Private Sub popExit_Click(Index As Integer)
    Select Case Index
        Case 0
            ExitWindowsNow "logoff"
        Case 1
            ExitWindowsNow "reboot"
        Case 2
            ExitWindowsNow "shutdown"
    End Select
End Sub

Private Sub popHelp_Click(Index As Integer)
    With Me
        Select Case Index
            Case 0
                OnlineHelp .hwnd, "www.e-freshware.com"
            Case 1
                OnlineHelp .hwnd, "sourceforge.net/projects/smpav"
            Case 3
                OnlineHelp .hwnd, SMP_SITE
            Case 4
                OnlineHelp .hwnd, SMP_SITE & "/index.php?page=faq"
            Case 5
                OnlineHelp .hwnd, SMP_SITE & "/index.php?page=contact"
            Case 6
                OnlineHelp .hwnd, SMP_SITE & "/index.php?page=download"
            Case 7
                MsgBox "Send your virus sample by e-mail" & vbNewLine & "to angga_seto@plasa.com" & vbNewLine & "with subject Sample Virus", vbInformation + vbSystemModal, "Send your virus sample"
            Case 8
            OnlineHelp .hwnd, "xsuck91.xm.com/xsuck91AV.html"
            End Select
    End With
End Sub

Private Sub popProcess_Click(Index As Integer)
    Dim lPID As Long
    Dim sProc As String
    With frmMain
        lPID = .lvwProcessExplorer.SelectedItem.SubItems(5)
        sProc = .lvwProcessExplorer.SelectedItem.SubItems(1)
        Select Case Index
            Case 0
                frmProcInfo.Show vbModal
            Case 2
                Screen.MousePointer = 13
                NTProcessList .lvwProcessExplorer, .ilsProcessExplorer
                .lvwProcessExplorer.SetFocus
                Screen.MousePointer = vbDefault
            Case 4
                OpenInFolder .lvwProcessExplorer, 1
                popProcess_Click 2
            Case 5
                OpenDosPrompt .lvwProcessExplorer, 1
                popProcess_Click 2
            Case 6
                ShowFileProperties .hwnd, .lvwProcessExplorer, 1
                popProcess_Click 2
            Case 11
                TerminateProcessID .lvwProcessExplorer, 5
                popProcess_Click 2
        End Select
    End With
End Sub

Private Sub popSystray_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            With frmMain
                .Show
                .OnSystray.Visible = False
            End With
            Unload frmFinish
            Case 1
            frmKarantina.Show
        Case 2
            OnlineHelp Me.hwnd, SMP_SITE & "/index.php?page=download"
        Case 3
            OnlineHelp Me.hwnd, SMP_SITE & "/index.php"
        Case 4
            OnlineHelp Me.hwnd, SMP_SITE & "/index.php?page=upload"
        Case 8
            If MsgBox("Are you sure you want to exit?", _
                vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation") = vbYes Then
                ExitNow
            End If
            Case 10
            Load frmRTP
                      
    End Select
End Sub

Private Sub popThreads_Click(Index As Integer)
    Select Case Index
        Case 0
            SetSuspendResumeThread frmMain.lvwProcessExplorer, 5, False
        Case 1
            SetSuspendResumeThread frmMain.lvwProcessExplorer, 5, True
    End Select
    popProcess_Click 2
End Sub
