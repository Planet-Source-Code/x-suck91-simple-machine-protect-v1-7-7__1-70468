VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProcInfo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Process Information"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin SimpleMachineProtect.chameleonButton cmdExit 
      Height          =   375
      Left            =   6360
      TabIndex        =   42
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProcInfo.frx":08CA
      PICN            =   "frmProcInfo.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox picMod 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   240
      ScaleHeight     =   4695
      ScaleWidth      =   7455
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   7455
      Begin MSComctlLib.ListView lvwMod 
         Height          =   3375
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilsMod"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox picInfoMod 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1485
         Left            =   5760
         Picture         =   "frmProcInfo.frx":0E80
         ScaleHeight     =   1485
         ScaleWidth      =   1350
         TabIndex        =   34
         Top             =   2760
         Width           =   1350
      End
      Begin VB.TextBox txtMod 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   ": C:\WINDOWS\explorer.exe"
         Top             =   120
         Width           =   5895
      End
      Begin VB.TextBox txtMod 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   ": Windows Explorer"
         Top             =   360
         Width           =   5895
      End
      Begin VB.TextBox txtMod 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   ": Application"
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox txtMod 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   ": 6.00.2800.1106 (xpsp1.020828-1920)"
         Top             =   840
         Width           =   5895
      End
      Begin MSComctlLib.ImageList ilsMod 
         Left            =   120
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Label lblMod 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblMod 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblMod 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   37
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblMod 
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   36
         Top             =   840
         Width           =   1455
      End
      Begin VB.Image imgEmpty 
         Height          =   240
         Left            =   120
         Picture         =   "frmProcInfo.frx":3540
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label lblNotFound 
         BackStyle       =   0  'Transparent
         Caption         =   "Cannot open file. Module not found."
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   4320
         Width           =   3015
      End
   End
   Begin VB.PictureBox picGen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   240
      ScaleHeight     =   4695
      ScaleWidth      =   7455
      TabIndex        =   1
      Top             =   360
      Width           =   7455
      Begin VB.PictureBox picIco 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   240
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picInfo 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1485
         Left            =   5760
         Picture         =   "frmProcInfo.frx":3ACA
         ScaleHeight     =   1485
         ScaleWidth      =   1350
         TabIndex        =   3
         Top             =   2760
         Width           =   1350
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Explorer"
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
         Left            =   960
         TabIndex        =   27
         Top             =   120
         Width           =   6375
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "Microsoft Corporation"
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   360
         Width           =   6375
      End
      Begin VB.Line linMod 
         BorderColor     =   &H00C0C0C0&
         X1              =   240
         X2              =   7200
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "6.00.2900.2180 (xpsp_sp2_rtm.040803-2158)"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label lblEmpty 
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblEmpty 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblEmpty 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblEmpty 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblEmpty 
         BackStyle       =   0  'Transparent
         Caption         =   "Attributes"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblEmpty 
         BackStyle       =   0  'Transparent
         Caption         =   "Created"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblEmpty 
         BackStyle       =   0  'Transparent
         Caption         =   "Process ID"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label lblEmpty 
         BackStyle       =   0  'Transparent
         Caption         =   "Threads"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lblEmpty 
         BackStyle       =   0  'Transparent
         Caption         =   "Memory"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   16
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label lblEmpty 
         BackStyle       =   0  'Transparent
         Caption         =   "Priority"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   15
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "© Microsoft Corporation. All rights reserved."
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4320
         Width           =   7095
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   ": explorer.exe"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   1080
         Width           =   5895
      End
      Begin VB.Label lblLocation 
         BackStyle       =   0  'Transparent
         Caption         =   ": C:\WINDOWS\"
         Height          =   495
         Left            =   1440
         TabIndex        =   12
         Top             =   1320
         Width           =   5895
      End
      Begin VB.Label lblType 
         BackStyle       =   0  'Transparent
         Caption         =   ": Application"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1920
         Width           =   5895
      End
      Begin VB.Label lblSize 
         BackStyle       =   0  'Transparent
         Caption         =   ": 981 KB"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   2160
         Width           =   5895
      End
      Begin VB.Label lblAttributes 
         BackStyle       =   0  'Transparent
         Caption         =   ": A"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   2400
         Width           =   5895
      End
      Begin VB.Label lblPID 
         BackStyle       =   0  'Transparent
         Caption         =   ": 1600"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label lblThreads 
         BackStyle       =   0  'Transparent
         Caption         =   ": 10"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label lblMemory 
         BackStyle       =   0  'Transparent
         Caption         =   ": 28,564 KB"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label lblPriority 
         BackStyle       =   0  'Transparent
         Caption         =   ": Normal"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label lblCreated 
         BackStyle       =   0  'Transparent
         Caption         =   ": Wednesday, August 04, 2004, 7:00:00 PM"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   2640
         Width           =   5895
      End
   End
   Begin VB.Frame fraModule 
      BackColor       =   &H00FFFFFF&
      Caption         =   "General Information"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
   Begin SimpleMachineProtect.XP_ProgressBar prgModule 
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   5280
      Width           =   6210
      _ExtentX        =   10954
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
      Color           =   12937777
      Scrolling       =   1
   End
   Begin VB.Label lblChange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Module Used By Application"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   40
      Top             =   5400
      Width           =   2325
   End
End
Attribute VB_Name = "frmProcInfo"
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

Option Explicit

Private Declare Function DrawIcon Lib _
    "user32" (ByVal hdc As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib _
    "shell32.dll" Alias "ExtractIconA" ( _
    ByVal hInst As Long, _
    ByVal lpszExeFileName As String, _
    ByVal nIconIndex As Long) As Long

Dim GetIco As New clsGetIconFile

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    AlwaysOnTop Me.hwnd, True
    MakeInvisible
    MakeProgress
    MakeInfo
    GetModuleProcessID frmMain.lvwProcessExplorer, 5, lvwMod, ilsMod
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
    x As Single, y As Single)
    lblChange.ForeColor = &HC0&
End Sub

Private Sub lblChange_Click()
    If lblChange.Caption = "Module Used By Application" Then
        MakeProgress
        CheckModule
        lblChange.Caption = "General Information"
        fraModule.Caption = "Module Used By Application"
        picGen.Visible = False
        picMod.Visible = True
    Else
        MakeProgress
        MakeInfo
        lblChange.Caption = "Module Used By Application"
        fraModule.Caption = "General Information"
        picGen.Visible = True
        picMod.Visible = False
    End If
End Sub

Private Sub CheckModule()
    Dim i As Integer
    If lvwMod.ListItems.Count = 0 Then
        For i = lblMod.LBound To lblMod.UBound
            lblMod.Item(i).Visible = False
            txtMod.Item(i).Visible = False
        Next i
        lvwMod.Visible = False
    Else
        For i = lblMod.LBound To lblMod.UBound
            lblMod.Item(i).Visible = True
            txtMod.Item(i).Visible = True
        Next i
        lvwMod.Visible = True
    End If
End Sub

Private Sub FillTextBoxes()
    On Error Resume Next
    Dim i As Integer
    txtMod(0).Text = ": " & lvwMod.SelectedItem
    For i = 1 To 3
        txtMod(i).Text = ": " & lvwMod.SelectedItem.SubItems(i)
    Next i
End Sub

Private Sub lblChange_MouseMove(Button As Integer, Shift As Integer, _
    x As Single, y As Single)
    lblChange.ForeColor = &HC00000
End Sub

Private Sub lvwMod_Click()
    FillTextBoxes
End Sub

Private Sub lvwMod_ItemClick(ByVal Item As MSComctlLib.ListItem)
    FillTextBoxes
End Sub

Private Sub MakeInfo()
    On Error Resume Next
    Dim strFile As String
    Dim hPID As Long
    Dim hVer As VERHEADER
    Dim hIcoExt As Long, hIcoDraw As Long
    picIco.Cls
    strFile = frmMain.lvwProcessExplorer.SelectedItem.SubItems(1)
    GetVerHeader strFile, hVer
    lblDescription = hVer.FileDescription
    lblCompany = hVer.CompanyName
    lblVersion = hVer.FileVersion
    lblCopyright = hVer.LegalCopyright
    lblFile = ": " & GetFileName(strFile)
    lblLocation = ": " & GetFilePath(strFile)
    lblType = ": " & GetPathType(strFile)
    lblSize = ": " & Format(GetSizeOfFile(strFile) / 1024, "###,###") & " KB"
    lblAttributes = ": " & GetAttribute(strFile)
    lblCreated = ": " & FormatDateTime(FileDateTime(strFile), vbLongDate) & _
        ", " & FormatDateTime(FileDateTime(strFile), vbLongTime)
    lblPID = ": " & frmMain.lvwProcessExplorer.SelectedItem.SubItems(5)
    lblThreads = ": " & frmMain.lvwProcessExplorer.SelectedItem.SubItems(6)
    lblMemory = ": " & frmMain.lvwProcessExplorer.SelectedItem.SubItems(7)
    lblPriority = ": " & frmMain.lvwProcessExplorer.SelectedItem.SubItems(8)
    hIcoExt = ExtractIcon(Me.hwnd, strFile, 0)
    If hIcoExt Then
        hIcoDraw = DrawIcon(picIco.hdc, 0, 0, hIcoExt)
    Else
        picIco.Picture = GetIco.Icon(strFile, LargeIcon)
    End If
    MakeVisible
End Sub

Private Sub MakeProgress()
    Screen.MousePointer = 13
    prgModule.Visible = True
    Dim i As Long
    i = 0
    Do
        DoEvents
        i = i + 2
        Sleep 1
        prgModule.Value = i
    Loop Until i = 100
    prgModule.Visible = False
    Screen.MousePointer = 0
End Sub

Private Sub MakeInvisible()
    lblDescription.Visible = False
    lblCompany.Visible = False
    lblVersion.Visible = False
    linMod.Visible = False
    Dim i As Integer
    For i = lblEmpty.LBound To lblEmpty.UBound
        lblEmpty.Item(i).Visible = False
    Next i
    lblFile.Visible = False
    lblLocation.Visible = False
    lblType.Visible = False
    lblSize.Visible = False
    lblAttributes.Visible = False
    lblCreated.Visible = False
    lblPID.Visible = False
    lblThreads.Visible = False
    lblMemory.Visible = False
    lblPriority.Visible = False
    lblCopyright.Visible = False
End Sub

Private Sub MakeVisible()
    lblDescription.Visible = True
    lblCompany.Visible = True
    lblVersion.Visible = True
    linMod.Visible = True
    Dim i As Integer
    For i = lblEmpty.LBound To lblEmpty.UBound
        lblEmpty.Item(i).Visible = True
    Next i
    lblFile.Visible = True
    lblLocation.Visible = True
    lblType.Visible = True
    lblSize.Visible = True
    lblAttributes.Visible = True
    lblCreated.Visible = True
    lblPID.Visible = True
    lblThreads.Visible = True
    lblMemory.Visible = True
    lblPriority.Visible = True
    lblCopyright.Visible = True
    picIco.Visible = True
End Sub
