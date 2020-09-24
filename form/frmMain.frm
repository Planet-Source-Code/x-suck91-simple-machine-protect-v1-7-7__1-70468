VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Main"
   ClientHeight    =   7515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":058A
   ScaleHeight     =   7515
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   143
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   142
      Top             =   7920
      Width           =   1455
   End
   Begin VB.ComboBox cboStrength 
      Height          =   315
      Left            =   2400
      TabIndex        =   135
      Text            =   "28"
      Top             =   8400
      Visible         =   0   'False
      Width           =   975
   End
   Begin SimpleMachineProtect.chameleonButton cmdHide 
      Height          =   495
      Left            =   360
      TabIndex        =   57
      ToolTipText     =   "Minimize to system tray"
      Top             =   5520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Hide"
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
      MICON           =   "frmMain.frx":3C60
      PICN            =   "frmMain.frx":3C7C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   2055
      TabIndex        =   28
      Top             =   6120
      Width           =   2055
      Begin SimpleMachineProtect.chameleonButton cmdExit 
         Height          =   495
         Left            =   0
         TabIndex        =   130
         ToolTipText     =   "Exit Simple Machine Protect"
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Exit"
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
         MICON           =   "frmMain.frx":4216
         PICN            =   "frmMain.frx":4232
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   1
      Left            =   360
      ScaleHeight     =   1095
      ScaleWidth      =   2055
      TabIndex        =   27
      Top             =   3960
      Width           =   2055
      Begin SimpleMachineProtect.chameleonButton cmdGeneral 
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   33
         ToolTipText     =   "About"
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "About"
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
         MICON           =   "frmMain.frx":47CC
         PICN            =   "frmMain.frx":47E8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdHelp 
         Height          =   495
         Left            =   0
         TabIndex        =   128
         ToolTipText     =   "Simple Machine Protect on the internet"
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Quick Help"
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
         MICON           =   "frmMain.frx":4D82
         PICN            =   "frmMain.frx":4D9E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   2
      Left            =   240
      TabIndex        =   26
      Top             =   5280
      Width           =   2295
      Begin VB.PictureBox picMenu 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   3
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   1995
         TabIndex        =   129
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Help"
      Height          =   1455
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "General"
      Height          =   3135
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   480
      Width           =   2295
      Begin VB.PictureBox picMenu 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2775
         Index           =   0
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   2055
         TabIndex        =   25
         Top             =   240
         Width           =   2055
         Begin SimpleMachineProtect.chameleonButton cmdGeneral 
            Height          =   615
            Index           =   0
            Left            =   0
            TabIndex        =   29
            ToolTipText     =   "Virus Scanner"
            Top             =   0
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            BTYPE           =   3
            TX              =   "SMP Virus Scanner"
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
            MICON           =   "frmMain.frx":5338
            PICN            =   "frmMain.frx":5354
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SimpleMachineProtect.chameleonButton cmdGeneral 
            Height          =   615
            Index           =   1
            Left            =   0
            TabIndex        =   30
            ToolTipText     =   "Control Administer"
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            BTYPE           =   3
            TX              =   "SMP Control Administer"
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
            MICON           =   "frmMain.frx":5C2E
            PICN            =   "frmMain.frx":5C4A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SimpleMachineProtect.chameleonButton cmdGeneral 
            Height          =   615
            Index           =   2
            Left            =   0
            TabIndex        =   31
            ToolTipText     =   "Process Explorer"
            Top             =   1440
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            BTYPE           =   3
            TX              =   "SMP Process Explorer"
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
            MICON           =   "frmMain.frx":6524
            PICN            =   "frmMain.frx":6540
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SimpleMachineProtect.chameleonButton cmdGeneral 
            Height          =   615
            Index           =   3
            Left            =   0
            TabIndex        =   32
            ToolTipText     =   "System Optimizer"
            Top             =   2160
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            BTYPE           =   3
            TX              =   "SMP System Optimizer"
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
            MICON           =   "frmMain.frx":6E1A
            PICN            =   "frmMain.frx":6E36
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   -1  'True
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
   End
   Begin VB.PictureBox picGeneral 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Index           =   4
      Left            =   2640
      Picture         =   "frmMain.frx":7710
      ScaleHeight     =   6495
      ScaleWidth      =   8205
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   8205
      Begin VB.PictureBox picHelpAbout 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   3495
         TabIndex        =   132
         Top             =   3600
         Width           =   3495
         Begin VB.Label lblHelpAbout 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "https://sourceforge.net/projects/smpav"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   1
            Left            =   240
            MousePointer    =   10  'Up Arrow
            TabIndex        =   133
            ToolTipText     =   "http://sourceforge.net/projects/smpav"
            Top             =   0
            Width           =   2925
         End
      End
      Begin VB.PictureBox picHelpAbout 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   2295
         TabIndex        =   131
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disclaimer"
         Enabled         =   0   'False
         Height          =   855
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   5400
         Width           =   7815
         Begin VB.Label lblEmpty 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   $"frmMain.frx":AB18
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   555
            Index           =   22
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   7515
         End
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Thanks"
         Enabled         =   0   'False
         Height          =   975
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   4320
         Width           =   7815
         Begin VB.Label lblEmpty 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   $"frmMain.frx":ABCF
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   675
            Index           =   21
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   7515
         End
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "History"
         Enabled         =   0   'False
         Height          =   2895
         Index           =   1
         Left            =   4920
         TabIndex        =   6
         Top             =   1320
         Width           =   3135
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   13
            Left            =   360
            Picture         =   "frmMain.frx":ACF7
            Top             =   2280
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   12
            Left            =   360
            Picture         =   "frmMain.frx":B281
            Top             =   1920
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   11
            Left            =   360
            Picture         =   "frmMain.frx":B80B
            Top             =   1560
            Width           =   240
         End
         Begin VB.Line linSMP 
            BorderColor     =   &H00C0C0C0&
            Index           =   3
            X1              =   480
            X2              =   960
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line linSMP 
            BorderColor     =   &H00C0C0C0&
            Index           =   2
            X1              =   480
            X2              =   960
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line linSMP 
            BorderColor     =   &H00C0C0C0&
            Index           =   1
            X1              =   480
            X2              =   480
            Y1              =   720
            Y2              =   1200
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   10
            Left            =   1080
            Picture         =   "frmMain.frx":BD95
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   9
            Left            =   1080
            Picture         =   "frmMain.frx":C31F
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   8
            Left            =   360
            Picture         =   "frmMain.frx":C8A9
            Top             =   360
            Width           =   240
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Minor bugs fixed."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   20
            Left            =   720
            TabIndex        =   53
            Top             =   2280
            Width           =   1245
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Online Help ready."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   19
            Left            =   720
            TabIndex        =   52
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Improved user interface."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   18
            Left            =   720
            TabIndex        =   51
            Top             =   1560
            Width           =   1800
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Updated."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   17
            Left            =   1440
            TabIndex        =   50
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Improved engine."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   16
            Left            =   1440
            TabIndex        =   49
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Improved performance."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   15
            Left            =   840
            TabIndex        =   48
            Top             =   360
            Width           =   1710
         End
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Simple Machine Protect"
         Enabled         =   0   'False
         Height          =   3495
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   4455
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X-Suck91 of SMAN 1 Tarogong Kidul"
            Height          =   195
            Left            =   720
            TabIndex        =   146
            Top             =   3120
            Width           =   2985
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   ": http://xsuck91.xm.com"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   1680
            TabIndex        =   141
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Homepage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   480
            TabIndex        =   140
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   ": Angga Sanusi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   1680
            TabIndex        =   139
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Modified"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   480
            TabIndex        =   138
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label lblHelpAbout 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   ": http://www.e-freshware.com"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   0
            Left            =   1680
            MousePointer    =   10  'Up Arrow
            TabIndex        =   137
            ToolTipText     =   "http://www.e-freshware.com"
            Top             =   2160
            Width           =   2235
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   ": Erwin Rusadi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   12
            Left            =   1680
            TabIndex        =   47
            Top             =   1920
            Width           =   1020
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   ": Bagus Judistirah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   11
            Left            =   1680
            TabIndex        =   46
            Top             =   1680
            Width           =   1275
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Homepage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   10
            Left            =   480
            TabIndex        =   45
            Top             =   2160
            Width           =   765
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Virus Definition"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   9
            Left            =   480
            TabIndex        =   44
            Top             =   1920
            Width           =   1065
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Developer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   8
            Left            =   480
            TabIndex        =   43
            Top             =   1680
            Width           =   735
         End
         Begin VB.Line linAbout 
            BorderColor     =   &H00C0C0C0&
            X1              =   480
            X2              =   3960
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   ": Microsoft Visual Basic 6.0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   7
            Left            =   1560
            TabIndex        =   42
            Top             =   1200
            Width           =   1905
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   ": GNU General Public License"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   6
            Left            =   1560
            TabIndex        =   41
            Top             =   960
            Width           =   2040
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   ": April 29, 2008"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   5
            Left            =   1560
            TabIndex        =   40
            Top             =   720
            Width           =   1110
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   ": 1.7.7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   4
            Left            =   1560
            TabIndex        =   39
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Compiler"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   38
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "License"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   37
            Top             =   960
            Width           =   525
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Release"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   36
            Top             =   720
            Width           =   570
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Version"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   35
            Top             =   480
            Width           =   525
         End
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X-Suck 91 of SMAN 1 Tarogong Kidul"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2880
         TabIndex        =   145
         Top             =   120
         Width           =   3030
      End
      Begin VB.Label lblEmpty 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Open Source Project"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   13
         Left            =   3345
         TabIndex        =   134
         Top             =   6240
         Width           =   1515
      End
      Begin VB.Image imgSMP 
         Height          =   480
         Index           =   0
         Left            =   6000
         Picture         =   "frmMain.frx":CE33
         Top             =   720
         Width           =   1980
      End
   End
   Begin VB.PictureBox picGeneral 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Index           =   3
      Left            =   2640
      Picture         =   "frmMain.frx":11947
      ScaleHeight     =   6495
      ScaleWidth      =   8205
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   8205
      Begin SimpleMachineProtect.chameleonButton cmdJumpTo 
         Height          =   375
         Index           =   10
         Left            =   5520
         TabIndex        =   72
         Top             =   5760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Windows Explorer"
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
         MICON           =   "frmMain.frx":163A8
         PICN            =   "frmMain.frx":163C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdJumpTo 
         Height          =   375
         Index           =   9
         Left            =   5520
         TabIndex        =   71
         Top             =   5280
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Task Manager"
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
         MICON           =   "frmMain.frx":1695E
         PICN            =   "frmMain.frx":1697A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdJumpTo 
         Height          =   375
         Index           =   8
         Left            =   5520
         TabIndex        =   70
         Top             =   4800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "System Restore"
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
         MICON           =   "frmMain.frx":16F14
         PICN            =   "frmMain.frx":16F30
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdJumpTo 
         Height          =   375
         Index           =   7
         Left            =   5520
         TabIndex        =   69
         Top             =   4320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "System Info"
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
         MICON           =   "frmMain.frx":174CA
         PICN            =   "frmMain.frx":174E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdJumpTo 
         Height          =   375
         Index           =   6
         Left            =   5520
         TabIndex        =   68
         Top             =   3840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "System Config"
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
         MICON           =   "frmMain.frx":17A80
         PICN            =   "frmMain.frx":17A9C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdJumpTo 
         Height          =   375
         Index           =   5
         Left            =   5520
         TabIndex        =   67
         Top             =   3360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Security Center"
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
         MICON           =   "frmMain.frx":18036
         PICN            =   "frmMain.frx":18052
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdJumpTo 
         Height          =   375
         Index           =   4
         Left            =   5520
         TabIndex        =   66
         Top             =   2880
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Registry Editor"
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
         MICON           =   "frmMain.frx":185EC
         PICN            =   "frmMain.frx":18608
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdJumpTo 
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   65
         Top             =   2400
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Disk Defragmenter"
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
         MICON           =   "frmMain.frx":18BA2
         PICN            =   "frmMain.frx":18BBE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdJumpTo 
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   64
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Control Panel"
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
         MICON           =   "frmMain.frx":19158
         PICN            =   "frmMain.frx":19174
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdJumpTo 
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   63
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Console Window"
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
         MICON           =   "frmMain.frx":1970E
         PICN            =   "frmMain.frx":1972A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdJumpTo 
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   62
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Clean Manager"
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
         MICON           =   "frmMain.frx":19CC4
         PICN            =   "frmMain.frx":19CE0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdWasher 
         Height          =   375
         Index           =   2
         Left            =   3600
         TabIndex        =   61
         Top             =   5760
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Junk File"
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
         MICON           =   "frmMain.frx":1A27A
         PICN            =   "frmMain.frx":1A296
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdWasher 
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   60
         Top             =   5760
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Recycle Bin"
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
         MICON           =   "frmMain.frx":1A830
         PICN            =   "frmMain.frx":1A84C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdWasher 
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   59
         Top             =   5760
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Registry"
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
         MICON           =   "frmMain.frx":1ADE6
         PICN            =   "frmMain.frx":1AE02
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvwSystemOptimizer 
         Height          =   4335
         Left            =   480
         TabIndex        =   58
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilsGlobal"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "System Optimizer Available For Windows"
            Object.Width           =   9260
         EndProperty
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jump To"
         Enabled         =   0   'False
         Height          =   5535
         Index           =   6
         Left            =   5400
         TabIndex        =   11
         Top             =   720
         Width           =   2535
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Washer"
         Enabled         =   0   'False
         Height          =   735
         Index           =   5
         Left            =   360
         TabIndex        =   10
         Top             =   5520
         Width           =   4815
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Windows Optimizer"
         Enabled         =   0   'False
         Height          =   4695
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   4815
      End
   End
   Begin VB.PictureBox picGeneral 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Index           =   2
      Left            =   2640
      Picture         =   "frmMain.frx":1B39C
      ScaleHeight     =   6495
      ScaleWidth      =   8205
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   8205
      Begin SimpleMachineProtect.chameleonButton cmdEndProc 
         Height          =   375
         Left            =   6240
         TabIndex        =   77
         Top             =   3720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "End Process"
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
         MICON           =   "frmMain.frx":1FB90
         PICN            =   "frmMain.frx":1FBAC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdNewProc 
         Height          =   375
         Left            =   4320
         TabIndex        =   76
         Top             =   3720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "New Process"
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
         MICON           =   "frmMain.frx":20146
         PICN            =   "frmMain.frx":20162
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdRefresh 
         Height          =   375
         Left            =   2280
         TabIndex        =   75
         Top             =   3720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Refresh"
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
         MICON           =   "frmMain.frx":206FC
         PICN            =   "frmMain.frx":20718
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdLocate 
         Height          =   375
         Left            =   360
         TabIndex        =   74
         Top             =   3720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Locate"
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
         MICON           =   "frmMain.frx":20CB2
         PICN            =   "frmMain.frx":20CCE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvwProcessExplorer 
         Height          =   2775
         Left            =   360
         TabIndex        =   73
         Top             =   840
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilsProcessExplorer"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Image Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Location"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Attributes"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Process ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Threads"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Memory Usage"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Priority"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Memory Information"
         Enabled         =   0   'False
         Height          =   1935
         Index           =   8
         Left            =   240
         TabIndex        =   13
         Top             =   4320
         Width           =   7815
         Begin VB.Label lblMem 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   7
            Left            =   7650
            TabIndex        =   93
            Top             =   1440
            Width           =   45
         End
         Begin VB.Label lblMem 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   6
            Left            =   7650
            TabIndex        =   92
            Top             =   1080
            Width           =   45
         End
         Begin VB.Label lblMem 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   5
            Left            =   7650
            TabIndex        =   91
            Top             =   720
            Width           =   45
         End
         Begin VB.Label lblMem 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   4
            Left            =   7650
            TabIndex        =   90
            Top             =   360
            Width           =   45
         End
         Begin VB.Label lblMem 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   3
            Left            =   3690
            TabIndex        =   89
            Top             =   1440
            Width           =   45
         End
         Begin VB.Label lblMem 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   3690
            TabIndex        =   88
            Top             =   1080
            Width           =   45
         End
         Begin VB.Label lblMem 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   1
            Left            =   3690
            TabIndex        =   87
            Top             =   720
            Width           =   45
         End
         Begin VB.Label lblMem 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   3690
            TabIndex        =   86
            Top             =   360
            Width           =   45
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "CPU Usage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   30
            Left            =   4080
            TabIndex        =   85
            Top             =   1440
            Width           =   795
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Memory Usage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   29
            Left            =   4080
            TabIndex        =   84
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Available Page File"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   28
            Left            =   4080
            TabIndex        =   83
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Total Page File"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   27
            Left            =   4080
            TabIndex        =   82
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Available Virtual Memory"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   26
            Left            =   240
            TabIndex        =   81
            Top             =   1440
            Width           =   1755
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Total Virtual Memory"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   80
            Top             =   1080
            Width           =   1470
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Available Physical Memory"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   24
            Left            =   240
            TabIndex        =   79
            Top             =   720
            Width           =   1875
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Total Physical Memory"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   23
            Left            =   240
            TabIndex        =   78
            Top             =   360
            Width           =   1590
         End
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Process List"
         Enabled         =   0   'False
         Height          =   3615
         Index           =   7
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   7815
      End
      Begin MSComctlLib.ImageList ilsProcessExplorer 
         Left            =   7320
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Timer tmrMemory 
         Interval        =   1000
         Left            =   7560
         Top             =   840
      End
   End
   Begin VB.PictureBox picGeneral 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Index           =   1
      Left            =   2640
      Picture         =   "frmMain.frx":21268
      ScaleHeight     =   6495
      ScaleWidth      =   8205
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   8205
      Begin VB.ListBox lstVirList 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   2010
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   110
         Top             =   4200
         Width           =   3375
      End
      Begin VB.CheckBox chkSound 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Warning Sound"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   109
         Top             =   3480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkScanMem 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scan Memory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   108
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkHidden 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hidden Recovery"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   107
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CheckBox chkRep 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Repair Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   106
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkFixReg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fix Error Registry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   105
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CheckBox chkTrans 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transparent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   104
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CheckBox chkHideTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hide Window Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   103
         Top             =   4680
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkOnTop 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Always On Top"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   102
         Top             =   4200
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.PictureBox picReport 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   480
         ScaleHeight     =   735
         ScaleWidth      =   3375
         TabIndex        =   98
         Top             =   5400
         Width           =   3375
         Begin VB.OptionButton optOffReport 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Turn Off Reporting Service"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   480
            Width           =   2415
         End
         Begin VB.OptionButton optFullReport 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Complete Reporting Service"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   2415
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Use Default Reporting Service"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   0
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin VB.PictureBox picOpt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   480
         ScaleHeight     =   975
         ScaleWidth      =   3375
         TabIndex        =   94
         Top             =   960
         Width           =   3375
         Begin VB.ComboBox cboExt 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmMain.frx":25CBB
            Left            =   360
            List            =   "frmMain.frx":25CD7
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   600
            Width           =   2900
         End
         Begin VB.OptionButton optExt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Use Extension List"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   2775
         End
         Begin VB.OptionButton optAllFiles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All Files Extension"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   0
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Entries Of The Detection List"
         Enabled         =   0   'False
         Height          =   2295
         Index           =   14
         Left            =   4200
         TabIndex        =   19
         Top             =   3960
         Width           =   3615
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporting Service"
         Enabled         =   0   'False
         Height          =   1095
         Index           =   12
         Left            =   360
         TabIndex        =   17
         Top             =   5160
         Width           =   3615
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Window Settings"
         Enabled         =   0   'False
         Height          =   1095
         Index           =   11
         Left            =   360
         TabIndex        =   16
         Top             =   3960
         Width           =   3615
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scan Options"
         Enabled         =   0   'False
         Height          =   1575
         Index           =   10
         Left            =   360
         TabIndex        =   15
         Top             =   2280
         Width           =   3615
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "File Extensions To Scan"
         Enabled         =   0   'False
         Height          =   1335
         Index           =   9
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   3615
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version Information"
         Enabled         =   0   'False
         Height          =   3135
         Index           =   13
         Left            =   4200
         TabIndex        =   18
         Top             =   720
         Width           =   3735
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Real Time Protector : 1.0.1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Left            =   1680
            TabIndex        =   147
            Top             =   1560
            Width           =   1950
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   720
            X2              =   1200
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "System Optimizer 1.0.5"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   36
            Left            =   1200
            TabIndex        =   116
            Top             =   2760
            Width           =   1680
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Process Explorer 2.1.1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   35
            Left            =   1200
            TabIndex        =   115
            Top             =   2400
            Width           =   1635
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Control Administer 1.0.8"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   34
            Left            =   1200
            TabIndex        =   114
            Top             =   2040
            Width           =   1755
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Signature: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   33
            Left            =   1680
            TabIndex        =   113
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Engine 1.3.7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   32
            Left            =   1680
            TabIndex        =   112
            Top             =   840
            Width           =   915
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Virus Scanner 1.2.5"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   31
            Left            =   1080
            TabIndex        =   111
            Top             =   480
            Width           =   1410
         End
         Begin VB.Line linSMP 
            BorderColor     =   &H00C0C0C0&
            Index           =   0
            X1              =   720
            X2              =   720
            Y1              =   840
            Y2              =   1680
         End
         Begin VB.Line linSMP 
            BorderColor     =   &H00C0C0C0&
            Index           =   5
            X1              =   720
            X2              =   1200
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line linSMP 
            BorderColor     =   &H00C0C0C0&
            Index           =   7
            X1              =   720
            X2              =   1200
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   2
            Left            =   600
            Picture         =   "frmMain.frx":25DBC
            Top             =   480
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   4
            Left            =   1320
            Picture         =   "frmMain.frx":26346
            Top             =   1200
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   6
            Left            =   720
            Picture         =   "frmMain.frx":268D0
            Top             =   2040
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   5
            Left            =   720
            Picture         =   "frmMain.frx":26E5A
            Top             =   2760
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   7
            Left            =   720
            Picture         =   "frmMain.frx":273E4
            Top             =   2400
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   3
            Left            =   1320
            Picture         =   "frmMain.frx":2796E
            Top             =   840
            Width           =   240
         End
      End
      Begin MSComctlLib.ImageList ilsGlobal 
         Left            =   7200
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":27EF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":28492
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":28A2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":28FC6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picGeneral 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Index           =   0
      Left            =   2640
      Picture         =   "frmMain.frx":29560
      ScaleHeight     =   6495
      ScaleWidth      =   8205
      TabIndex        =   0
      Top             =   560
      Width           =   8205
      Begin SimpleMachineProtect.chameleonButton cmdShred 
         Height          =   375
         Left            =   6600
         TabIndex        =   144
         Top             =   4440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Shred Virus"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14869218
         BCOLO           =   14869218
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":2DCE7
         PICN            =   "frmMain.frx":2DD03
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvwVirFound 
         Height          =   2655
         Left            =   360
         TabIndex        =   124
         Top             =   1680
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilsGlobal"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Virus Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Location"
            Object.Width           =   9834
         EndProperty
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   122
         Top             =   840
         Width           =   5655
      End
      Begin SimpleMachineProtect.chameleonButton cmdDelete 
         Height          =   375
         Left            =   5040
         TabIndex        =   121
         Top             =   4440
         Width           =   1335
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Delete"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":2E29D
         PICN            =   "frmMain.frx":2E2B9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdReport 
         Height          =   375
         Left            =   3480
         TabIndex        =   120
         Top             =   4440
         Width           =   1335
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Report"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":2E853
         PICN            =   "frmMain.frx":2E86F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdStop 
         Height          =   375
         Left            =   1920
         TabIndex        =   119
         Top             =   4440
         Width           =   1335
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Stop"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":2EE09
         PICN            =   "frmMain.frx":2EE25
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdScan 
         Height          =   375
         Left            =   360
         TabIndex        =   118
         Top             =   4440
         Width           =   1335
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Scan"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":2F3BF
         PICN            =   "frmMain.frx":2F3DB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SimpleMachineProtect.chameleonButton cmdBrowse 
         Height          =   375
         Left            =   6120
         TabIndex        =   117
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Browse"
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
         MICON           =   "frmMain.frx":2F975
         PICN            =   "frmMain.frx":2F991
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Current Progress"
         Enabled         =   0   'False
         Height          =   1215
         Index           =   17
         Left            =   240
         TabIndex        =   22
         Top             =   5040
         Width           =   7815
         Begin SimpleMachineProtect.XP_ProgressBar prgScan 
            Height          =   300
            Left            =   120
            TabIndex        =   123
            Top             =   550
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   12632319
            Orientation     =   1
            Scrolling       =   1
         End
         Begin VB.Shape shpScan 
            BorderColor     =   &H00C0C0C0&
            Height          =   255
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   7575
         End
         Begin VB.Label lblAnim 
            BackColor       =   &H00FFFFFF&
            Caption         =   "[-]"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   135
            Left            =   120
            TabIndex        =   127
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Waiting For Instruction"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Left            =   3075
            TabIndex        =   126
            Top             =   885
            Width           =   1665
         End
         Begin VB.Label lblScan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   225
            Left            =   240
            TabIndex        =   125
            Top             =   255
            Width           =   7335
         End
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Virus Found"
         Enabled         =   0   'False
         Height          =   3495
         Index           =   16
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   7815
      End
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Location To Scan"
         Enabled         =   0   'False
         Height          =   735
         Index           =   15
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   7815
      End
      Begin VB.Timer tmrScan 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7560
         Top             =   5160
      End
   End
   Begin SimpleMachineProtect.OnSystray OnSystray 
      Left            =   360
      Top             =   5520
      _ExtentX        =   476
      _ExtentY        =   476
      Icon            =   "frmMain.frx":2FF2B
      ToolTipText     =   "Simple Machine Protect"
   End
   Begin VB.Line linSMP 
      BorderColor     =   &H00C0C0C0&
      Index           =   6
      X1              =   2880
      X2              =   3360
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line linSMP 
      BorderColor     =   &H00C0C0C0&
      Index           =   4
      X1              =   7560
      X2              =   8040
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   11
      Left            =   6480
      Picture         =   "frmMain.frx":304C5
      Top             =   8880
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   10
      Left            =   5880
      Picture         =   "frmMain.frx":3118F
      Top             =   8880
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   9
      Left            =   5280
      Picture         =   "frmMain.frx":31E59
      Top             =   8880
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   8
      Left            =   4680
      Picture         =   "frmMain.frx":32B23
      Top             =   8880
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   7
      Left            =   6480
      Picture         =   "frmMain.frx":337ED
      Top             =   8280
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   6
      Left            =   5880
      Picture         =   "frmMain.frx":344B7
      Top             =   8280
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   5280
      Picture         =   "frmMain.frx":35181
      Top             =   8280
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   4680
      Picture         =   "frmMain.frx":35E4B
      Top             =   8280
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   6480
      Picture         =   "frmMain.frx":36B15
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Left            =   5880
      Picture         =   "frmMain.frx":377DF
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   5280
      Picture         =   "frmMain.frx":384A9
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   4680
      Picture         =   "frmMain.frx":39173
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   495
      Index           =   12
      Left            =   4800
      Top             =   9480
      Width           =   495
   End
   Begin VB.Label lblStrength 
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
      Height          =   255
      Left            =   3360
      TabIndex        =   136
      Top             =   8400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblCopyLeft 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright © 2007 - 2008 BJ`s Software Studios && X-Suck91 Community. All Rights Reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   56
      Top             =   7155
      Width           =   6735
   End
   Begin VB.Image imgSMP 
      Height          =   480
      Index           =   1
      Left            =   360
      Picture         =   "frmMain.frx":39E3D
      Top             =   6840
      Width           =   1980
   End
   Begin VB.Label lblMove 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SIMPLE MACHINE PROTECT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   10815
   End
End
Attribute VB_Name = "frmMain"
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
' History    : Minor bugs fixed.
'            : Penambahan Modul Shred File
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
'              * UDARAMAYA [http://www.udaramaya.com]
'              * X-Suck91 Community [http://xsuck91.xm.com]
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

Private Sub chkHideTitle_Click()
    If chkHideTitle.Value = vbChecked Then
        App.TaskVisible = False
        App.Title = GenerateRandomTitle(True)
    Else
        App.TaskVisible = True
        App.Title = GenerateMainTitle
    End If
End Sub

Private Sub chkOnTop_Click()
    If chkOnTop.Value = 1 Then
        AlwaysOnTop Me.hwnd, True
    Else
        AlwaysOnTop Me.hwnd, False
    End If
End Sub

Private Sub chkTrans_Click()
    If chkTrans.Value = 1 Then
        SetOpagueForm False
    Else
        SetOpagueForm True
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim BFF As String
    BFF = BrowseForFolder(Me.hwnd, _
        "Select Drive Or Directory: ")
    If Len(BFF) > 0 Then
        txtLocation.Text = BFF
    End If
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    Dim lValue As Long
    lValue = CheckVirusItem
    If lValue > 0 Then
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
    End If
    If MsgBox("Are you sure you want to delete selected file(s)?", _
        vbYesNo + vbQuestion, "SMP Virus Scanner") = vbYes Then
        lblStatus.Caption = "Deleting File": DoEvents
        cmdDelete.Enabled = False
        DeleteVirus
    Else
        cmdDelete.Enabled = True
    End If
    With lvwVirFound
        If .ListItems.Count > 0 Then
            .Enabled = True
        Else
            .Enabled = False
            .Refresh
        End If
    End With
    lblStatus.Caption = ""
End Sub

Private Sub cmdExit_Click()
    If MsgBox("Are you sure you want to exit?", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation") = vbYes Then
        Form_Terminate
    End If
End Sub

Private Sub cmdGeneral_Click(Index As Integer)
    tmrMemory.Enabled = False
    frmInfo.prgInfo.Color = &H8080&
    Dim i As Integer
    With picGeneral
        For i = .LBound To .UBound
            .Item(i).Visible = False
            cmdGeneral.Item(i).UseGreyscale = True
        Next i
        .Item(Index).Visible = True
        cmdGeneral.Item(Index).UseGreyscale = False
        If Index = 2 Then
            frmInfo.Caption = "Getting Information"
            frmInfo.prgInfo.Color = &HC56A31
            lvwProcessExplorer.Visible = False
            frmInfo.Show vbModal
            tmrMemory.Enabled = True
            Call cmdRefresh_Click
        End If
    End With
End Sub

Private Sub cmdHelp_Click()
    PopupMenu frmInfo.mnuHelp
End Sub

Private Sub cmdhide_Click()
    Me.Hide
    OnSystray.Visible = True
End Sub

Private Sub cmdJumpTo_Click(Index As Integer)
    frmInfo.Caption = cmdJumpTo(Index).Caption
    frmInfo.Show vbModal
    With Me
        Select Case Index
            Case 0: OpenXPTool .hwnd, "cleanmgr.exe"
            Case 1: OpenXPTool .hwnd, "cmd.exe"
            Case 2: OpenXPTool .hwnd, "control.exe"
            Case 3: OpenXPTool .hwnd, "dfrg.msc"
            Case 4: OpenXPTool .hwnd, "regedit.exe"
            Case 5: OpenXPTool .hwnd, "wscui.cpl"
            Case 6: OpenXPTool .hwnd, "msconfig.exe"
            Case 7: OpenXPTool .hwnd, "winmsd.exe"
            Case 8: OpenXPTool .hwnd, GetSystem32Path & "restore\rstrui.exe"
            Case 9: OpenXPTool .hwnd, "taskmgr.exe"
            Case 10: OpenXPTool .hwnd, "explorer.exe"
        End Select
    End With
End Sub

Private Sub cmdLocate_Click()
    OpenInFolder lvwProcessExplorer, 1
    cmdRefresh_Click
End Sub

Private Sub cmdNewProc_Click()
    ShowRunApp Me.hwnd
    cmdRefresh_Click
End Sub

Private Sub cmdRefresh_Click()
    NTProcessList lvwProcessExplorer, ilsProcessExplorer
End Sub

Private Sub cmdReport_Click()
    frmReport.Show vbModal
End Sub

Private Sub cmdScan_Click()
    On Error Resume Next
    picGeneral(0).MousePointer = vbHourglass
    cmdhide.SetFocus
    cmdBrowse.Enabled = False
    cmdScan.Enabled = False
    cmdReport.Enabled = False
    cmdDelete.Enabled = False
    cmdShred.Enabled = False
    With lvwVirFound
        .ListItems.Clear
        .Checkboxes = False
        .ColumnHeaders(1).Width = 2000
        .ColumnHeaders(2).Width = 5575
        .Enabled = False
    End With
    Dim i As Integer
    For i = picMenu.LBound To picMenu.UBound
        picMenu.Item(i).Enabled = False
    Next i
    StopScan = False
    nMemory = 0
    If chkScanMem.Value = 1 Then
        lblStatus.Caption = "Scanning Memory": DoEvents
        ScanProcess True
    End If
    Sleep 250
    If chkFixReg.Value = 1 Then
        lblStatus.Caption = "Repairing Registry": DoEvents
        Clean_Registry
    End If
    Sleep 250
    With prgScan
        .Value = 0
        .ShowText = True
        .Orientation = ccOrientationHorizontal
        .Scrolling = ccScrollingSearch
    End With
    tmrScan.Enabled = True
    lblStatus.Caption = "Preparing To Scan": lblScan.Caption = _
        "Please wait and do not open any application.": DoEvents
    MakeStartReporting
    CalcFileNow
    tmrScan.Enabled = False
    Sleep 500
    With prgScan
        .Value = 0
        .ShowText = True
        .Scrolling = ccScrollingSmooth
    End With
    cmdStop.Enabled = True
    lblStatus.Caption = "Scanning File": DoEvents
    ScanVirus txtLocation.Text, frmMain
    
    cmdScan.Enabled = True
    cmdStop.Enabled = False
    If chkHidden.Value = 1 Then
        lblAnim.Visible = True
        lblStatus.Caption = "Applying Attributes": DoEvents
        HiddenRecovery txtLocation.Text, lblAnim
        lblAnim.Visible = False
    End If
    If optOffReport.Value = True Then
        cmdReport.Enabled = False
    Else
        cmdReport.Enabled = True
    End If
    MakeFinishReporting
    prgScan.Value = 100
    lblStatus.Caption = "Scan Progress Finished"
    lblScan.Caption = ""
    Sleep 500
    frmFinish.Show vbModal, Me
    Me.Show
    If optOffReport.Value = False Then
        CreateLogFile App.path & "\SMP.LOG", frmReport.txtLog.Text
    End If
    With lvwVirFound
        If lvwVirFound.ListItems.Count > 0 Then
            LV_AutoSizeColumn lvwVirFound, .ColumnHeaders.Item(1)
            LV_AutoSizeColumn lvwVirFound, .ColumnHeaders.Item(2)
            .Enabled = True
            .Checkboxes = True
            For i = 1 To .ListItems.Count
                .ListItems(i).Checked = True
            Next i
            cmdDelete.Enabled = True
            cmdShred.Enabled = True
        Else
            .ColumnHeaders(1).Width = 2000
            .ColumnHeaders(2).Width = 5575
            .Enabled = False
            cmdDelete.Enabled = False
            cmdShred.Enabled = False
        End If
    End With
    For i = picMenu.LBound To picMenu.UBound
        picMenu.Item(i).Enabled = True
    Next i
    picGeneral(0).MousePointer = vbDefault
    cmdGeneral(0).SetFocus
    cmdBrowse.Enabled = True
End Sub



Private Sub cmdShred_Click()
On Error Resume Next
    Dim lValue As Long
    lValue = CheckVirusItem
    If lValue > 0 Then
        cmdShred.Enabled = False
    Else
        cmdShred.Enabled = True
    End If
    If MsgBox("This will shred the selected file(s)?", _
        vbYesNo + vbQuestion, "SMP Virus Scanner") = vbYes Then
                        If MsgBox("This will take a very long time to shred the selected file", _
        vbYesNo + vbQuestion, "Delete virus permanently") = vbYes Then
                lblStatus.Caption = "Shreding File": DoEvents
        cmdShred.Enabled = False
        cmdDelete.Enabled = False
        ShredVirus
        MsgBox "Virus has been delete permanently", vbInformation + vbSystemModal, "Shred Virus"
    Else
        cmdShred.Enabled = True
        End If
    End If
    With lvwVirFound
        If .ListItems.Count > 0 Then
            .Enabled = True
        Else
            .Enabled = False
            .Refresh
            cmdDelete.Enabled = False
            cmdShred.Enabled = False
            MsgBox "Virus has been delete permanently", vbInformation + vbSystemModal, "Shred Virus"
        End If
        
    End With
    
    lblStatus.Caption = ""
End Sub

Private Sub cmdStop_Click()
    If MsgBox("Are you sure you want to stop scan progress?", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation") = vbYes Then
        StopScan = True
    End If
End Sub

Private Sub cmdWasher_Click(Index As Integer)
    frmInfo.Caption = cmdWasher(Index).Caption
    frmInfo.Show vbModal
    Select Case Index
        Case 0
            Call Clean_Registry
        Case 1
            Call EmptyTheRecycle(Me.hwnd)
        Case 2
            Call ClearJunkFile
    End Select
End Sub


Private Sub Form_Activate()

    
    Unload frmLoading
    lblEmpty(33).Caption = "Signature: " & lstVirList.ListCount
    AlwaysOnTop Me.hwnd, True
    RegisterFile ".SMP", "This File Is Quarantined By Simple Machine Protect", "Simple Machine Protect", App.path & "\" & "SMP.exe /Q %1", App.path & "\secicon.ico"
IconCompare = True
Akurat = "100"
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    GetCPUInfo lblMem(7)
    cboExt.ListIndex = 4
    Me.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    chkOnTop_Click
End Sub

Private Sub Form_Terminate()
    ExecuteOptimizer lvwSystemOptimizer
    'SaveAppSettings
    ExitNow
End Sub

Private Sub lblHelpAbout_Click(Index As Integer)
    If Index = 0 Then
        OnlineHelp Me.hwnd, "www.e-freshware.com"
    Else
        OnlineHelp Me.hwnd, "souceforge.net/projects/smpav"
    End If
End Sub

Private Sub lblMove_MouseDown(Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        MoveForm Me.hwnd
    End If
End Sub

Private Sub lvwProcessExplorer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lPID As Long
    lPID = lvwProcessExplorer.SelectedItem.SubItems(5)
    If GetAppID <> lPID Then
        cmdEndProc.Enabled = True
    Else
        cmdEndProc.Enabled = False
    End If
End Sub

Private Sub lvwProcessExplorer_MouseDown(Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu frmInfo.mnuProcess
    End If
End Sub

Private Sub lvwVirFound_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If CheckVirusItem > 0 Then
        cmdDelete.Enabled = True
        cmdShred.Enabled = True
    Else
        cmdDelete.Enabled = False
        cmdShred.Enabled = False
    End If
End Sub



Private Sub OnSystray_DoubleClick()
    On Error Resume Next
    Me.Show
    Unload frmFinish
    OnSystray.Visible = False
End Sub

Private Sub OnSystray_MouseUp(Button As Integer)
    If Button = 2 Then
        PopupMenu frmInfo.mnuSystray
    End If
End Sub

Private Sub optAllFiles_Click()
    cboExt.Enabled = False
End Sub

Private Sub optExt_Click()
    cboExt.Enabled = True
End Sub



Private Sub tmrMemory_Timer()
    MemoryInfo lblMem(5), lblMem(1), lblMem(3), _
        lblMem(4), lblMem(0), lblMem(2), lblMem(6)
    UpdateValues lblMem(7)
End Sub

Private Sub cmdEndProc_Click()
    Dim lExitCode As Long
    If MsgBox("Are you sure you want to terminate the process?", _
        vbYesNo + vbExclamation, "SMP Process Explorer Warning") = vbYes Then
        lExitCode = TerminateProcessID(lvwProcessExplorer, 5)
        If lExitCode = 0 Then MsgBox "SMP Process Explorer cannot end this process.", _
            vbExclamation, "Unable To Terminate Process"
        Call cmdRefresh_Click
    End If
End Sub

Private Sub tmrScan_Timer()
    If prgScan.Value < 100 Then
        prgScan.Value = prgScan.Value + 1
    Else
        prgScan.Value = 0
    End If
End Sub

Private Sub txtLocation_Change()
    If Mid$(txtLocation.Text, 2, 2) <> ":\" Then
        cmdScan.Enabled = False
    Else
        cmdScan.Enabled = True
    End If
End Sub

Private Sub MakeStartReporting()
    With frmReport.txtLog
        VirLog = ""
        .Text = ""
        .Text = .Text & "SIMPLE MACHINE PROTECT" & " - Report File Date: " & _
            FormatDateTime(Now, vbLongDate) & vbCrLf & vbCrLf
        If optFullReport.Value = True Then
            .Text = .Text & "    Version" & vbTab & vbTab & vbTab & ": 1.6.8" & vbCrLf
            .Text = .Text & "    Engine" & vbTab & vbTab & vbTab & ": 1.3.3" & vbCrLf
            .Text = .Text & "    Signature" & vbTab & vbTab & vbTab & ": " & _
                lstVirList.ListCount & vbCrLf
            .Text = .Text & vbCrLf & "    Start Of The Scan" & vbTab & vbTab & _
                ": " & Time & vbCrLf
            .Text = .Text & "    Location Of The Scan" & vbTab & ": " & _
                txtLocation.Text & vbCrLf
            .Text = .Text & "    File Extensions To Scan" & vbTab & _
                CheckFileScanValue(optAllFiles, cboExt) & vbCrLf & vbCrLf
            .Text = .Text & "    Fix Error Registry" & vbTab & vbTab & _
                CheckBoxesValues(chkFixReg) & vbCrLf
            .Text = .Text & "    Repair Data" & vbTab & vbTab & vbTab & _
                CheckBoxesValues(chkRep) & vbCrLf
            .Text = .Text & "    Hidden Recovery" & vbTab & vbTab & _
                CheckBoxesValues(chkHidden) & vbCrLf
            .Text = .Text & "    Scan Memory" & vbTab & vbTab & vbTab & _
                CheckBoxesValues(chkScanMem) & vbCrLf
            .Text = .Text & "    Warning Sound" & vbTab & vbTab & _
                CheckBoxesValues(chkSound) & vbCrLf & vbCrLf
        Else
            .Text = .Text & "    Start Of The Scan" & vbTab & vbTab & ": " & _
                Time & vbCrLf
            .Text = .Text & "    Location Of The Scan" & vbTab & ": " & _
                txtLocation.Text & vbCrLf
            .Text = .Text & "    File Extensions To Scan" & vbTab & _
                CheckFileScanValue(optAllFiles, cboExt) & vbCrLf & vbCrLf
        End If
    End With
End Sub
Private Sub MakeFinishReporting()
    Dim sSaveTime As String
    With frmReport.txtLog
        If chkScanMem.Value = vbChecked Then
            Dim sStrMemScan As String
            If nMemory > 1 Then
                sStrMemScan = ": " & nMemory & " Processes"
            Else
                sStrMemScan = ": " & nMemory & " Process"
            End If
            .Text = .Text & "    Memory Scanned" & vbTab & vbTab & sStrMemScan & vbCrLf
        End If
        .Text = .Text & "    File Scanned" & vbTab & vbTab & _
            CheckValueData(nFile, "scanned") & vbCrLf
        .Text = .Text & "    File Infected" & vbTab & vbTab & _
            CheckValueData(nInfect, "infected") & vbCrLf
        .Text = .Text & "    File Repaired" & vbTab & vbTab & _
            CheckValueData(nRepair, "repaired") & vbCrLf
        .Text = .Text & "    File Detected" & vbTab & vbTab & _
            CheckValueData(lvwVirFound.ListItems.Count, "detected") & vbCrLf
        sSaveTime = Time
        sLastScan = CStr(FormatDateTime(Now, vbLongDate) & " - " & sSaveTime)
        .Text = .Text & "    End Of The Scan" & vbTab & vbTab & ": " & _
            sSaveTime & vbCrLf & vbCrLf
        If StopScan = False Then
            .Text = .Text & "    The scan has been done completely."
        Else
            .Text = .Text & "    The scan was not completed succesfully."
        End If
        If optFullReport.Value = True Then
            .Text = .Text & vbCrLf & vbCrLf & VirLog
        End If
    End With
End Sub

Private Function CheckVirusItem() As Long
    Dim i As Double
    With lvwVirFound
        For i = 1 To .ListItems.Count
            If .ListItems.Item(i).Checked = True Then
                CheckVirusItem = CheckVirusItem + 1
            End If
        Next i
    End With
End Function

Public Sub DeleteVirus()
    On Error Resume Next
    Dim sDelete As String
    Dim i As Long, lRet As Long
    With lvwVirFound
        For i = 1 To .ListItems.Count
            If .ListItems.Item(i).Checked = True Then
                sDelete = .ListItems.Item(i).SubItems(1)
                lRet = KillVirusNow(sDelete)
                If lRet <> 0 Then
                    .ListItems.Remove i
                Else
                    .ListItems(i).Checked = False
                End If
                .ListItems.Remove i
                Call DeleteVirus
                Exit For
            End If
        Next i
    End With
End Sub
Public Sub ShredVirus()
    On Error Resume Next
    Dim sDelete As String
    Dim i As Long, lRet As Long
    With lvwVirFound
        For i = 1 To .ListItems.Count
            If .ListItems.Item(i).Checked = True Then
                sDelete = .ListItems.Item(i).SubItems(1)
                lRet = ShredVirusNow(sDelete)
                If lRet <> 0 Then
                    .ListItems.Remove i
                Else
                    .ListItems(i).Checked = False
                End If
                .ListItems.Remove i
                Call ShredVirus
                Exit For
            End If
        Next i
    End With
End Sub
Private Sub lvwVirFound_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
With lvwVirFound
'For i = 1 To .ListItems.Count
If Button = 1 Then
Text1.Text = .SelectedItem.SubItems(1)
Text2.Text = .SelectedItem
'Exit For
End If
'Next i
If Button = 2 Then
Text1.Text = .SelectedItem.SubItems(1)
Text2.Text = .SelectedItem
PopupMenu frmFinish.mnuFile
End If
End With
End Sub
Private Sub SetOpagueForm(lMode As Boolean)
    Me.Enabled = False
    Dim i As Integer
    Select Case lMode
        Case False
            i = 255
            Do
                i = i - 5
                DoEvents
                MakeTransparent Me.hwnd, i
                Sleep 10
            Loop While i >= 192
        Case True
            i = 192
            Do
                i = i + 5
                DoEvents
                MakeTransparent Me.hwnd, i
                Sleep 10
            Loop Until i >= 255
    End Select
    Me.Enabled = True
    Me.Refresh
End Sub
Public Sub Karantina()
On Error Resume Next
Dim sXor As New clsSimpleXOR
Dim i As String
X = Pengacakan(0, 999)
MkDir App.path & "\" & "karantina\"
sXor.EncryptFile Text1.Text, App.path & "\" & "karantina" & "\" & Text2.Text & X & ".smp", Text2.Text
Set sXor = Nothing
lvwVirFound.SelectedItem.ListSubItems.Remove 1
i = lvwVirFound.SelectedItem
lvwVirFound.ListItems.Remove (i)
MsgBox i & " Virus Has Been Added To Karantina Folder", vbInformation + vbSystemModal, "Secure Your Virus In Karantina Folder"

End Sub
Public Function Pengacakan(ByVal Low As Long, ByVal High As Long) As Long
Randomize
Pengacakan = Int((High - Low + 1) * Rnd) + Low
End Function

