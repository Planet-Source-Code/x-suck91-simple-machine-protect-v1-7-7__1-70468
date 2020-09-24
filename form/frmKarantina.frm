VERSION 5.00
Begin VB.Form frmKarantina 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Karantina"
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   Picture         =   "frmKarantina.frx":0000
   ScaleHeight     =   2730
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrOnTop 
      Interval        =   1
      Left            =   5160
      Top             =   360
   End
   Begin SimpleMachineProtect.chameleonButton cmdhide 
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Hide"
      ENAB            =   -1  'True
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
      MICON           =   "frmKarantina.frx":360C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.FileListBox flsVirus 
      Height          =   1650
      Left            =   360
      Pattern         =   "*.smp"
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   4440
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblQuarantined 
      BackStyle       =   0  'Transparent
      Caption         =   "                  No File's In Quarantined"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   4485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "                  List of quarantine file"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   5985
   End
End
Attribute VB_Name = "frmKarantina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Added By X-Suck91

Private Sub cmdhide_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
AlwaysOnTop Me.hwnd, True
flsVirus.path = App.path & "\karantina\"
If flsVirus.List(1) = "" Then
lblQuarantined.Visible = True
Else
lblQuarantined.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tmrOnTop.Enabled = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
MoveForm Me.hwnd
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
MoveForm Me.hwnd
End If
End Sub

Private Sub Label2_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
MoveForm Me.hwnd
End If
End Sub

Private Sub lblQuarantined_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
MoveForm Me.hwnd
End If
End Sub

Private Sub tmrOnTop_Timer()
On Error Resume Next
'AlwaysOnTop Me.hwnd, True
End Sub
