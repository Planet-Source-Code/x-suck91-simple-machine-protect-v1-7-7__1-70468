VERSION 5.00
Begin VB.Form frmRTP 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Real Time Protector"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3600
      Top             =   2520
   End
   Begin VB.Timer tmrScan 
      Interval        =   1000
      Left            =   120
      Top             =   480
   End
   Begin SimpleMachineProtect.chameleonButton cmdHapus 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Hapus"
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
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   65280
      FCOLO           =   65280
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRTP.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Timer tmrGetAddExplo 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblLokasi 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   2280
      TabIndex        =   6
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Lokasi Viri   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Viri Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Virus Terdeteksi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   45
   End
End
Attribute VB_Name = "frmRTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Added By X-Suck91
'Simple Machine Protect - Real Time Protector
'Sebenarnya code ini adalah potongan kode yang ada pada
'virus aksika dan saya menggunakannya untuk menjadi kode RTP ini



Private Declare Function SendMessage Lib _
"user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib _
"user32" Alias "FindWindowA" _
(ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib _
"user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Const WM_GETTEXT = &HD
Public Aktif As Boolean

Private Sub cmdHapus_Click()
On Error Resume Next
KillVirusNow (lblLokasi.Caption)
Me.Visible = False
Label1.Caption = ""
lblLokasi.Caption = ""
tmrGetAddExplo.Enabled = True
tmrScan.Enabled = True

End Sub

Private Sub Form_Load()
AlwaysOnTop Me.hwnd, True
tmrGetAddExplo.Enabled = True
tmrScan.Enabled = True
frmInfo.popSystray(10).Caption = "Tutup Real Time Protector"
End Sub



Private Sub Form_Unload(Cancel As Integer)
frmInfo.popSystray(10).Caption = "Real Time Protector"
Unload frmScanFD
tmrGetAddExplo.Enabled = False
tmrScan.Enabled = False
Timer1.Enabled = False

Unload Me
End Sub

Private Sub Timer1_Timer()
ScanFlashDisk
End Sub

Private Sub tmrGetAddExplo_Timer()
On Error Resume Next
Dim hand1 As Long
Dim hand2 As Long
Dim hand3 As Long
Dim hand4 As Long
Dim hand5 As Long
Dim hand6 As Long
Dim hand7 As Long
Dim temp As String * 256
Dim temp2 As String * 256
Dim AlamatFile1 As String
Dim JudulCaption1 As String


Dim hand10 As Long

        
   
    'membaca address bar pada windows explorer sebagai media penyebaran
    hand1 = FindWindow("ExploreWClass", vbNullString)
    hand10 = FindWindow("CabinetWClass", vbNullString)
    If hand1 = GetForegroundWindow Then
        hand2 = FindWindowEx(hand1, 0&, "WorkerW", vbNullString)
        SendMessage hand1, WM_GETTEXT, 200, ByVal temp2
    ElseIf hand10 = GetForegroundWindow Then
        hand2 = FindWindowEx(hand10, 0&, "WorkerW", vbNullString)
        SendMessage hand10, WM_GETTEXT, 200, ByVal temp2
    
    End If
    'dapatkan string pada address bar
    hand3 = FindWindowEx(hand2, 0&, "RebarWindow32", vbNullString)
    hand4 = FindWindowEx(hand3, 0&, "ComboBoxEx32", vbNullString)
    hand5 = FindWindowEx(hand4, 0&, "ComboBox", vbNullString)
    hand6 = FindWindowEx(hand5, 0&, "Edit", vbNullString)
    SendMessage hand6, WM_GETTEXT, 200, ByVal temp
    

    'ambil lokasi folder yang aktif pada windows explorer
    AlamatFile1 = Mid$(temp, 1, InStr(temp, Chr$(0)) - 1)
    'ambil caption windows explorer
    JudulCaption1 = Mid$(temp2, 1, InStr(temp2, Chr$(0)) - 1)
    Text1.Text = AlamatFile1
 
End Sub

Private Sub tmrScan_Timer()
On Error Resume Next

   If Text1.Text = "" Then
    
    Else
    Hitung
    RealTimeProtector Text1.Text, frmMain

    End If
    
    
End Sub
Private Sub ScanFlashDisk()
On Error GoTo Batal
Dim AdaFlashDisk As Boolean
Dim ObjFSO As Object
Dim ObjDrive As Object
'buat file scripting object
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
AdaFlashDisk = False
For Each ObjDrive In ObjFSO.Drives
    'Asumsi semua removable drive diatas huruf C adalah flash disk
    '1 - Removable drive
    '2 - Fixed drive (hard disk)
    '3 - Mapped network drive
    '4 - CD-ROM drive
    '5 - RAM disk
    If ObjDrive.DriveType = 1 Then
       AdaFlashDisk = True
    End If
    If AdaFlashDisk = True Then
    frmScanFD.Text1.Text = ObjDrive.Driveletter + ":\"
    Load frmScanFD
    End If
    Next

Batal:
End Sub
