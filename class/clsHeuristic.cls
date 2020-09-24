VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHeuristic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Saya menggunakan modul heuristic dari Peradnya Anti Virus
'Mohon ijin nya ya APTX demi pengembangan SMP yang lebih baik



Dim i As Integer
Dim j As Integer
Public IconCompare As Boolean
Public TipeHeuristic As String
Public Akurat As Integer
Public Function CekHeuristic(Alamatfile As String, Formku As Form) As Boolean
    On Error GoTo keluar
    Dim tmp As String
    If Right(Alamatfile, 3) = "exe" And IconCompare = True Then
        If CekIcon(Alamatfile, Formku.Picture1, Formku.Picture2, Akurat) = True Then CekHeuristic = True
    Else
 
    Open Alamatfile For Input As 1
        Do While Not EOF(1)
            Input #1, tmp
            Select Case Right(Alamatfile, 3)
                Case "vbs": If CekVBS(tmp) = True Then CekHeuristic = True
                Case "bat": If CekBat(tmp) = True Then CekHeuristic = True
                Case "inf": If CekInf(tmp) = True Then CekHeuristic = True
            End Select
        Loop
    Close #1
   End If
    Exit Function
keluar:
    CekHeuristic = False
End Function
Private Function CekVBS(baris As String) As Boolean
    Dim bhy(15) As String
    bhy(0) = "scripting.filesystemobject"
    bhy(1) = "wscript.shell"
    bhy(2) = "wscript.scriptfullname"
    bhy(3) = "createobject"
    bhy(4) = "regsetvalue"
    bhy(5) = "copyfile"
    bhy(6) = "exitwindowsex"
    bhy(7) = "persistmoniker=file:"
    bhy(8) = "runit"
    bhy(9) = "attachments.add"
    bhy(10) = "outlook.application"
    bhy(11) = "-f -r -t 0"
    bhy(12) = "HKEY_CLASSES_ROOT\tukul.ndeso\shell\open\command\"
    bhy(13) = "flashdrive"
    baris = LCase(baris)
    j = 0
    For i = 0 To 10
        j = InStr(1, baris, bhy(i), vbTextCompare)
        If j <> 0 Then
            CekVBS = True
            TipeHeuristic = "Heuristic.Susp.Vbs (" & i & ")"
            Exit Function
        End If
    Next
    CekVBS = False
End Function

'Untuk script Bat (file bat)
Private Function CekBat(baris As String) As Boolean
    Dim bhy(11) As String
    bhy(0) = "format "
    bhy(1) = "reg "
    bhy(2) = "%0"
    bhy(3) = "attrib "
    bhy(4) = "run"
    bhy(5) = "hidden"
    bhy(6) = "disable"
    bhy(7) = "startup"
    bhy(8) = "NoFolderOptions"
    bhy(9) = "HideFileExt"
    bhy(10) = "tskill "
    
    baris = LCase(baris)
    j = 0
    For i = 0 To 10
        j = InStr(1, baris, bhy(i), vbTextCompare)
        If j <> 0 Then
            CekBat = True
            TipeHeuristic = "Heuristic.Susp.Bat (" & i & ")"
            Exit Function
        End If
    Next
    CekBat = False
End Function
Private Function CekInf(baris As String) As Boolean
    Dim bhy(3) As String
    bhy(0) = "wscript"
    bhy(1) = ".vbs"
    bhy(2) = ".bat"
    bhy(3) = ".com"
    bhy(3) = ".scr"
    
    baris = LCase(baris)
    j = 0
    For i = 0 To 4
        j = InStr(1, baris, bhy(i), vbTextCompare)
        If j <> 0 Then
            CekInf = True
            TipeHeuristic = "Heuristic.Susp.Inf (" & i & ")"
            Exit Function
        End If
    Next
    CekInf = False
End Function
Private Sub PaP(p1 As PictureBox, p2 As PictureBox, W As Long, H As Long, c As Long, DDD As Byte)
    Dim X, Y, a, AP, DD1, DD2
    For X = 0 To W - 15 Step c  ' Wah baca koordinat (x,y) pixel nih...
        For Y = 0 To H - 15 Step c
            If p2.Point(X, Y) = p1.Point(X, Y) Then AP = AP + 1 ' Wah ketemu yang sama , tambah satu ah...
            a = a + 1  ' Pokoknya Tambah Satu (untuk total pixel)...
        Next
    Next
    DDD = (AP * 100) \ a 'Mencari Persentase ....
End Sub
Private Function CekIcon(PathFile As String, Picture1 As PictureBox, Picture2 As PictureBox, Akuratyuk As Integer) As Boolean
   
    CekIcon = False
    Picture1.Cls
    Picture2.Picture = frmMain.Image1(12).Picture
    GetLargeIcon PathFile, Picture1
    
    Dim i As Integer
    Dim re As Byte
    For i = 0 To 11 'Banyak Array
        Picture2.Cls
        Picture2.Picture = frmMain.Image1(i).Picture
        PaP Picture1, Picture2, Picture1.Width, Picture1.Height, 15, re
        'Utama.Text2(i) = re
        If re >= Akuratyuk Then
            'Factor = re
            TipeHeuristic = "Heuristic.Susp.Icon (" & i & ") [Acc = " & re & " %]"
            CekIcon = True
            Exit For
        End If
    Next
End Function
