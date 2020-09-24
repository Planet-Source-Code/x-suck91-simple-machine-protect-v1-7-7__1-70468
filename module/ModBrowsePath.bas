Attribute VB_Name = "ModBrowsePath"
'---------------------------------------------------------------------------------------'
'Module     : SMP Virus Scanner - Source Code Builder (Browse Path)                     '
'Author     : Bagus Judistirah (bagus_badboy)                                           '
'Company    : BJ's Software Studio                                                      '
'E-Mail     : bagus.judistirah@hotmail.com                                              '
'Homepage   : http://www.cyber-freak.net                                                '
'License    : GNU General Public License                                                '
'Study      : Accounting Department - State Polytechnic Of Malang                       '
'History    : Some bugs fixed.                                                          '
'Note       : I try to keep my software as bug-free as possible.                        '
'             But it's a general rule that no software ever is error free,              '
'             and the number of errors increases with the complexity of the program.    '
'Thanks     : Aat Shadewa, Boby Ertanto, Irwan Halim, Dony Wahyu ISP, Yusuf TP,         '
'             Erwin, MI People, Husni, BillyInferno, Paul, Marx, VM, Wardana,           '
'             Achmad Darmal, My Sweety, Andy, Dream Theater, Evanescence, Virologi,     '
'             OGnizer - Online Virus Community, VB-BEGO, all of you who's support me.   '
'---------------------------------------------------------------------------------------'

Option Explicit

Private Type BrowseInfo
    lnghWnd As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_EDITBOX As Long = &H10
Private Const MAX_PATH As Integer = 260
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Declare Function ShellExecuteEx Lib _
    "shell32" Alias "ShellExecuteExA" _
    (SEI As SHELLEXECUTEINFO) As Long
Private Declare Sub CoTaskMemFree Lib _
    "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib _
    "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, _
    ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib _
    "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib _
    "shell32" (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long
Private Declare Function GetFileTitle Lib _
    "COMDLG32.DLL" Alias "GetFileTitleA" _
    (ByVal lpszFile As String, _
    ByVal lpszTitle As String, _
    ByVal cbBuf As Integer) As Integer
Private Declare Function GetSystemDirectory _
    Lib "kernel32.dll" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib _
    "kernel32.dll" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long

Enum Win32Path
    Windows
    System32
End Enum

Private Function BrowseForFolder(ByVal hWndOwner As Long, _
    ByVal strPrompt As String) As String
    
    On Error GoTo ErrHandle
    
    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As BrowseInfo
    
    With udtBI
        .lnghWnd = hWndOwner
        .lpszTitle = lstrcat(strPrompt, "")
        .ulFlags = BIF_NEWDIALOGSTYLE + BIF_EDITBOX
    End With
    
    lngIDList = SHBrowseForFolder(udtBI)
    
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH, 0)
        lngResult = SHGetPathFromIDList(lngIDList, _
            strPath)
        Call CoTaskMemFree(lngIDList)
        intNull = InStr(strPath, vbNullChar)
            If intNull > 0 Then
                strPath = Left(strPath, intNull - 1)
            End If
    End If
     
    BrowseForFolder = strPath
    
    Exit Function
    
ErrHandle:
    BrowseForFolder = Empty
    
End Function

Public Function LocateFile(lvwFilePath As ListView, _
    ItemExepath As Integer) As Double

    On Error Resume Next
    
    Shell "explorer.exe " & GetFilePath(lvwFilePath. _
        SelectedItem.SubItems(ItemExepath)), _
        vbNormalFocus

End Function

Private Function GetFilePath(ByVal sPath As String) _
    As String
    
    Dim i As Integer
    
    For i = Len(sPath) To 1 Step -1
        If Mid$(sPath, i, 1) = "\" Then
            GetFilePath = Mid$(sPath, 1, i)
            Exit For
        End If
    Next i

End Function

Public Function ShowFileProperties1(hWndOwner As Long, _
     lvwFilePath As ListView, ItemExepath As Integer) _
     As Long

    Dim SEI As SHELLEXECUTEINFO
    
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
            SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = hWndOwner
        .lpVerb = "properties"
        .lpFile = lvwFilePath.SelectedItem.SubItems _
            (ItemExepath)
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 1
        .lpIDList = 0
    End With

    Call ShellExecuteEx(SEI)

End Function

Public Function StripNulls1(sStr As String) As String

    StripNulls1 = Left$(sStr, InStr(1, sStr, Chr$(0)) - 1)
    
End Function

Public Function GetFileName1(sFilename As String) As String

    Dim buffer As String
    
    buffer = String(255, 0)
    
    GetFileTitle sFilename, buffer, Len(buffer)
    
    buffer = StripNulls1(buffer)
    GetFileName1 = buffer
    
End Function

Public Function WinSysDirPath(ByVal WinOrSys As Win32Path) _
    As String

    Dim buffer As String * 255
    Dim WinSys As Long

    Select Case WinOrSys
        Case Win32Path.Windows
            WinSys = GetWindowsDirectory(buffer, 255)
        Case Win32Path.System32
            WinSys = GetSystemDirectory(buffer, 255)
    End Select
    
    WinSysDirPath = Left(buffer, WinSys) '& "\"
    
End Function


