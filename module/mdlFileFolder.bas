Attribute VB_Name = "mdlFileFolder"
 

Option Explicit
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hdcDest&, ByVal X&, ByVal Y&, ByVal flags&) As Long

Public Declare Sub Sleep Lib _
    "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SHGetSpecialFolderLocation Lib _
    "shell32.dll" (ByVal hWndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib _
    "shell32" (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib _
    "kernel32.dll" Alias "GetWindowsDirectoryA" ( _
    ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib _
    "kernel32.dll" Alias "GetSystemDirectoryA" ( _
    ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Private Declare Function SHRunDialog Lib _
    "shell32" Alias "#61" ( _
    ByVal hOwner As Long, _
    ByVal Unknown1 As Long, _
    ByVal Unknown2 As Long, _
    ByVal szTitle As String, _
    ByVal szPrompt As String, _
    ByVal uFlags As Long) As Long
Private Declare Function ShellExecuteEx Lib _
    "shell32" Alias "ShellExecuteExA" ( _
    SEI As SHELLEXECUTEINFO) As Long
Private Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Private Declare Function SHGetFileInfo Lib _
    "shell32.dll" Alias "SHGetFileInfoA" ( _
    ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbFileInfo As Long, _
    ByVal uFlags As Long) As Long
    
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
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

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Enum SpecialFolder
    CSIDL_RECENT = &H8
    CSIDL_PROFILER = &H28
    CSIDL_HISTORY = &H22
End Enum

Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_EDITBOX As Long = &H10
Private Const MAX_PATH As Integer = 260
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_EXPLORER = &H80000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_HIDEREADONLY = &H4
Private Const SHGFI_DISPLAYNAME As Long = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Declare Sub CoTaskMemFree Lib _
    "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib _
    "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, _
    ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib _
    "shell32" (lpBI As BrowseInfo) As Long

Private SIconInfo As SHFILEINFO
Public Sub GetIcon(icPath$, pDisp As PictureBox)
pDisp.Cls
Dim hImgSmall&: hImgSmall = SHGetFileInfo(icPath$, 0&, SIconInfo, Len(SIconInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
 ImageList_Draw hImgSmall, SIconInfo.iIcon, pDisp.hDC, 0, 0, ILD_TRANSPARENT
End Sub
Public Sub GetLargeIcon(icPath$, pDisp As PictureBox)
Dim hImgLrg&: hImgLrg = SHGetFileInfo(icPath$, 0&, SIconInfo, Len(SIconInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
 ImageList_Draw hImgLrg, SIconInfo.iIcon, pDisp.hDC, 0, 0, ILD_TRANSPARENT
End Sub
Public Function BrowseForFolder(ByVal lnghWnd As Long, _
    ByVal strPrompt As String) As String
    On Error GoTo ehBrowseForFolder
    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As BrowseInfo
    With udtBI
        .lnghWnd = lnghWnd
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
ehBrowseForFolder:
    BrowseForFolder = Empty
End Function

Public Function GetSpecialFolder(FolderType As SpecialFolder) As String
    Dim r As Long, sPath As String
    Dim IDL As ITEMIDLIST
    r = SHGetSpecialFolderLocation(100, FolderType, IDL)
    sPath = Space$(512)
    r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    GetSpecialFolder = Left$(sPath, InStr(1, sPath, Chr$(0)) - 1)
End Function

Public Function GetWindowsPath() As String
    Dim lpBuffer As String * 255
    Dim nSize As Long
    nSize = GetWindowsDirectory(lpBuffer, 255)
    GetWindowsPath = Left(lpBuffer, nSize) & "\"
End Function

Public Function GetSystem32Path() As String
    Dim lpBuffer As String * 255
    Dim nSize As Long
    nSize = GetSystemDirectory(lpBuffer, 255)
    GetSystem32Path = Left(lpBuffer, nSize) & "\"
End Function

Public Function OpenInFolder(lvwItemExe As ListView, ItemId As Integer) As Double
    On Error Resume Next
    OpenInFolder = Shell("explorer.exe /select, " & _
        lvwItemExe.SelectedItem.SubItems(ItemId), vbNormalFocus)
End Function

Public Function OpenDosPrompt(lvwFilePath As ListView, _
    ItemExepath As Integer) As Long
    On Error Resume Next
    OpenDosPrompt = ShellExecute(1, vbNullString, "command.com", _
        vbNullString, GetFilePath(lvwFilePath.SelectedItem.SubItems(ItemExepath)), 1)
End Function

Public Function ShowRunApp(ByVal hWnd As Long) As Long
    On Error Resume Next
    ShowRunApp = SHRunDialog(hWnd, 0, 0, _
        StrConv("New Process", vbUnicode), _
        StrConv("Type the name of a program, folder, document, or Internet Resource," _
        & "and Windows will open it for you.", vbUnicode), 0)
End Function

Public Function ShowFileProperties(hWndOwner As Long, _
    lvwFilePath As ListView, ItemExepath As Integer, _
    Optional lUseSubItem As Boolean = True) _
     As Long
    On Error Resume Next
    Dim SEI As SHELLEXECUTEINFO
    Dim slpFileName As String
    If lUseSubItem Then
        slpFileName = lvwFilePath.SelectedItem.SubItems(ItemExepath)
    Else
        slpFileName = lvwFilePath.SelectedItem
    End If
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
            SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = hWndOwner
        .lpVerb = "properties"
        .lpFile = slpFileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 1
        .lpIDList = 0
    End With
    Call ShellExecuteEx(SEI)
End Function

Public Function GetFilePath(sPath As String) As String
    Dim i As Integer
    For i = Len(sPath) To 1 Step -1
        If Mid$(sPath, i, 1) = "\" Then
            GetFilePath = Mid$(sPath, 1, i)
            Exit For
        End If
    Next i
End Function

Public Function GetPathType(ByVal Path As String) As String
    Dim FileInfo As SHFILEINFO, lngRet As Long
    lngRet = SHGetFileInfo(Path, 0, FileInfo, _
        Len(FileInfo), SHGFI_DISPLAYNAME Or SHGFI_TYPENAME)
    If lngRet = 0 Then GetPathType = _
        Trim$(GetFileExtension(Path) & " File"): Exit Function
    GetPathType = Left$(FileInfo.szTypeName, _
        InStr(1, FileInfo.szTypeName, vbNullChar) - 1)
End Function

Public Function GetFileExtension(ByVal Path As String) As String
    Dim intRet As Integer: intRet = InStrRev(Path, ".")
    If intRet = 0 Then Exit Function
    GetFileExtension = UCase(Mid$(Path, intRet + 1))
End Function
