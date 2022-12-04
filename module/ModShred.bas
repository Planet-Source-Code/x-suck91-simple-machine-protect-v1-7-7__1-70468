Attribute VB_Name = "ModShred"
'Added By X-Suck91


Declare Function FlushFileBuffers Lib "kernel32" (ByVal hfile As Long) As Long
Dim SBox(255) As Long
Dim KeyArr(255) As Long
Dim I1 As Long, J1 As Long

Public Sub RC4Seed(key As String)
Dim lngSwapVar As Long
Dim i As Long
Dim j As Long

For i = 0 To 255
    KeyArr(i) = Asc(Mid(key, (i Mod Len(key)) + 1, 1))
    SBox(i) = i
Next

j = 0
For i = 0 To 255
    j = (j + SBox(i) + KeyArr(i)) Mod 256
    lngSwapVar = SBox(i)
    SBox(i) = SBox(j)
    SBox(j) = lngSwapVar
Next
End Sub

Public Function RC4Rnd() As Double

Dim a As Long
Dim k As Long
Dim tmpVar As Long

I1 = (I1 + 1) Mod 256
J1 = (J1 + SBox(I1)) Mod 256
tmpVar = SBox(I1)
SBox(I1) = SBox(J1)
SBox(J1) = tmpVar

k = SBox((SBox(I1) + SBox(J1)) Mod 256)

RC4Rnd = CDbl(k) / CDbl(255)
End Function

Public Function ReadBinary(FileName As String)

Dim BinContent As String
Dim FileNo As Integer, FileSize As Long

FileNo = FreeFile
FileSize = FileLen(FileName)
BinContent = String(FileSize, " ")

Open FileName For Binary As FileNo
Get FileNo, , BinContent
Close FileNo

ReadBinary = BinContent
End Function

Public Function WriteBinary(FileName As String, BinContent As String)

Dim FileNo As Integer, FileSize As Long
FileNo = FreeFile

Open FileName For Binary As FileNo
FlushFileBuffers (FileNo)
Put FileNo, , BinContent
FlushFileBuffers (FileNo)
Close FileNo

End Function

Public Function ShredFile(FileName As String, Strength As Integer)
On Error Resume Next

Dim FileContentLen As Long
FileContentLen = FileLen(FileName)

Dim FName As String
Dim TmpF As String
FName = FileName

Dim FileNo As Integer
Dim i As Long, j As Long
Dim BinString0 As String
Dim BinString255 As String
BinString0 = String(FileContentLen, Chr(0))
BinString255 = String(FileContentLen, Chr(255))


For i = 1 To Strength
    WriteBinary FName, BinString0
    WriteBinary FName, BinString255
Next


Dim Seed As String
Seed = CStr(Date) & CStr(Timer) & CStr(Time) & Chr(CLng(Rnd * 255))
RC4Seed Seed


For i = 1 To Strength

    WriteBinary FName, RandBinary(FileContentLen)

    DoEvents
Next


For i = 1 To Strength
    TmpF = RandFileName()
    Name FName As TmpF
    FName = TmpF
Next


Kill FName
End Function

Function RandFileName() As String
Dim RFN As String
Dim RndNum As Double
Dim Chars As String
Chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890-_ "

For i = 0 To 100
    RndNum = RC4Rnd() * Len(Chars)
    RFN = RFN & Mid(Chars, 1 + CInt(RndNum), 1)
Next

RandFileName = RFN & ".tmp"
End Function

Function RandBinary(Length As Long) As String
Dim i As Long
Dim ByteArray() As Byte
Dim RandChar As String
Dim RC4Rand As Double

ReDim Preserve ByteArray(Length) As Byte

For i = 0 To Length
    RC4Rand = RC4Rnd() * 255
    ByteArray(i) = CByte(RC4Rand)
Next

RandBinary = StrConv(ByteArray, vbUnicode)
End Function


