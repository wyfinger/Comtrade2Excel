Attribute VB_Name = "Comtrade2Excel"
'
' Comtrade2Excel Excel2Comtrade Converter
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' https://github.com/wyfinger/Comtrade2Excel
' ����� �������, miv@prim.so-ups.ru
' 2014
'

Option Explicit

Private Declare PtrSafe Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare PtrSafe Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long

Dim objSheet As Variant


Private Function ReplaceExt(strFileName As String, strNewExt As String) As String
'
' ����� ���������� �����

Dim dot_pos As Long

dot_pos = InStrRev(strFileName, ".")
If dot_pos > 0 Then
  ReplaceExt = Left(strFileName, dot_pos) & strNewExt
Else
  ReplaceExt = "strFileName"
End If

End Function


Private Function ExtractFileName(strPath As String) As String
'
' �������� ��� �����

Dim slash_pos As Long

If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)

slash_pos = InStrRev(Replace$(strPath, "/", "\"), "\")
If slash_pos > 0 Then
  ExtractFileName = Right$(strPath, Len(strPath) - slash_pos)
Else
  ExtractFileName = ""
End If
  
End Function


Private Function FileExists(strFileName As String) As Boolean
'
' �������� ������������� �����

FileExists = PathFileExists(strFileName)
  
End Function


Private Function ArrGet(varArr, intNo As Integer) As String
'
' ������� ����������� ���������� ������ �� �������

If (intNo >= LBound(varArr)) And (intNo <= UBound(varArr)) Then
  ArrGet = varArr(intNo)
Else
  ArrGet = ""
End If

End Function


Private Function GetInt(strVal As String) As Double
'
' ������� �� ������ ��� ���������� ������� � ��������������� � �����

Dim i As Integer
Dim strRez As String

Const strDights = "0123456789."

If Len(strVal) = 0 Then
  GetInt = 0
Else
  For i = 1 To Len(strVal)
    If InStr(strDights, Mid(strVal, i, 1)) Then strRez = strRez & Mid(strVal, i, 1)
  Next
  If strRez = "" Then
    GetInt = 0
  Else
    GetInt = Val(strRez)
  End If
End If

End Function


Private Function ReadASCIILine(ByVal intFile As Integer)
'
' ������ ����� ������ ������ �� CFG ��� DAT ����� ASCII �������

Dim ReadedLine As String

Line Input #intFile, ReadedLine

' ���������� OEM (DOS) to ANSI (Windows)
Dim RVar As Integer
Dim DecodedLine As String

DecodedLine = ReadedLine
'RVar = OemToChar(ReadedLine, DecodedLine)

ReadASCIILine = Split(DecodedLine, ",")

End Function


Private Function Combine4Byte(ByVal A As Byte, ByVal B As Byte, ByVal C As Byte, ByVal D As Byte) As Double
'
' �� ������� ���� ������ Double (������ ��� ������������� �����)

Combine4Byte = D * 2 ^ 24 + C * 2 ^ 16 + B * 2 ^ 8 + A

End Function


Private Function Combine2Byte(ByVal A As Byte, ByVal B As Byte) As Integer
'
' �� ���� ���� ������ Double (����� ���� �������������)

Combine2Byte = (B And 127) * (2 ^ 8) + A
If (B And 128) = 128 Then Combine2Byte = Combine2Byte - (2 ^ 15)

End Function


Private Function ReadBINARYLine(ByVal intFile As Integer, ByVal intLen As Integer, ByVal intASig As Integer, ByVal intDSig As Integer)
'
' ������ ����� ������ ������ �� DAT ����� BINARY �������

' � BINARY DAT ����� ������ ��� ��������: ����� � ����� � �� �������� � 4 ����� ������, ����� ���� ����������
' ������� �� 2 ����� �� ������, � ����� �������� ��������, ���������� ���� ���������� ����� ����������� 1

Dim ByteLine() As Byte
ReDim ByteLine(intLen - 1) As Byte

Get #intFile, , ByteLine

Dim ResultStr As String

' ����� � ����� �������
ResultStr = Combine4Byte(ByteLine(0), ByteLine(1), ByteLine(2), ByteLine(3))
ResultStr = ResultStr & "," & Combine4Byte(ByteLine(4), ByteLine(5), ByteLine(6), ByteLine(7))

' ������ ���������� �������
Dim i As Integer
For i = 0 To intASig - 1
  ResultStr = ResultStr & "," & Combine2Byte(ByteLine(8 + i * 2), ByteLine(8 + i * 2 + 1))
Next

' ������ ���������� �������
For i = 0 To intDSig - 1
  If (ByteLine(8 + intASig * 2 + i \ 8) And (2 ^ (i Mod 8))) = (2 ^ (i Mod 8)) Then
    ResultStr = ResultStr & ",1"
  Else
    ResultStr = ResultStr & ",0"
  End If
Next

' ���������� ������
ReadBINARYLine = Split(ResultStr, ",")

End Function

Private Function ReadBINARYLineEasy(ByVal intFile As Integer, ByVal intLen As Integer, ByVal intASig As Integer, ByVal intDSig As Integer)
'
' ������ ����� ������ ������ �� DAT ����� BINARY EASY �������,
' � ���������� ��� ������ � ������� �������.

Dim ByteLine() As Byte
ReDim ByteLine(intLen - 1) As Byte

Get #intFile, , ByteLine

Dim ResultStr As String

' ������ ���������� �������
Dim i As Integer
For i = 0 To intASig - 1
  ResultStr = ResultStr & "," & Combine2Byte(ByteLine(i * 2), ByteLine(i * 2 + 1))
Next

' ������ ���������� �������
For i = 0 To intDSig - 1
  If (ByteLine(intASig * 2 + i \ 8) And (2 ^ (i Mod 8))) = (2 ^ (i Mod 8)) Then
    ResultStr = ResultStr & ",1"
  Else
    ResultStr = ResultStr & ",0"
  End If
Next

' ������� ������ �������
ResultStr = Mid(ResultStr, 2, Len(ResultStr))

' ���������� ������
ReadBINARYLineEasy = Split(ResultStr, ",")

End Function

Private Function LoadDataASCII(strDATFile As String) As Boolean
'
' ������ ����� ������ � ������� ASCII

Dim i As Integer
Dim j As Integer
Dim intDATFile As Integer
Dim Line

On Error GoTo err_read_dat_file
  Dim errReadDatFile As Boolean
  intDATFile = FreeFile
  Open strDATFile For Input Access Read Lock Write As #intDATFile
  Seek #intDATFile, 1

  i = 20
  Do While Not EOF(intDATFile)
    Line = ReadASCIILine(intDATFile)
    For j = 1 To UBound(Line) + 2
      objSheet.Cells(i, j + 1).Value = ArrGet(Line, j - 1)
    Next
    i = i + 1
  Loop

  Close #intDATFile
  errReadDatFile = True
  
err_read_dat_file:
  If Not errReadDatFile Then
    LoadDataASCII = False  ' Error, �������� ������ ��� ������ ASCII DAT �����
    Exit Function
  End If
On Error GoTo 0

LoadDataASCII = True

End Function


Private Function LoadDataBINARY(strDATFile As String, ByVal intASig As Integer, ByVal intDSig As Integer) As Boolean
'
' ������ ����� ������ � ������� BINARY

Dim intDATFile As Integer
Dim i As Integer
Dim j As Integer
Dim Line
Dim LineLen As Integer
Dim FileSize As Long

' ���������� ������ � ����� ������
LineLen = 8 + (intASig * 2) + (intDSig \ 8) + ((intDSig Mod 8) And 1)

On Error GoTo err_read_dat_file
  Dim errReadDatFile As Boolean
  intDATFile = FreeFile
  FileSize = FileLen(strDATFile)
  Open strDATFile For Binary Access Read Lock Write As #intDATFile Len = LineLen
  Seek #intDATFile, 1

  i = 20
  Do While (Not EOF(intDATFile)) And (Seek(intDATFile) < FileSize)
    Line = ReadBINARYLine(intDATFile, LineLen, intASig, intDSig)
    For j = 1 To UBound(Line) + 2
      objSheet.Cells(i, j + 1).Value = ArrGet(Line, j - 1)
    Next
    i = i + 1
  Loop

  Close #intDATFile
  errReadDatFile = True
  
err_read_dat_file:
  If Not errReadDatFile Then
    LoadDataBINARY = False  ' Error, �������� ������ ��� ������ DAT �����
    Exit Function
  End If
On Error GoTo 0

LoadDataBINARY = True

End Function


Private Function LoadDataBINARYEasy(strDATFile As String, ByVal intASig As Integer, ByVal intDSig As Integer) As Boolean
'
' ������ ����� ������ � ������� BINARY ���������� ��������,
' ��� � ����� CFG ����� �������� EASY= 1 � �� ����� ������ 8 ���� � DAT �����:
' ����� � ����� �������.

Dim intDATFile As Integer
Dim i As Integer
Dim j As Integer
Dim Line
Dim LineLen As Integer
Dim FileSize As Long
Dim stepPeriod As Long

' ���������� ������ � ����� ������, �� 8 ������, ��� � LoadDataBINARY
LineLen = (intASig * 2) + (intDSig \ 8) + ((intDSig Mod 8) And 1)

On Error GoTo err_read_dat_file
  Dim errReadDatFile As Boolean
  intDATFile = FreeFile
  FileSize = FileLen(strDATFile)
  Open strDATFile For Binary Access Read Lock Write As #intDATFile Len = LineLen
  Seek #intDATFile, 1

  stepPeriod = 1000000 / objSheet.Cells(5, 2).Value  ' ��� � ����� ����

  i = 20
  Do While (Not EOF(intDATFile)) And (Seek(intDATFile) < FileSize) And (i <= objSheet.Cells(6, 3).Value + 19)
    Line = ReadBINARYLineEasy(intDATFile, LineLen, intASig, intDSig)
    objSheet.Cells(i, 2).Value = i - 20 + 1     ' �����
    objSheet.Cells(i, 3).Value = (i - 20) * stepPeriod ' �����
    For j = 0 To UBound(Line)
      objSheet.Cells(i, j + 4).Value = ArrGet(Line, j)
    Next
    i = i + 1
  Loop

  Close #intDATFile
  errReadDatFile = True
  
err_read_dat_file:
  If Not errReadDatFile Then
    LoadDataBINARYEasy = False  ' Error, �������� ������ ��� ������ DAT �����
    Exit Function
  End If
On Error GoTo 0

LoadDataBINARYEasy = True

End Function

Private Function OpenComtrade(strFileName As String) As Integer
'
' ������ CFG ����� � �������� ������

' ������ CFG ����

Dim intCFGFile As Integer
Dim strDATFile As String

strDATFile = ReplaceExt(strFileName, "dat")

' �������� ������������� ������
If Not FileExists(strFileName) Then
  OpenComtrade = 1    ' Error, CFG ���� �� ������
  Exit Function
End If
If Not FileExists(strDATFile) Then
  OpenComtrade = 2    ' Error, DAT ���� �� ������
  Exit Function
End If

' ��������� CFG ����, ������������ ������ �������
On Error GoTo err_cfg_file_access
  Dim errCFGFileAccess As Boolean
  intCFGFile = FreeFile
  Open strFileName For Input Access Read Lock Write As #intCFGFile
  errCFGFileAccess = True
err_cfg_file_access:
  If Not errCFGFileAccess Then
    OpenComtrade = 3  ' Error, ������ ������� � CFG �����
    Exit Function
  End If
On Error GoTo 0
  
' ������� ����� ����
Set objSheet = ActiveWorkbook.Worksheets.Add
On Error GoTo err_sheet_exists
  Dim errSheetExists As Boolean
  objSheet.Name = ExtractFileName(strFileName)
  errSheetExists = True
err_sheet_exists:
  If Not errSheetExists Then
    ' ���� ���� ������������� �� ������� - ���� � ���
  
    'OpenComtrade = 4  ' Error, ���� � ����� ������ ��� ����������
    ' TO-DO: ������� �� ��� ������� ������ ��� ��������� ���� ��� �������
    'objSheet.Delete
    'Exit Function
  End If
On Error GoTo 0
  
' ������ CFG ����, ��������� ����
Dim Line
Dim intASig As Integer
Dim intDSig As Integer

objSheet.Cells(1, 1).Value = "����:": objSheet.Cells(1, 2).Value = strFileName

' ������ ������: �������� ������� � ������������� ������������
Line = ReadASCIILine(intCFGFile)
objSheet.Cells(2, 1).Value = "������������:": objSheet.Cells(2, 2).Value = ArrGet(Line, 0)
objSheet.Cells(3, 1).Value = "�����:": objSheet.Cells(3, 2).Value = ArrGet(Line, 1)

' ������ ������: ����� ���������� �������, ���������� ����������, ���������� ����������
Line = ReadASCIILine(intCFGFile)
intASig = GetInt(ArrGet(Line, 1))
intDSig = GetInt(ArrGet(Line, 2))

' ���������
objSheet.Cells(10, 1).Value = "SignalNo"
objSheet.Cells(11, 1).Value = "SignalName"
objSheet.Cells(12, 1).Value = "SignalPhase"
objSheet.Cells(13, 1).Value = "Component"
objSheet.Cells(14, 1).Value = "Meas"
objSheet.Cells(15, 1).Value = "A"
objSheet.Cells(16, 1).Value = "B"
objSheet.Cells(17, 1).Value = "Skew"
objSheet.Cells(18, 1).Value = "Min"
objSheet.Cells(19, 1).Value = "Max"

' ������ ���������� ������
Dim i As Integer
Dim j As Integer

For i = 1 To intASig
  Line = ReadASCIILine(intCFGFile)
  For j = 1 To 10
    objSheet.Cells(j + 9, i + 3).Value = Replace(ArrGet(Line, j - 1), ",", ".")
  Next
Next

' ������ ���������� ������
For i = 1 To intDSig
  Line = ReadASCIILine(intCFGFile)
  For j = LBound(Line) To UBound(Line)
    objSheet.Cells(j + 10, i + 3 + intASig).Value = ArrGet(Line, j)
  Next
Next

' ����� ���������� ������� ������������ ����� ������
' ������� �������
Line = ReadASCIILine(intCFGFile)
objSheet.Cells(4, 1).Value = "�������:": objSheet.Cells(4, 2).Value = GetInt(ArrGet(Line, 0))

' ���������� ������ �������������, �� �������, ��� ����� ������ 1
Line = ReadASCIILine(intCFGFile)

' ������������� � ���������� ��������
Line = ReadASCIILine(intCFGFile)
objSheet.Cells(5, 1).Value = "�������������:": objSheet.Cells(5, 2).Value = GetInt(ArrGet(Line, 0))
objSheet.Cells(6, 1).Value = "��������:": objSheet.Cells(6, 3).Value = GetInt(ArrGet(Line, 1))

' ����� ������ / �����
Line = ReadASCIILine(intCFGFile)
objSheet.Cells(7, 1).Value = "����/����� ������:":
objSheet.Cells(7, 2).NumberFormat = "@"
objSheet.Cells(7, 3).NumberFormat = "@"
objSheet.Cells(7, 2).Value = ArrGet(Line, 0)
objSheet.Cells(7, 3).Value = ArrGet(Line, 1)

Line = ReadASCIILine(intCFGFile)
objSheet.Cells(8, 1).Value = "����/����� �����:":
objSheet.Cells(8, 2).NumberFormat = "@"
objSheet.Cells(8, 3).NumberFormat = "@"
objSheet.Cells(8, 2).Value = ArrGet(Line, 0)
objSheet.Cells(8, 3).Value = ArrGet(Line, 1)

' ������ ������������� ������ (ASCII / BINARY)
Dim strDATFormat As String
Line = ReadASCIILine(intCFGFile)
strDATFormat = UCase$(ArrGet(Line, 0))

' ���� BINARY - �������� ��� �� ������ "EASY= 1"
On Error GoTo err_easy
  Dim strEASY As String
  Line = ReadASCIILine(intCFGFile)
  strEASY = UCase$(Replace(ArrGet(Line, 0), " ", ""))
err_easy:
Close #intCFGFile

' ������ ������ � ����������� �� �������
Dim ReadOk As Boolean
If strDATFormat = "ASCII" Then
  ReadOk = LoadDataASCII(strDATFile)
ElseIf strDATFormat = "BINARY" Then
  If strEASY = "EASY=1" Then
    ReadOk = LoadDataBINARYEasy(strDATFile, intASig, intDSig)
  Else
    ReadOk = LoadDataBINARY(strDATFile, intASig, intDSig)
  End If
Else
  OpenComtrade = 5  ' Error, ������ � ������� DAT �����, ������ ���� ASCII ��� BINARY
  Exit Function
End If

If Not ReadOk Then OpenComtrade = 6  ' Error, ������ ��� ������ DAT �����

' ��� ����� �� ������� ������ ��� ���������� ��������
objSheet.Cells(6, 2).NumberFormat = "0"
objSheet.Cells(6, 2).Formula = "=MAX(B20:B999999)"

OpenComtrade = 0

End Function


Private Function SaveComtrade(strFileName As String)
'
' ���������� COMTRADE ����� �� Excel (������ ASCII)

Dim objSheet
Set objSheet = ActiveWorkbook.ActiveSheet

Dim intCFGFile As Integer
Dim strDATFile As String
Dim intDATFile As Integer

strDATFile = ReplaceExt(strFileName, "dat")

' ���������� CFG

On Error GoTo err_cfg_file_access
  Dim errCFGFileAccess As Boolean
  intCFGFile = FreeFile
  Open strFileName For Output Access Write Lock Read Write As #intCFGFile
  Print #intCFGFile, objSheet.Cells(2, 2).Value & "," & objSheet.Cells(3, 2).Value
  errCFGFileAccess = True
err_cfg_file_access:
  If Not errCFGFileAccess Then
    SaveComtrade = 1  ' Error, ������ ��� ������ CFG �����
    Exit Function
  End If
On Error GoTo 0


' ������� ���������� ���������� � ���������� ��������
Dim intSig As Integer
Dim intASig As Integer
Dim intDSig As Integer
Dim i As Integer

For i = objSheet.UsedRange.Columns.Count To 4 Step -1
  If objSheet.Cells(10, i).Value <> "" Then
    intSig = i - 3
    Exit For
  End If
Next

For i = intSig + 4 To 4 Step -1
  If objSheet.Cells(15, i).Value <> "" Then
    intASig = i - 3
    Exit For
  End If
Next

intDSig = intSig - intASig

Print #intCFGFile, intSig & "," & intASig & "A," & intDSig & "D"

' ����� �������� �������
Dim Line As String
Dim j As Integer

For i = 1 To intASig
  Line = ""
  For j = 1 To 10
    Line = Line & "," & Replace(objSheet.Cells(j + 9, i + 3).Value, ",", ".")
  Next
  Line = Right(Line, Len(Line) - 1)
  Print #intCFGFile, Line
Next

For i = 1 To intDSig
  Line = ""
  For j = 1 To 3
    Line = Line & "," & Replace(objSheet.Cells(j + 9, i + 3 + intASig).Value, ",", ".")
  Next
  Line = Right(Line, Len(Line) - 1)
  Print #intCFGFile, Line
Next

' ���������� ������
Print #intCFGFile, Trim(objSheet.Cells(4, 2).Value)
Print #intCFGFile, "1"
Print #intCFGFile, objSheet.Cells(5, 2).Value & "," & objSheet.Cells(6, 2).Value
Print #intCFGFile, objSheet.Cells(7, 2).Value & "," & objSheet.Cells(7, 3).Value
Print #intCFGFile, objSheet.Cells(8, 2).Value & "," & objSheet.Cells(8, 3).Value
Print #intCFGFile, "ASCII"

Close #intCFGFile

' ����� DAT ���� � �������
On Error GoTo err_dat_file_access
  Dim errDATFileAccess As Boolean
  intDATFile = FreeFile
  Open strDATFile For Output Access Write Lock Read Write As intDATFile
  errDATFileAccess = True
err_dat_file_access:
  If Not errDATFileAccess Then
    SaveComtrade = 2  ' Error, ������ ��� ������ DAT �����
    Exit Function
  End If
On Error GoTo 0


Dim Rates As Long
Rates = objSheet.Cells(6, 2).Value

For i = 1 To Rates
  Line = ""
  For j = 1 To intSig + 2
    Line = Line & "," & objSheet.Cells(i + 19, j + 1).Value
  Next
  Line = Right(Line, Len(Line) - 1)
  Print #intDATFile, Line
Next

Close #intDATFile

End Function


Private Sub Prepare()
'
' ���������� ������ ���������, ����� �������� ������ �������

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

End Sub


Private Sub Ended()
'
' ���������� ������� ���������

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
    
End Sub


Public Sub Comtrade2Excel()
'
' ������ �������: ������ COMTRADE

Dim COMTRADEFile As String

COMTRADEFile = Application.GetOpenFilename("COMTRADE Files (*.cfg), *.cfg")

Prepare
Select Case OpenComtrade(COMTRADEFile)
  Case 1
     MsgBox "Error, CFG ���� �� ������", vbOKOnly, "Error"
  Case 2
     MsgBox "Error, DAT ���� �� ������", vbOKOnly, "Error"
  Case 3
     MsgBox "Error, ������ ������� � CFG �����", vbOKOnly, "Error"
  Case 4
     MsgBox "Error, ���� � ����� ������ ��� ����������", vbOKOnly, "Error"
  Case 5
     MsgBox "Error, ������ � ������� DAT �����, ������ ���� ASCII ��� BINARY", vbOKOnly, "Error"
  Case 6
     MsgBox "Error, ������ ��� ������ DAT �����", vbOKOnly, "Error"
End Select
Ended

End Sub


Public Sub Excel2Comtrade()
'
' ������ �������: ������ COMTRADE

Dim COMTRADEFile As String

COMTRADEFile = Application.GetSaveAsFilename(, "COMTRADE Files (*.cfg), *.cfg")

Prepare
Select Case SaveComtrade(COMTRADEFile)
  Case 1
     MsgBox "Error, ������ ��� ������ CFG �����", vbOKOnly, "Error"
  Case 2
     MsgBox "Error, ������ ��� ������ DAT �����", vbOKOnly, "Error"
End Select
Ended

End Sub
