Attribute VB_Name = "Comtrade2Excel"
'
' Comtrade2Excel Excel2Comtrade Converter
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' https://github.com/wyfinger/Comtrade2Excel
' Игорь Матвеев, miv@prim.so-ups.ru
' 2014
'

Option Explicit

Private Declare PtrSafe Function GetOEMCP Lib "kernel32" () As Long
Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
Private Declare PtrSafe Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long

Function ReplaceExt(ByVal strFileName, ByVal strNewExt As String) As String
'
' Смена расширения файла

Dim dot_pos As Long

dot_pos = InStrRev(strFileName, ".")
If dot_pos > 0 Then
  ReplaceExt = Left(strFileName, dot_pos) & strNewExt
Else
  ReplaceExt = "strFileName"
End If

End Function


Private Function ExtractFileName(ByVal strPath As String) As String
'
' Получить имя файла

Dim slash_pos As Long

If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)

slash_pos = InStrRev(Replace$(strPath, "/", "\"), "\")
If slash_pos > 0 Then
  ExtractFileName = Right$(strPath, Len(strPath) - slash_pos)
Else
  ExtractFileName = ""
End If
  
End Function

Function ReadNext(ByVal nFile)
'
' Читаем строку, преобразуем кодировку, сплитим ее в массив

Line Input #nFile, ReadedLine$

' сначала можно проверить номера кодовых таблиц
OemCP& = GetOEMCP    ' номер OEM (DOS) таблицы
AnsiCP& = GetACP     ' номер ANSI (Windows) таблицы

DecodedLine$ = Space$(Len(ReadedLine$))
Code& = OemToChar(ReadedLine$, DecodedLine$)

ReadNext = Split(DecodedLine, ",")

End Function

Function Ansii2Oem(ByVal strInput As String) As String
'
' Преобразование строки ANSII - OEM866 для записи

OemCP& = GetOEMCP
AnsiCP& = GetACP

rez$ = Space$(Len(strInput$))
Code& = OemToChar(strInput$, rez$)
'Ansii2Oem = rez$
Ansii2Oem = strInput

End Function

Private Sub LoadComtrade(ByVal strFileName As String)
'
' Загрузка и парсинг COMTRADE файла

Dim strConfig As String
Dim strData As String

strConfig = strFileName
strData = ReplaceExt(strFileName, "dat")

' Разбираем конфиг файл

Dim cName
Dim cNo
Dim cSignals
Dim cAnalogSignals
Dim cDigitalSignals

Dim dSignalNo
Dim dSignalName
Dim dSignalPhase
Dim dComponent
Dim dMeas
Dim dA
Dim dB
Dim dSkew
Dim dMin
Dim dMax

Dim eFreq
Dim eNRates           ' Количество частот дискретизации, обычно 1
Dim eRatesPerSec      ' частота дискретизации
Dim eRates           ' номер последней выборки

Dim nFile

Dim ReadedArr

' Создаем новый лист
Dim objRez
Set objRez = ActiveWorkbook.Worksheets.Add

' Читаем конфиг

nFile = FreeFile
Open strConfig For Input As #nFile

' Читаем первые две строки: Название и номер устройства, количество сигналов
ReadedArr = ReadNext(nFile)
cName = ReadedArr(0)
cNo = ReadedArr(1)
ReadedArr = ReadNext(nFile)
cSignals = ReadedArr(0)
cAnalogSignals = Left$(ReadedArr(1), Len(ReadedArr(1)) - 1)
cDigitalSignals = Left$(ReadedArr(2), Len(ReadedArr(2)) - 1)

' Записываем на лист общую информацию по осциллограмме
objRez.Name = ExtractFileName(strConfig)
objRez.Cells(1, 1).Value = "Файл:": objRez.Cells(1, 2).Value = strConfig
objRez.Cells(2, 1).Value = "Наименование:": objRez.Cells(2, 2).Value = cName
objRez.Cells(3, 1).Value = "Номер:": objRez.Cells(3, 2).Value = cNo



objRez.Cells(10, 1).Value = "SignalNo":
objRez.Cells(11, 1).Value = "SignalName":
objRez.Cells(12, 1).Value = "SignalPhase":
objRez.Cells(13, 1).Value = "Component":
objRez.Cells(14, 1).Value = "Meas":
objRez.Cells(15, 1).Value = "A":
objRez.Cells(16, 1).Value = "B":
objRez.Cells(17, 1).Value = "Skew":
objRez.Cells(18, 1).Value = "Min":
objRez.Cells(19, 1).Value = "Max":

' Читаем аналоговые / дискретные каналы
For i = 1 To cAnalogSignals
  ReadedArr = ReadNext(nFile)
  For j = 1 To 10
    objRez.Cells(j + 9, i + 3).Value = ReadedArr(j - 1)
  Next
Next
For i = 1 To cDigitalSignals
  ReadedArr = ReadNext(nFile)
  For j = 1 To 3
    objRez.Cells(j + 9, i + 3 + cAnalogSignals).Value = ReadedArr(j - 1)
  Next
Next

' Снова общие параметры
ReadedArr = ReadNext(nFile)
objRez.Cells(4, 1).Value = "Частота:": objRez.Cells(4, 2).Value = ReadedArr(0)
ReadedArr = ReadNext(nFile)   ' это количество частот дискретизации в файле, всегда 1
ReadedArr = ReadNext(nFile)
objRez.Cells(5, 1).Value = "Дискретизация:": objRez.Cells(5, 2).Value = ReadedArr(0)
objRez.Cells(6, 1).Value = "Отсчетов:": objRez.Cells(6, 3).Value = ReadedArr(1)
ReadedArr = ReadNext(nFile)
objRez.Cells(7, 1).Value = "Дата/Время старта:":
objRez.Cells(7, 2).NumberFormat = "@"
objRez.Cells(7, 3).NumberFormat = "@"
objRez.Cells(7, 2).Value = ReadedArr(0)
objRez.Cells(7, 3).Value = ReadedArr(1)
ReadedArr = ReadNext(nFile)
objRez.Cells(8, 1).Value = "Дата/Время старта:":
objRez.Cells(8, 2).NumberFormat = "@"
objRez.Cells(8, 3).NumberFormat = "@"
objRez.Cells(8, 2).Value = ReadedArr(0):
objRez.Cells(8, 3).Value = ReadedArr(1)

Close #nFile

' Читаем данные
nFile = FreeFile
Open strData For Input As #nFile

i = 20
Do While Not EOF(nFile)
  ReadedArr = ReadNext(nFile)
  For j = 1 To cSignals + 2
    objRez.Cells(i, j + 1).Value = ReadedArr(j - 1)
    'objRez.Cells(i, 1).Formula = "=(1/B$5)*(B" & i & "-1)"
    'objRez.Cells(i, 1).NumberFormat = "0.0000"
  Next
  i = i + 1
Loop
  
Close #nFile

' Это чтобы не править каждый раз количество отсчетов
objRez.Cells(6, 2).Formula = "=МАКС(B20:B999999)":

End Sub

Private Sub SaveComtrade(ByVal strFileName As String)
'
' Сохранение COMTRADE файла

Dim strConfig As String
Dim strData As String

Dim objRez
Set objRez = ActiveWorkbook.ActiveSheet

strConfig = strFileName
strData = ReplaceExt(strFileName, "dat")

' Записываем конфиг

nFile = FreeFile
Open strConfig For Output As #nFile
Print #nFile, Ansii2Oem(objRez.Cells(2, 2).Value & "," & objRez.Cells(3, 2).Value)

Dim cSignals
Dim cAnalogSignals
Dim cDigitalSignals

For i = 1 To 1000
  If objRez.Cells(10, 4 + i).Value = "" Then
    cSignals = i
    Exit For
  End If
Next

For i = cSignals + 4 To 2 Step -1
  If objRez.Cells(15, 4 + i).Value <> "" Then
    cDigitalSignals = cSignals - i - 1
    Exit For
  End If
Next

cAnalogSignals = cSignals - cDigitalSignals

Print #nFile, cSignals & "," & cAnalogSignals & "A," & cDigitalSignals & "D"

' Пишем сигналы аналог / дискрет

For i = 1 To cAnalogSignals
  strLine = ""
  For j = 1 To 10
    strLine = strLine & "," & objRez.Cells(j + 9, i + 3).Value
  Next
  strLine = Right(strLine, Len(strLine) - 1)
  Print #nFile, Ansii2Oem(strLine)
Next

For i = 1 To cDigitalSignals
  strLine = ""
  For j = 1 To 3
    strLine = strLine & "," & objRez.Cells(j + 9, i + 3 + cAnalogSignals).Value
  Next
  strLine = Right(strLine, Len(strLine) - 1)
  Print #nFile, Ansii2Oem(strLine)
Next

' Остаточные данные
Print #nFile, Ansii2Oem(objRez.Cells(4, 2).Value)
Print #nFile, "1"
Print #nFile, objRez.Cells(5, 2).Value & "," & objRez.Cells(6, 2).Value
Print #nFile, objRez.Cells(7, 2).Value & "," & objRez.Cells(7, 2).Value
Print #nFile, objRez.Cells(8, 2).Value & "," & objRez.Cells(8, 2).Value
Print #nFile, "ASCII"

Close #nFile

' Пишем данные

nFile = FreeFile
Open strData For Output As #nFile

Rates = objRez.Cells(6, 2).Value

For i = 1 To Rates
  strLine = ""
  For j = 1 To cSignals + 2
    strLine = strLine & "," & objRez.Cells(i + 19, j + 1).Value
  Next
  strLine = Right(strLine, Len(strLine) - 1)
  Print #nFile, Ansii2Oem(strLine)
Next

Close #nFile

End Sub


Public Sub Comtrade2Excel()
'
' Запуск макроса: Чтение COMTRADE

'LoadComtrade ("C:\Users\wyfinger\Desktop\Склейка осцилограмм\2Р_ДЗШ_ФC.cfg")
Application.ScreenUpdating = False
LoadComtrade (Application.GetOpenFilename("COMTRADE Files (*.cfg), *.cfg"))
Application.ScreenUpdating = True

End Sub

Public Sub Excel2Comtrade()
'
' Запуск макроса: Запись COMTRADE

'Application.ScreenUpdating = False
SaveComtrade (Application.GetSaveAsFilename(, "COMTRADE Files (*.cfg), *.cfg"))
'Application.ScreenUpdating = True

End Sub
