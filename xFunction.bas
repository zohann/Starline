Option Compare Database

Function CreateTable(strTableName) As Recordset
'--------------------------------------------------------------------
Dim tbl As TableDef       'объект таблица
Dim idx As Index          'объект индекс
Dim fld As Field          'объект поле
Dim rst As Recordset      'объект набор записей
Dim i As Integer          'счетчик дней
 

On Error Resume Next

'создание объектной переменной таблицы, полей и индекса в ней
    Set tbl = CurrentDb.CreateTableDef(strTableName)
    With tbl
        Set fld = tbl.CreateField("idPoint") 'Создание
        fld.Type = dbLong
        fld.Properties("Attributes") = dbAutoIncrField 'Назначение атрибутов счетчика
        .Fields.Append fld 'Добавить поле
'        .Fields.Append tbl.CreateField("idPoint", dbLong)
        .Fields.Append tbl.CreateField("idTrack", dbLong)
        .Fields.Append tbl.CreateField("lat", dbText, 20)
        .Fields.Append tbl.CreateField("lon", dbText, 20)
        .Fields.Append tbl.CreateField("timestamp", dbLong)
        .Fields.Refresh '
            'создание уникального индекса
            Set idx = .CreateIndex("Primary Key")
                With idx
                    .Fields.Append .CreateField("idPoint") 'добавление поля в индекс
                    .Unique = True   'Уникальный
                    .Primary = True  'Первичный
                End With
            .Indexes.Append idx
           'индекс создан
    End With
' Фактическое добавление таблицы из объектной переменной описанной выше
    CurrentDb.TableDefs.Append tbl
 
    Set CreateTable = CurrentDb.OpenRecordset(strTableName)
' Заполнение таблицы данными
'    Set rst = CurrentDb.OpenRecordset(strTableName, dbOpenDynaset)
'    With rst
'        For i = 1 To 7
'            .AddNew
'            !DayID = i
'            !DayName = DayName(i)
'            .Update
'        Next i
'    End With
 
CreateWeekDaysTableBye:
    On Error Resume Next
    Set idx = Nothing
    Set tbl = Nothing
    rst.Close
    Set rst = Nothing
    Exit Function
 
'CreateWeekDaysTableErr:
'    MsgBox "Произошла ошибка при выполнении процедуры " & _
'    "[CreateWeekDaysTable] :" & vbCrLf & _
'    Err.Description & vbCrLf & _
'    "Номер ошибки = " & Err.Number, vbCritical
'    Resume CreateWeekDaysTableBye
End Function
 
'--------------------------------------------------------------------
 
Private Function DayName(DayNo As Integer) As String
'es 26.10.2000
'Вспомагательная = Возвращает название дня недели по его номеру
'--------------------------------------------------------------------
On Error GoTo DayNameErr
    Select Case DayNo
        Case 1: DayName = "Понедельник"
        Case 2: DayName = "Вторник"
        Case 3: DayName = "Среда"
        Case 4: DayName = "Четверг"
        Case 5: DayName = "Пятница"
        Case 6: DayName = "Суббота"
        Case 7: DayName = "Воскресенье"
    End Select
DayNameBye:     Exit Function
DayNameErr:     DayName = "#Error#": Resume DayNameBye
End Function

' ********************************* Чтение текстового файла в переменную
Function ReadTXTfile(ByVal filename As String) As String
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.OpenTextFile(filename, 1, True): ReadTXTfile = ts.ReadAll: ts.Close
    Set ts = Nothing: Set fso = Nothing
End Function

' ********************************* Запись в текстовый файл из переменной
Function SaveTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    On Error Resume Next: Err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.CreateTextFile(filename, True)
    ts.Write txt: ts.Close
    SaveTXTfile = Err = 0
    Set ts = Nothing: Set fso = Nothing
End Function

' ********************************* Добавление в текстовый файл из переменной
Function AddIntoTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    On Error Resume Next: Err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.OpenTextFile(filename, 8, True): ts.Write txt: ts.Close
    Set ts = Nothing: Set fso = Nothing
    AddIntoTXTfile = Err = 0
End Function

' *****************************************
Function ConvertTimestampToDate(timestamt)
    '1700106124 = 16.11.2023, 08:42:04
'    timestamt = 1700106124
    ConvertTimestampToDate = DateAdd("h", utc, fromUnix10(timestamt))
End Function

' *****************************************
Public Function fromUnix10(ts) As Date
    fromUnix10 = DateAdd("s", CDbl(ts), "1/1/1970")
End Function

Public Function toUnix(dt) As Long
    toUnix = DateDiff("s", "1/1/1970", dt)
End Function

' *****************************************
Function fromUNIX13Digits(uT) As Date
   fromUNIX13Digits = CDbl(uT) / 86400000 + DateSerial(1970, 1, 1)
End Function

' ******************************************* Формат даты
Function getDateFormat(dat, Optional addtime As Boolean) As String
    If addtime Then
        getDateFormat = "#" & Format(dat, "mm\/dd\/yyyy hh:mm:ss") & "#"
    Else
        getDateFormat = "#" & Format(dat, "mm\/dd\/yyyy") & "#"
    End If
End Function
