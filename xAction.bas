Option Compare Database
Option Explicit

' ******************************************* Запуск импорта треков
Private Function aStart()
    Dim collFiles As Collection, file_name As Variant, tx As String, txt As String
    Dim obj As Object, objData As Object
    ' ****************************
    Set collFiles = setCollFiles() ' выбор файлов треков, коллекция файлов с треками (формат json)
    If collFiles.Count = 0 Then Exit Function
    For Each file_name In collFiles ' цикл по файлам
        tx = tx & getNameFile(file_name) & " | "
        ' **************************** Чтение файла в переменную
        txt = ReadTXTfile(file_name)
        If Len(txt) > 0 Then
            ' **************************** Конвертация в Json
            Set obj = JsonConverter.ParseJson(txt)
            Set objData = obj("data")
            Call readTrackFile(objData, tx)  ' чтение файла
        End If
    Next
    Set obj = Nothing
    Set objData = Nothing
    MsgBox tx
End Function

' ******************************************* Имя файла
Function getNameFile(file_name) As String
    getNameFile = Split(file_name, "\")(UBound(Split(file_name, "\")))
End Function

' ******************************************* Чтение файла треков
Function readTrackFile(ByRef objData As Object, tx As String) As String
    Dim rs As Recordset, dat As Variant, point_count As Long, point_count_all As Long, track_count As Long
    Dim nods As Object, noda As Object
    Dim mile As String, datetime_start As String, atime As String, idTrack As Long
    track_count = 0
    ' ****************************
    For Each dat In objData ' цикл по коллекции треков текущего файла
        If dat("type") = "TRACK" Then
            track_count = track_count + 1
            If dat("mileage") > 500 Then
                point_count = 0
                Set nods = dat("nodes") ' коллекция точек
                For Each noda In nods ' цикл по точкам трека
                    mile = noda("mileage")
                    If mile <> "" And mile <> "0" Then
                        If point_count = 0 Then ' первая точка
                            datetime_start = ConvertTimestampToDate(noda("t")) ' дата/время
                            idTrack = addTrackData(datetime_start) ' запись первого трека
                            If idTrack = 0 Then Exit For ' выход - если запись уже имеется
                            ' ***************************** Проверка наличия таблицы
                            Call checkTable(rs, datetime_start)
                        End If
                        point_count = point_count + 1
                        atime = noda("t")
                        Call addPointToTrack(rs, idTrack, mile, atime, noda("x"), noda("y"))
                    End If
                Next
                If idTrack > 0 Then Call addTrackEndData(idTrack, datetime_start, ConvertTimestampToDate(atime))
                point_count_all = point_count_all + point_count
            End If
        End If
    Next
    tx = tx & "Tracks: " & track_count & " | Points: " & point_count_all & Chr(13)
End Function

' ******************************************* Таблица трека
Function checkTable(rs As Recordset, datetime As String)
    Dim table_name As String
    table_name = Format(datetime, "yyyy_mm")
    If TypeName(rs) = "Nothing" Then
        Set rs = CreateTable(table_name)
    Else
        If rs.name <> table_name Then
            Set rs = CreateTable(table_name)
        End If
    End If
End Function

' *******************************************
Function addPointToTrack(rs As Recordset, idTrack, mile, atime, lat, lon) As Long
    With rs
        .AddNew
        !idTrack = idTrack
        !lat = Replace(lat, ",", ".")
        !lon = Replace(lon, ",", ".")
        !timestamp = atime
        .Update
    End With
End Function

' *******************************************
Function addTrackEndData(idTrack, time_start, time_end) As Long
    Dim diff_min As Long, diff_hour As Date, diff As Date
    diff_min = DateDiff("n", time_start, time_end)
    diff_hour = TimeSerial(0, diff_min, 0)
    diff = Format(diff_hour, "HH:mm")
    CurrentDb.Execute ("UPDATE tracks SET date_end = '" & Format(time_end, "dd.mm.yyyy") & "', time_end = '" & Format(time_end, "hh:mm:ss") & "', time_go = '" & diff & "' WHERE idTrack = " & idTrack)
End Function

' ******************************************* Запись трека
Function addTrackData(datetime) As Long
    Dim date_start As Date, time_start As Date, crit As String
    date_start = Format(datetime, "dd.mm.yyyy")
    time_start = Format(datetime, "hh:mm:ss")
    crit = "date_start = " & getDateFormat(date_start) & " AND time_start =#" & time_start & "#"
    With CurrentDb.OpenRecordset("SELECT * FROM tracks")
        .FindFirst (crit)
        If .NoMatch Then
            .AddNew
            !date_start = date_start
            !time_start = time_start
            .Update
            addTrackData = DMax("idTrack", "tracks", crit)
        End If
    End With
End Function
