Option Compare Database

Private Function test()
    Set colDates = New Collection
'    colDates.add "20.08.2023"
    For i = 11 To 30
        colDates.add i & ".09.2023"
    Next
    Call trackToGPS(colDates)
End Function

Function trackToGPS(colDates)
    ' **************************** Файл шаблона
    filepath = CurrentProject.Path & "\track.gpx"
    txfile = LoadTextFromTextFile(filepath)
    For Each dat In colDates
        trackfilepath = CurrentProject.Path & "\" & Format(dat, "yyyy-mm-dd_") & "[" & Format(Now, "yyyy-mm-dd HH_mm") & "]" & "_track.gpx"
        table_name = Format(dat, "yyyy_mm")
        txP = ""
        With CurrentDb.OpenRecordset("SELECT * FROM [tracks] WHERE date_start = " & getDateFormat(dat))
            Do While Not .EOF
                With CurrentDb.OpenRecordset("SELECT * FROM [" & table_name & "] WHERE idTrack = " & !idTrack & " ORDER BY idPoint")
                    Do While Not .EOF
                        dt = ConvertTimestampToDate(!timestamp)
                        atime = "<time>" & Format(dt, "yyyy-mm-ddThh:mm:ssZ") & "</time>"
                        txP = txP & Chr(13) & "<trkpt lat=""" & !lat & """ lon=""" & !lon & """>" & atime & "</trkpt>"
                        .MoveNext
                    Loop
                End With
                .MoveNext
            Loop
        End With
        Call putTxt(trackfilepath, Replace(txfile, "[TRKPT]", txP), True)
    Next
End Function
