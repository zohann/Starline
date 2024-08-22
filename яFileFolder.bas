Public cInitial As String

' ******************************************* Коллекция файлов с треками (формат json)
Function setCollFiles() As Collection
    Set setCollFiles = New Collection
    With Application.FileDialog(1)
        .InitialFileName = CurrentProject.Path & "\"
        If .Show = False Then Exit Function
        For Each f In .SelectedItems
            setCollFiles.add f
        Next
    End With
End Function

' ******************************************* Считать файл
Function getTxt(filepath)
    getTxt = LoadTextFromTextFile(filepath, "UTF-8")
End Function

' ******************************************* Записать текст в файл
Function txtToFile(txt, inn)
    filepath = CurrentProject.Path & "\" & Format(Now, "yyyy-mm-dd_hh.mm") & "_[" & inn & "]" & ".txt" ' путь и имя файла
    Call putTxt(filepath, txt, True) ' запись ответа в файл
End Function

' ******************************************* Получить текст из файла
Function ReadAllTextFile(path_file)
   Const ForReading = 1, ForWriting = 2
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
'   Set f = fso.OpenTextFile(path_file, ForWriting, True)
'   f.Write "Hello world!"
   Set f = fso.OpenTextFile(path_file, ForReading)
   ReadAllTextFile = f.ReadAll
End Function

' ******************************************* Получить текст из файла
Function LoadTextFromTextFile(ByVal filename, Optional ByVal encoding) As String
    On Error Resume Next: Dim txt$
    If Trim(encoding) = "" Then encoding = "windows-1251"
    With CreateObject("ADODB.Stream")
        .Type = 2:
        If Len(encoding) Then .Charset = encoding
        .Open
        .LoadFromFile filename        ' загружаем данные из файла
        LoadTextFromTextFile = .ReadText        ' считываем текст файла
        .Close
    End With
End Function

' ******************************************* Запись в файл
Function putTxt(filetxt, txt, Optional utf As Boolean)
    Dim fsT As Object
    If txt = "" Then Exit Function
    If utf Then
        With CreateObject("ADODB.Stream")
            .Type = 2
            .Charset = "utf-8"
            .Open
            .WriteText txt
            .SaveToFile filetxt, 2
        End With
    Else
        Open filetxt For Output As #1: Print #1, txt: Close #1
    End If
End Function

'Function path_folder_clients() As String
'    path_folder_clients = "\\portal-05.bank\diskU\ДВПА\УВЗФЛ\ФИНАНСОВОЕ ОЗДОРОВЛЕНИЕ\Клиенты\"
'End Function

' ******************************************* Открыть папку
'Function FolderClientOpen() As Boolean
'    cl = Forms("Рабочая")!client0
'    folder_clients = path_folder_clients()
'    folder_path = folder_clients & cl
'    Set fso = CreateObject("scripting.filesystemobject")
'    If fso.FolderExists(folder_clients) Then
'        If fso.FolderExists(folder_path) Then
'            Application.FollowHyperlink folder_path
'        Else
'            Application.FollowHyperlink folder_clients
'        End If
'    Else
'        MsgBox "Папка ''" & folder_clients & "'' не найдена!", vbExclamation, "Внимание!"
'    End If
'End Function

' ******************************************* Создание папки
'Function FolderClientAdd(cl) As Boolean
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    folder_clients = path_folder_clients()
'    folder_path = folder_clients & cl
'    If fso.FolderExists(folder_clients) Then
'        If fso.FolderExists(folder_path) Then
'            Application.FollowHyperlink folder_path
'        Else
'            If MsgBox("Создать папку ''" & cl & "''?", vbYesNo, "Внимание!") = vbYes Then
'                MkDir folder_path
'                For i = 1 To 200000000: Next
'                If fso.FolderExists(folder_path) Then
'                    Application.FollowHyperlink folder_path
'                End If
'            End If
'        End If
'    Else
'        MsgBox "Папка ''" & folder_clients & "'' не найдена!", vbExclamation, "Внимание!"
'    End If
'End Function

Function getPathDoc(txPath) As String
    getPathDoc = "<a href=""file:///" & txPath & """>" & txPath & "</a>"
End Function

' ******************************************* Проверка наличия папки
Function ifPath(mPath) As Boolean
On Error Resume Next
    ifPath = Dir(mPath, vbDirectory) <> ""
End Function

' *******************************************
Function getPathDB() As String
    strBackEndPath = CurrentDb.TableDefs("ColumnOrders").Connect
    j = InStrRev(strBackEndPath, "=") + 1
    getPathDB = Mid(strBackEndPath, j)
End Function

Function Get_All_File_from_SubFolders()
    Dim sFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then Exit Function
        sFolder = .SelectedItems(1)
    End With
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    GetSubFolders sFolder
    Set objFolder = Nothing
    Set objFSO = Nothing
End Function

' *******************************************
Private Function GetSubFolders(sPath)
    Dim sPathSeparator As String, sObjName As String
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(sPath)
    For Each objFile In objFolder.Files
        fl = Replace(objFile.name, objFSO.GetBaseName(objFile), "")
        If fl Like ".odt*" Or fl Like ".doc*" Then
            colFiles.add objFolder & "\" & objFile.name
        End If
    Next
    For Each objFolder In objFolder.SubFolders
        GetSubFolders objFolder.Path
    Next
End Function

' *******************************************
Function Get_All_File_from_Folder()
    Dim sFolder As String, sFiles As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then Exit Function
        sFolder = .SelectedItems(1)
    End With
    sFiles = Dir(sFolder & "\*.odt*")
    Do While sFiles <> ""
        sFiles = Dir
    Loop
End Function

' *******************************************
Function ChangeFolder() ' для отдельной кнопки - если вдруг надо поменять ранее выбранную папку
    On Error Resume Next: GetFolder , True
End Function

' *******************************************
Function GetFolder(Optional ByVal FolderIndex& = 0, Optional ByVal ShowDialog As Boolean = False, _
                   Optional ByVal title$ = "Выберите папку", Optional ByVal InitialFolder$) As String
    ' При первом вызове выводит диалогое окно выбора папки
    ' Запоминает выбранную папку, и при следующих вызовах диалоговое окно не выводит,
    ' а возвращает путь к ранее выбиравшейся папке
    ' Используйте вызов с параметром ShowDialog=TRUE для принудительного отображения диалогового окна
    On Error Resume Next: Err.Clear
    ProjectName$ = IIf(Len(PROJECT_NAME$) > 0, PROJECT_NAME$, "SelectFolder")
    PreviousFolder$ = GetSetting(Application.name, ProjectName$, "folder" & FolderIndex&, "")
    If Len(PreviousFolder$) > 0 And Not ShowDialog Then
        If Dir(PreviousFolder$, vbDirectory) <> "" Then GetFolder = PreviousFolder$: Exit Function
    End If
 
    If InitialFolder$ = "" Then
        If Len(PreviousFolder$) > 0 And Dir(PreviousFolder$, vbDirectory) <> "" Then
            InitialFolder$ = PreviousFolder$    ' начинаем обзор с ранее выбранной папки
        Else
            InitialFolder$ = ThisWorkbook.Path & "\"    ' начинаем с текущей папки
        End If
    End If
 
    With Application.FileDialog(msoFileDialogFolderPicker)    ' вывод диалогового окна
        .ButtonName = "Выбрать"
        .title = title
        .InitialFileName = InitialFolder$
        .AllowMultiSelect = True
        If .Show <> -1 Then Exit Function    ' если пользователь отказался от выбора папки
        GetFolder = .SelectedItems(1)
        If Not Right$(GetFolder, 1) = "\" Then GetFolder = GetFolder & "\"
        SaveSetting Application.name, ProjectName$, "folder" & FolderIndex&, GetFolder
    End With
End Function

' ******************************************* Проверка открыта ли книга
Function IsBookOpen(wbFullName As String) As Boolean
    On Error Resume Next
    Open wbFullName For Random Access Read Write Lock Read Write As 1
    Close 1
    IsBookOpen = Err
End Function

' *******************************************
Function ВыборФайлов(Optional initial = "") As Collection
    With Application.FileDialog(1)
        .AllowMultiSelect = True
        .InitialView = msoFileDialogViewList
'        .title = title 'заголовок окна диалога
        .Filters.Clear
        .InitialFileName = IIf(initial = "", CurrentProject.Path & "\", initial)
        .Filters.add "файлы Excel", "*.xls*"
        .FilterIndex = 1
        Set ВыборФайлов = New Collection
        If .Show = 0 Then Exit Function
        For i = 1 To .SelectedItems.Count
            ВыборФайлов.add .SelectedItems(i)
        Next
        cInitial = .InitialFileName
    End With
End Function

' *******************************************
Function getWord(Optional filename = "", Optional tempa As Boolean)
    On Error Resume Next
    Dim wordApp, Doc As Object
    If filename = "" Then
        Set xlApp = CreateObject("Excel.Application")
        Set getWb_ = xlApp.Workbooks.add
        xlApp.visible = True
    Else
        Set wordApp = GetObject(, "Word.Application")
        If wordApp Is Nothing Then
            Set wordApp = CreateObject("Word.Application")
            If tempa Then
                Set getWord = wordApp.Documents.add(filename)
            Else
                wordApp.Documents.Open filename
                Set getWord = wordApp.Documents(filename)
            End If
            wordApp.visible = True
        Else
            For Each w In wordApp.Documents
                If w.name = Dir(filename) Then
                    Set getWord = w
                    Exit For
                End If
            Next
        End If
        If TypeName(getWord) <> "Document" Then  '= "Nothing" Or TypeName(getWord) = "Empty" Then
'            DoEvents: SysCmd acSysCmdSetStatus, "Импорт данных | Открытие файла | " & Dir(filename)
            If tempa Then
                Set getWord = wordApp.Documents.add(filename)
            Else
                wordApp.Documents.Open filename
                Set getWord = wordApp.Documents(filename)
            End If
            wordApp.visible = True
        End If
    End If
    If TypeName(getWord) <> "Document" Then
        Set getWord = New Collection
        getWord.add Err.Description
    End If
End Function

' *******************************************
Function getWb(Optional filename = "")
    On Error Resume Next
    Dim xlApp As Object
    If filename = "" Then
        Set xlApp = CreateObject("Excel.Application")
        Set getWb = xlApp.Workbooks.add
    Else
        Set xlApp = GetObject(, "Excel.Application")
        If TypeName(xlApp) = "Nothing" Then Set xlApp = CreateObject("Excel.Application")
        For Each w In xlApp.Workbooks
            If w.name = Dir(filename) Then Set getWb = w
        Next
        If TypeName(getWb) <> "Workbook" Then
            Set getWb = xlApp.Workbooks.Open(filename)
            getWb.Activate: getWb.sheets(1).Activate
            getWb.Parent.visible = True
        End If
    End If
End Function

' *******************************************
Function LoadArrayFromWorkbook(ByVal filename$, ByVal FirstCellAddress$, Optional ByVal ColumnsCount& = 0, Optional sheetname = "", Optional title = "") As Variant
    On Error Resume Next: Err.Clear
    Set wb = GetObject(filename$)
    If wb Is Nothing Then MsgBox "Не удалось загрузить файл " & filename$: Exit Function
    If sheetname = "" Then
        Set sh = wb.Worksheets(1)
    Else
        Set sh = wb.Worksheets(sheetname)
    End If
    Set ra = sh.Range(sh.Range(FirstCellAddress$), sh.Range(FirstCellAddress$).EntireColumn.Cells(sh.Rows.Count).End(xlUp))
    If ra Is Nothing Then MsgBox "Не удалось обработать таблицу из файла " & filename$, vbExclamation, "Ошибка": wb.Close False: Exit Function
    If ColumnsCount& = 0 Then ColumnsCount& = sh.Columns.Count - sh.Range(FirstCellAddress$).Column + 1
    Err.Clear
    Set ra = ra.Resize(, ColumnsCount&)
    If Err Then MsgBox "Не удалось обработать таблицу из файла " & filename$, vbExclamation, "Ошибка": wb.Close False: Exit Function
    LoadArrayFromWorkbook = ra.Value
    wb.Close False
    Set wb = Nothing
End Function

' *******************************************
Function getШаблон(doc_name) As String
    strFullPath = getPathDB()
    getШаблон = Left(strFullPath, InStrRev(strFullPath, "\")) & doc_name
    If Dir(getШаблон) = "" Then MsgBox "Шаблон ''" & doc_name & "''не найден!", vbExclamation, "Внимание!": getШаблон = ""
End Function

' *******************************************
Function getDoc(mFile)
    Set wordApp = CreateObject("Word.Application")
    wordApp.visible = True
    Set getDoc = wordApp.Documents.add(mFile)
End Function

' *******************************************
Function ptxt(txt)
    Open CurrentProject.Path & "\" & Format(Now, "yyyy-MM-dd_hh-mm") & ".txt" For Output As #1
    Print #1, txt
    Close #1
End Function

' *******************************************
Function addtxt(ByVal txt As String) As Boolean
    On Error Resume Next: Err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.OpenTextFile(filelog, 8, True)
    ts.Write Format(Now, "yyyy-MM-dd hh-mm-ss ") & txt
    ts.Close
    Set ts = Nothing: Set fso = Nothing
End Function

' *******************************************
Function mkdir2(Путь)
On Error GoTo error_
    Set fso = CreateObject("Scripting.FileSystemObject")
    a = Split(Путь, "\")
    For i = 0 To UBound(a)
        If a(i) <> "" Then
            aa = aa & a(i) & "\"
        If fso.FolderExists(aa) = False Then MkDir aa
        End If
    Next
exiterror:
    Exit Function
error_:
    MsgBox "Ошибка:" & Err.Number & vbCrLf & Err.Description, vbCritical, "Warning", Err.HelpFile, Err.HelpContext
    Resume exiterror:
End Function
