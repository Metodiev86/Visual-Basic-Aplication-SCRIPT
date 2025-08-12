Sub SendEmailsWithAttachments()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim folderPath As String
    Dim cell As Range
    Dim emailAddress As String
    Dim allCCEmails As String
    Dim searchName As String
    Dim attachmentPath As String
    Dim emailSent As Boolean
    Dim fileSystem As Object
    Dim fileDictionary As Object
    Dim signature As String
    
    ' Създаване на Outlook приложение
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If OutApp Is Nothing Then
        MsgBox "Outlook не е инсталиран или конфигуриран.", vbExclamation
        Exit Sub
    End If

    ' Дефиниране на подписа
        signature = "<div style='color: rgb(47, 84, 150);'>" & _
            "С Уважение,<br><br>" & _
            "Стоян Методиев<br>" & _
            "Администратор База Данни (DBA)<br>" & _
            "СТАБИЛ ДИ ЕООД<br>" & _
            "Пловдив, Бул. Кукленско шосе 17<br>" & _
            "<a href='https://stabil-di.com/'>https://stabil-di.com/</a><br>" & _
            "Мобилен: +359 888 415 383<br>" & _
            "Имейл: s.metodiev@stabil-di.com</div>"

    ' Работен лист
    Set ws = ThisWorkbook.Sheets("Лист1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Инициализация на FileSystemObject и Dictionary
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set fileDictionary = CreateObject("Scripting.Dictionary")
    
    ' Избор на папка с файлове
    Dim folderDialog As FileDialog
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With folderDialog
        .Title = "Изберете папката с файловете"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "Не е избрана папка. Процесът се прекратява.", vbExclamation
            Exit Sub
        End If
    End With

    ' Записване на файловете в Dictionary
    If fileSystem.FolderExists(folderPath) Then
        Dim file As Object
        For Each file In fileSystem.GetFolder(folderPath).Files
            fileDictionary(file.Name) = file.Path
        Next file
    End If
    
    ' Събиране на всички имейли от колона B
    allCCEmails = ""
    Dim ccCell As Range
    For Each ccCell In ws.Range("B2:B" & lastRow)
        If Trim(ccCell.Value) <> "" Then
            If allCCEmails = "" Then
                allCCEmails = Trim(ccCell.Value)
            Else
                allCCEmails = allCCEmails & "; " & Trim(ccCell.Value)
            End If
        End If
    Next ccCell
    
    ' Обхождане на данните
    For Each cell In ws.Range("A2:A" & lastRow)
        emailSent = False
        emailAddress = Trim(cell.Value)
        searchName = Trim(ws.Cells(cell.Row, "C").Value)
        attachmentPath = ""

        ' Търсене на файл за прикачане
        Dim key As Variant
        For Each key In fileDictionary.Keys
            If InStr(1, key, searchName, vbTextCompare) > 0 Then
                attachmentPath = fileDictionary(key)
                Exit For
            End If
        Next key
        
        ' Изпращане на имейла
        If attachmentPath <> "" Then
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = emailAddress
                .CC = allCCEmails
                .Subject = "Напомняне за настъпил падеж - " & searchName & " - " & Format(Date, "dd.mm.yyyy")
                .HTMLBody = "Уважаеми партньори, <br><br>" & _
                        "За ваше удобство ви изпращаме подробна справка за текущите ви задължения към дружеството ни.  <br><br>" & _
                        "Моля, извършете плащане по маркираните документи с изтекъл падеж. <br>" & _
                        "Ако междувременно посочените задължения вече са платени, моля да ни извините!<br><br><br><br>" & _
                        Replace(signature, vbCrLf, "<br>") ' Добавяне на подписа с HTML форматиране
                .Attachments.Add attachmentPath
                .Send
                emailSent = True
            End With
            Set OutMail = Nothing
        End If

        ' Оцветяване на клетката според резултата
        If emailSent Then
            cell.Interior.Color = RGB(144, 238, 144)
        Else
            cell.Interior.Color = RGB(255, 0, 0)
        End If
    Next cell

    ' Освобождаване на обектите
    Set OutApp = Nothing
    Set fileSystem = Nothing
    Set fileDictionary = Nothing
    MsgBox "Процесът приключи успешно!", vbInformation
End Sub


