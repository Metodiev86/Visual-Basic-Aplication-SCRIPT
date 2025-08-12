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
    
    ' ��������� �� Outlook ����������
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If OutApp Is Nothing Then
        MsgBox "Outlook �� � ���������� ��� ������������.", vbExclamation
        Exit Sub
    End If

    ' ���������� �� �������
        signature = "<div style='color: rgb(47, 84, 150);'>" & _
            "� ��������,<br><br>" & _
            "����� ��������<br>" & _
            "������������� ���� ����� (DBA)<br>" & _
            "������ �� ����<br>" & _
            "�������, ���. ��������� ���� 17<br>" & _
            "<a href='https://stabil-di.com/'>https://stabil-di.com/</a><br>" & _
            "�������: +359 888 415 383<br>" & _
            "�����: s.metodiev@stabil-di.com</div>"

    ' ������� ����
    Set ws = ThisWorkbook.Sheets("����1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' ������������� �� FileSystemObject � Dictionary
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set fileDictionary = CreateObject("Scripting.Dictionary")
    
    ' ����� �� ����� � �������
    Dim folderDialog As FileDialog
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With folderDialog
        .Title = "�������� ������� � ���������"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "�� � ������� �����. �������� �� ����������.", vbExclamation
            Exit Sub
        End If
    End With

    ' ��������� �� ��������� � Dictionary
    If fileSystem.FolderExists(folderPath) Then
        Dim file As Object
        For Each file In fileSystem.GetFolder(folderPath).Files
            fileDictionary(file.Name) = file.Path
        Next file
    End If
    
    ' �������� �� ������ ������ �� ������ B
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
    
    ' ��������� �� �������
    For Each cell In ws.Range("A2:A" & lastRow)
        emailSent = False
        emailAddress = Trim(cell.Value)
        searchName = Trim(ws.Cells(cell.Row, "C").Value)
        attachmentPath = ""

        ' ������� �� ���� �� ���������
        Dim key As Variant
        For Each key In fileDictionary.Keys
            If InStr(1, key, searchName, vbTextCompare) > 0 Then
                attachmentPath = fileDictionary(key)
                Exit For
            End If
        Next key
        
        ' ��������� �� ������
        If attachmentPath <> "" Then
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = emailAddress
                .CC = allCCEmails
                .Subject = "��������� �� �������� ����� - " & searchName & " - " & Format(Date, "dd.mm.yyyy")
                .HTMLBody = "�������� ���������, <br><br>" & _
                        "�� ���� �������� �� ��������� �������� ������� �� �������� �� ���������� ��� ����������� ��.  <br><br>" & _
                        "����, ��������� ������� �� ����������� ��������� � ������� �����. <br>" & _
                        "��� ������������� ���������� ���������� ���� �� �������, ���� �� �� ��������!<br><br><br><br>" & _
                        Replace(signature, vbCrLf, "<br>") ' �������� �� ������� � HTML �����������
                .Attachments.Add attachmentPath
                .Send
                emailSent = True
            End With
            Set OutMail = Nothing
        End If

        ' ���������� �� �������� ������ ���������
        If emailSent Then
            cell.Interior.Color = RGB(144, 238, 144)
        Else
            cell.Interior.Color = RGB(255, 0, 0)
        End If
    Next cell

    ' ������������� �� ��������
    Set OutApp = Nothing
    Set fileSystem = Nothing
    Set fileDictionary = Nothing
    MsgBox "�������� �������� �������!", vbInformation
End Sub


