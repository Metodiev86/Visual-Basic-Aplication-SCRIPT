Sub CashBook_Mail()
    Dim objMail As Outlook.MailItem
    Set objMail = Application.CreateItem(olMailItem)
        
    objMail.Display ' Показва черновата, можеш да замениш с .Send за автоматично изпращане
    objMail.Subject = "Корекции в Kасовата Kнига  за - " & Format(Date - 1, "dd.mm.yyyy")
    objMail.To = "Деан Стоянов"
    objMail.CC = "Красимир Илиев; Василка Трангова"
    objMail.HTMLBody = "Уважаеми г-н Стоянов, <br> <p style='margin-left:30px;'>   Изпращам Ви справка за корекциите в Касовата Книга за - " & Format(Date - 1, "dd.mm.yyyy") & "</p>" & objMail.HTMLBody
          
End Sub

Function GetFridayOfWeek() As Date
    Dim today As Date
    Dim daysToFriday As Integer
    
    today = Date ' Днешната дата
    daysToFriday = 6 - Weekday(today, vbMonday) ' Изчисляваме дните до петък (vbMonday започва от понеделник)

    GetFridayOfWeek = today + daysToFriday ' Добавяме дните към текущата дата
End Function

Sub TestFriday()
    MsgBox "Петъкът тази седмица е: " & Format(GetFridayOfWeek(), "dd-mm-yyyy")
End Sub
