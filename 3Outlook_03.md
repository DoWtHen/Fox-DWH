```vba
'====================================
'       VÁLASZ SZÖVEG SABLONOK
'         FoxConn és egyebek
'====================================

Sub ValaszSzoveg_I()
' DoWtHen Makró 2026.10.01
' Válasz Szöbeg Sablonok
' CSAK a levél legelejére tud beszúrni szöveget!
' IC tesztesek

Dim sel As Outlook.Selection
Dim mail As Outlook.MailItem

    ' Kijelölt elem lekérése
    Set sel = Application.ActiveExplorer.Selection
    If sel.Count = 0 Then
        MsgBox "Nincs kijelölt levél.", vbExclamation
        Exit Sub
    End If

    ' Csak MailItem esetén
    If TypeOf sel.Item(1) Is Outlook.MailItem Then
        Set mail = sel.Item(1)
    Else
        MsgBox "Ez nem e-mail.", vbExclamation
        Exit Sub
    End If

    ' --- SZÖVEG BESZÚRÁSA A VÁLASZ ELEJÉRE ---
    Dim beszur As String
    beszur = "Sziasztok. <br> <br> Átvettük, rendezésig elzártuk.<br><br>"

    mail.HTMLBody = beszur & mail.HTMLBody
End Sub


Sub ValaszSzoveg_II()
' DoWtHen Makró 2026.10.01
' Válasz Szöbeg Sablonok
' CSAK a levél legelejére tud beszúrni szöveget!
' DeBug

Dim sel As Outlook.Selection
Dim mail As Outlook.MailItem

    ' Kijelölt elem lekérése
    Set sel = Application.ActiveExplorer.Selection
    If sel.Count = 0 Then
        MsgBox "Nincs kijelölt levél.", vbExclamation
        Exit Sub
    End If

    ' Csak MailItem esetén
    If TypeOf sel.Item(1) Is Outlook.MailItem Then
        Set mail = sel.Item(1)
    Else
        MsgBox "Ez nem e-mail.", vbExclamation
        Exit Sub
    End If

    ' --- SZÖVEG BESZÚRÁSA A VÁLASZ ELEJÉRE ---
    Dim beszur As String
    beszur = "Sziasztok. <br> <br> Könyvelve, kiadtam.<br><br>"

    mail.HTMLBody = beszur & mail.HTMLBody
End Sub


Sub ValaszSzoveg_III()
' DoWtHen Makró 2026.10.01
' Válasz Szöbeg Sablonok
' CSAK a levél legelejére tud beszúrni szöveget!
' IQAC

Dim sel As Outlook.Selection
Dim mail As Outlook.MailItem

    ' Kijelölt elem lekérése
    Set sel = Application.ActiveExplorer.Selection
    If sel.Count = 0 Then
        MsgBox "Nincs kijelölt levél.", vbExclamation
        Exit Sub
    End If

    ' Csak MailItem esetén
    If TypeOf sel.Item(1) Is Outlook.MailItem Then
        Set mail = sel.Item(1)
    Else
        MsgBox "Ez nem e-mail.", vbExclamation
        Exit Sub
    End If

    ' --- SZÖVEG BESZÚRÁSA A VÁLASZ ELEJÉRE ---
    Dim beszur As String
    beszur = "Sziasztok. <br> <br> Elhoztuk, könyvelés alatt.<br><br>"

    mail.HTMLBody = beszur & mail.HTMLBody
End Sub


```
