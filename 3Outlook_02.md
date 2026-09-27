```vba
'====================================
'           LEVÉL SABLONOK
'         FoxConn és egyebek
'====================================

Sub UjEmailSablonbol_1()
' DoWtHen Makró 2026.05.01
' Sablon levél fájl megnyítása

    Dim MyItem As Outlook.MailItem
    Dim path As String
    Dim fajlNev As String

fajlNev = "dwh.oft"
path = Environ$("APPDATA") & "\Microsoft\Templates\" & fajlNev
Set MyItem = Application.CreateItemFromTemplate(path)

    MyItem.Display
End Sub


Sub UjEmailSablonbol_2()
' DoWtHen Makró 2026.05.01
' Sablon levél fájl megnyítása

    Dim MyItem As Outlook.MailItem
    Dim path As String
    Dim fajlNev As String

fajlNev = "proba2.oft"
path = Environ$("APPDATA") & "\Microsoft\Templates\" & fajlNev
Set MyItem = Application.CreateItemFromTemplate(path)
    
    MyItem.Display
End Sub


Sub UjEmailFoxconn()
' DoWtHen Makró 2026.05.01
' Sablon levél fájl megnyítása

    Dim MyItem As Outlook.MailItem
    Dim path As String
    Dim fajlNev As String

fajlNev = "foxconn.oft"
path = Environ$("APPDATA") & "\Microsoft\Templates\" & fajlNev
Set MyItem = Application.CreateItemFromTemplate(path)

    MyItem.Display
End Sub

```
