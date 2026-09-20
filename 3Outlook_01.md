# Fox-DWH Outlook:
## Outlook makró a levelek áthelyezésére:

```vba
Sub Kijelolt_Emailek_Athelyezese_Tallozas()
' DoWtHen Makró 2026.05.01
' CSAK A KIJELÖLT LEVELEKET MÁSOLJA ÁT MAPPA TALLÓZÁS ABLAKKAL
    
    On Error GoTo ErrHandler

    Dim ns As Outlook.NameSpace
    Dim destFolder As Outlook.MAPIFolder
    Dim itm As Object

  If MsgBox("A kijelölt levelek áthelyezése Tallózás ablakkal." & vbCrLf & "Biztosan futtatod a makrót?", vbQuestion + vbYesNo, "Megerősítés") = vbNo Then
    MsgBox "Akkor kilépek."
    Exit Sub
  End If

    Set ns = Application.GetNamespace("MAPI")

    ' --- Mappa tallózó ablak megnyitása ---
    Set destFolder = ns.PickFolder
    If destFolder Is Nothing Then
        MsgBox "Nincs kiválasztott célmappa.", vbExclamation
        Exit Sub
    End If

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Nincs kijelölt elem.", vbExclamation
        Exit Sub
    End If

    For Each itm In Application.ActiveExplorer.Selection
        If TypeOf itm Is Outlook.MailItem Then
            itm.Move destFolder
        End If
    Next itm

    MsgBox "Áthelyezés kész!", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Hiba történt: " & Err.Description, vbCritical
End Sub


Sub Kijelolt_Emailek_Athelyezese()
' DoWtHen Makró 2026.05.01
' v2 2026.09.20
' CSAK A KIJELÖLT LEVELEKET MÁSOLJA ÁT

    On Error GoTo ErrHandler

    Dim ns As Outlook.NameSpace
    Dim root As Outlook.MAPIFolder
    Dim destFolder As Outlook.MAPIFolder
    Dim itm As Object

    If MsgBox("A kijelölt levelek áthelyezése az Archivum mappába." & vbCrLf & _
              "Biztosan futtatod a makrót?", vbQuestion + vbYesNo, "Megerősítés") = vbNo Then
        MsgBox "Akkor kilépek."
        Exit Sub
    End If

    Set ns = Application.GetNamespace("MAPI")

    ' --- POSTAFIÓK GYÖKERÉNEK MEGTALÁLÁSA ---
    Set root = ns.Folders("dowthen@freemail.hu")

    If root Is Nothing Then
        MsgBox "Nem találom a postafiókot.", vbCritical
        Exit Sub
    End If

    ' --- ARCHIVUM MAPPA MEGTALÁLÁSA ---
    On Error Resume Next
    Set destFolder = root.Folders("Archivum")
    On Error GoTo ErrHandler

    If destFolder Is Nothing Then
        MsgBox "Nem található az 'Archivum' mappa.", vbCritical
        Exit Sub
    End If

    ' --- KIJELÖLT ELEMEK ELLENŐRZÉSE ---
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Nincs kijelölt elem.", vbExclamation
        Exit Sub
    End If

    ' --- ÁTHELYEZÉS ---
    For Each itm In Application.ActiveExplorer.Selection
        If TypeOf itm Is Outlook.MailItem Then
            itm.Move destFolder
        End If
    Next itm

    MsgBox "Áthelyezés kész.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Hiba történt: " & Err.Description, vbCritical
End Sub


Sub Minden_Email_Athelyezese()
' DoWtHen Makró 2026.05.01
' v2 2026.09.20
' MINDEN LEVELET ÁTHELYEZ AMI A BEJÖVŐ MAPPÁBAN VAN

    On Error GoTo ErrHandler

    Dim ns As Outlook.NameSpace
    Dim root As Outlook.MAPIFolder
    Dim inbox As Outlook.MAPIFolder
    Dim destFolder As Outlook.MAPIFolder
    Dim itm As Object

    If MsgBox("Minden levél áthelyezése az Archivum mappába." & vbCrLf & _
              "Biztosan futtatod a makrót?", vbQuestion + vbYesNo, "Megerősítés") = vbNo Then
        MsgBox "Akkor kilépek."
        Exit Sub
    End If

    Set ns = Application.GetNamespace("MAPI")

    ' --- POSTAFIÓK GYÖKERÉNEK MEGTALÁLÁSA ---
    Set root = ns.Folders("dowthen@freemail.hu")

    If root Is Nothing Then
        MsgBox "Nem találom a postafiókot.", vbCritical
        Exit Sub
    End If

    ' --- ARCHIVUM MAPPA MEGTALÁLÁSA ---
    On Error Resume Next
    Set destFolder = root.Folders("Archivum")
    On Error GoTo ErrHandler

    If destFolder Is Nothing Then
        MsgBox "Nem található az 'Archivum' mappa.", vbCritical
        Exit Sub
    End If

    ' --- BEJÖVŐ MAPPAK MEGTALÁLÁSA ---
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' --- ÁTHELYEZÉS ---
    While inbox.Items.Count > 0
        Set itm = inbox.Items(1)

        If TypeOf itm Is Outlook.MailItem Then
            itm.Move destFolder
        Else
            itm.Delete
        End If
    Wend

    MsgBox "Minden levél áthelyezve.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Hiba történt: " & Err.Description, vbCritical
End Sub



Sub PrintFolders(ByVal fld As Outlook.MAPIFolder, ByVal indent As String)
' DoWtHen Makró 2026.05.01
' Az Immediate ablakban sorolja fel az Outlook mappaneveket
' Ez a függvény része!

    Dim subFld As Outlook.MAPIFolder

    Debug.Print indent & fld.Name

    For Each subFld In fld.Folders
        PrintFolders subFld, indent & "    "
    Next subFld
End Sub


Sub Mappanevek_Listaja()
' DoWtHen Makró 2026.05.01
' Az Immediate ablakban sorolja fel az Outlook mappaneveket

    Dim ns As Outlook.NameSpace
    Dim root As Outlook.MAPIFolder

    Set ns = Application.GetNamespace("MAPI")
    Set root = ns.Folders(1) ' első postafiók

    Debug.Print "=== Mappák listája ==="
    Call PrintFolders(root, "")
End Sub


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


Sub TemplatesMappaMegnyitasa()
' DoWtHen Makró 2026.09.12
' Megnyitja a Templates mappát

    Dim path As String
    path = Environ$("APPDATA") & "\Microsoft\Templates\"
    Shell "explorer.exe """ & path & """", vbNormalFocus
End Sub
```
