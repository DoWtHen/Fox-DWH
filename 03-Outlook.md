# Fox-DWH Outlook:
## Outlook makró a levelek áthelyezésére:

```vba
Sub UjEmailSablonbol()
    Dim MyItem As Outlook.MailItem
    Set MyItem = Application.CreateItemFromTemplate( _
        "C:\Users\dowth\AppData\Roaming\Microsoft\Templates\dwh.oft")
    MyItem.Display
End Sub
```
```vba
Sub Kijelolt_Emailek_Athelyezese_Tallozas()
' CSAK A KIJELÖLT LEVELEKET MÁSOLJA ÁT MAPPA TALLÓZÁS ABLAKKAL
    
    On Error GoTo ErrHandler

    Dim ns As Outlook.NameSpace
    Dim destFolder As Outlook.MAPIFolder
    Dim itm As Object

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

    MsgBox "Áthelyezés kész.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Hiba történt: " & Err.Description, vbCritical
End Sub


Sub Kijelolt_Emailek_Athelyezese()
' CSAK A KIJELÖLT LEVELEKET MÁSOLJA ÁT

    On Error GoTo ErrHandler

    Dim ns As Outlook.NameSpace
    Dim destFolder As Outlook.MAPIFolder
    Dim itm As Object

    Set ns = Application.GetNamespace("MAPI")

    ' ?? IDE ÍRD A CÉL MAPPÁT
    ' Példa: Inbox › Archiválás
    Set destFolder = ns.GetDefaultFolder(olFolderInbox).Folders("Archiválás")

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Nincs kijelölt elem.", vbExclamation
        Exit Sub
    End If

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
' MINDEN LEVELET ÁTMÁSOL AMI A MAPPÁBAN VAN
    On Error GoTo ErrHandler

    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.MAPIFolder
    Dim destFolder As Outlook.MAPIFolder
    Dim itm As Outlook.MailItem

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' ?? Célmappa
    Set destFolder = inbox.Folders("Archiválás")

    While inbox.Items.Count > 0
        If TypeOf inbox.Items(1) Is Outlook.MailItem Then
            inbox.Items(1).Move destFolder
        Else
            inbox.Items(1).Delete
        End If
    Wend

    MsgBox "Minden levél áthelyezve.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Hiba történt: " & Err.Description, vbCritical
End Sub


Sub Mappanevek_Listaja()
    Dim ns As Outlook.NameSpace
    Dim root As Outlook.MAPIFolder

    Set ns = Application.GetNamespace("MAPI")
    Set root = ns.Folders(1) ' első postafiók

    Debug.Print "=== Mappák listája ==="
    Call PrintFolders(root, "")
End Sub

Sub PrintFolders(ByVal fld As Outlook.MAPIFolder, ByVal indent As String)
    Dim subFld As Outlook.MAPIFolder

    Debug.Print indent & fld.Name

    For Each subFld In fld.Folders
        PrintFolders subFld, indent & "    "
    Next subFld
End Sub
```

