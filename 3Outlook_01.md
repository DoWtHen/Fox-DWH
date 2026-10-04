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

  'If MsgBox("A kijelölt levelek áthelyezése Tallózás ablakkal." & vbCrLf & "Biztosan futtatod a makrót?", vbQuestion + vbYesNo, "Megerősítés") = vbNo Then
  '  MsgBox "Akkor kilépek."
  '  Exit Sub
  'End If

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

    Dim objSelection As Outlook.Selection
    Dim objItem As Object
    Dim objNamespace As Outlook.NameSpace
    Dim objPstStore As Outlook.Store
    Dim objArchiveFolder As Outlook.MAPIFolder
    Dim targetPstName As String
    Dim targetFolderName As String
    
    If MsgBox("A kijelölt levelek áthelyezése az Archivum mappába." & vbCrLf & _
              "Biztosan futtatod a makrót?", vbQuestion + vbYesNo, "Megerősítés") = vbNo Then
        MsgBox "Akkor kilépek."
        Exit Sub
    End If
    
    ' --- BEÁLLÍTÁSOK ---
    ' Az Outlook oldalsávján megjelenő PST adatfájl pontos neve
    targetPstName = "Archívumok"
    ' A PST fájlon belüli célmappa neve (pl. "Beérkezett üzenetek", "Archívum" vagy "Mappa")
    targetFolderName = "Beérkezett üzenetek"
    ' -------------------
    
    ' Kijelölt elemek lekérése
    Set objSelection = Application.ActiveExplorer.Selection
    
    ' Ellenőrzés, hogy van-e kijelölt elem
    If objSelection.Count = 0 Then
        MsgBox "Nincs kijelölve levél az archiváláshoz!", vbExclamation, "Hiba"
        Exit Sub
    End If
    
    Set objNamespace = Application.GetNamespace("MAPI")
    
    ' Megkeressük a megadott nevű PST adatfájlt a csatolt tárolók között
    On Error Resume Next
    Dim tempStore As Outlook.Store
    For Each tempStore In objNamespace.Stores
        If tempStore.DisplayName = targetPstName Then
            Set objPstStore = tempStore
            Exit For
        End If
    Next tempStore
    On Error GoTo 0
    
    ' Ha nem találja az adatfájlt
    If objPstStore Is Nothing Then
        MsgBox "A(z) '" & targetPstName & "' nevű adatfájl nem található az Outlookban! Kérjük, ellenőrizze a nevet az oldalsávon.", vbCritical, "Hiba"
        Exit Sub
    End If
    
    ' Megkeressük a célmappát a PST fájlon belül
    On Error Resume Next
    Set objArchiveFolder = objPstStore.GetRootFolder.Folders(targetFolderName)
    On Error GoTo 0
    
    ' Ha a megadott mappa nem létezik az adatfájlban, létrehozzuk
    If objArchiveFolder Is Nothing Then
        On Error Resume Next
        Set objArchiveFolder = objPstStore.GetRootFolder.Folders.Add(targetFolderName)
        On Error GoTo 0
    End If
    
    ' Végső ellenőrzés a mappára
    If objArchiveFolder Is Nothing Then
        MsgBox "Nem sikerült elérni vagy létrehozni a(z) '" & targetFolderName & "' mappát!", vbCritical, "Hiba"
        Exit Sub
    End If
    
    ' Levelek áthelyezése hátulról előre haladva
    Dim i As Long
    Dim movedCount As Long
    movedCount = 0
    
    For i = objSelection.Count To 1 Step -1
        Set objItem = objSelection.Item(i)
        If TypeOf objItem Is Outlook.MailItem Then
            objItem.Move objArchiveFolder
            movedCount = movedCount + 1
        End If
    Next i
    
    ' Opcionális visszajelzés az állapotsoron (nem zavar felugró ablakkal)
    Application.ActiveExplorer.ClearSelection
    StatusBar = movedCount & " levél sikeresen áthelyezve a(z) " & targetPstName & " adatfájlba."
End Sub


Sub Minden_Email_Athelyezese()
' DoWtHen Makró 2026.05.01
' v2 2026.09.20
' MINDEN LEVELET ÁTHELYEZ AMI A BEJÖVŐ MAPPÁBAN VAN

    Dim objNamespace As Outlook.NameSpace
    Dim objInboxFolder As Outlook.MAPIFolder
    Dim objPstStore As Outlook.Store
    Dim objArchiveFolder As Outlook.MAPIFolder
    Dim objItem As Object
    Dim targetPstName As String
    Dim targetFolderName As String
    Dim i As Long
    Dim movedCount As Long
    
    ' --- BEÁLLÍTÁSOK ---
    ' Az Outlook oldalsávján megjelenő PST adatfájl pontos neve
    targetPstName = "Archívumok"
    ' A PST fájlon belüli célmappa neve
    targetFolderName = "Beérkezett üzenetek"
    ' -------------------
    
    Set objNamespace = Application.GetNamespace("MAPI")
    
    ' 1. Az aktuális fő Beérkezett üzenetek mappa lekérése
    Set objInboxFolder = objNamespace.GetDefaultFolder(olFolderInbox)
    
    ' Ellenőrzés, hogy van-e benne egyáltalán levél
    If objInboxFolder.Items.Count = 0 Then
        MsgBox "A Beérkezett üzenetek mappa már teljesen üres!", vbInformation, "Információ"
        Exit Sub
    End If
    
    ' 2. Az "Archívumok" nevű PST adatfájl megkeresése
    On Error Resume Next
    Dim tempStore As Outlook.Store
    For Each tempStore In objNamespace.Stores
        If tempStore.DisplayName = targetPstName Then
            Set objPstStore = tempStore
            Exit For
        End If
    Next tempStore
    On Error GoTo 0
    
    ' Ha nem találja az adatfájlt
    If objPstStore Is Nothing Then
        MsgBox "A(z) '" & targetPstName & "' nevű adatfájl nem található az Outlookban! Kérjük, ellenőrizze a nevet a bal oldali sávban.", vbCritical, "Hiba"
        Exit Sub
    End If
    
    ' 3. Célmappa megkeresése vagy létrehozása a PST fájlon belül
    On Error Resume Next
    Set objArchiveFolder = objPstStore.GetRootFolder.Folders(targetFolderName)
    On Error GoTo 0
    
    ' Ha a mappa még nem létezik az adatfájlban, létrehozzuk
    If objArchiveFolder Is Nothing Then
        On Error Resume Next
        Set objArchiveFolder = objPstStore.GetRootFolder.Folders.Add(targetFolderName)
        On Error GoTo 0
    End If
    
    ' Végső ellenőrzés a célmappára
    If objArchiveFolder Is Nothing Then
        MsgBox "Nem sikerült elérni vagy létrehozni a(z) '" & targetFolderName & "' mappát a(z) " & targetPstName & " fájlban!", vbCritical, "Hiba"
        Exit Sub
    End If
    
    ' 4. ÖSSZES levél áthelyezése (hátulról előre haladva a számozásban a hibák elkerülése végett)
    movedCount = 0
    For i = objInboxFolder.Items.Count To 1 Step -1
        Set objItem = objInboxFolder.Items(i)
        
        ' Csak a leveleket mozgatjuk (értekezlet-meghívókat, feladatokat nem, ha esetleg lennének ott)
        If TypeOf objItem Is Outlook.MailItem Then
            objItem.Move objArchiveFolder
            movedCount = movedCount + 1
        End If
    Next i
    
    ' Visszajelzés a sikeres futásról
    MsgBox movedCount & " darab levél sikeresen átmozgatva a(z) '" & targetPstName & "' adatfájlba.", vbInformation, "Kész"
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


Sub TemplatesMappaMegnyitasa()
' DoWtHen Makró 2026.09.12
' Megnyitja a Templates mappát

    Dim path As String
    path = Environ$("APPDATA") & "\Microsoft\Templates\"
    Shell "explorer.exe """ & path & """", vbNormalFocus
End Sub

```
