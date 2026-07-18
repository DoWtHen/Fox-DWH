# IC leltár elküldendő adatok
Ezzel már előrébb vagyok...

```vba

Sub Programozott_ICk()
' DoWtHen Makró 2026.07.20
' IC-k lista szűrése programozott SU számokra (nem _00 végű SU számok)
' Copilot szerkesztette

Dim UtolsoA As Long
Dim SzurtUtolsoA As Long
Dim ListaUtolsoA As Long
Dim i As Long ' oszlopszélesség változója
Dim kerdes As Integer
Dim lang As Long 'Windows nyelv keresése
Dim wordYes As String 'Windows nyelv keresése

lang = Application.LanguageSettings.LanguageID(msoLanguageIDUI) 'Nyelvi kódtábla számát adja vissza

    'kódtáblához igazodva írja ki az Igen szót
    Select Case lang
        Case 1033: wordYes = "Yes"
        Case 1038: wordYes = "Igen"
        Case 1031: wordYes = "Ja"
        Case Else: wordYes = "Yes"   ' alapértelmezett
    End Select

kerdes = MsgBox("Ez a makró a Programozott IC-k munkalapot szűri le." & vbCrLf & "Az  " & wordYes & "-re kattintva kezdi a formázást." & vbCrLf & "Mehet??", vbYesNo + vbQuestion, "Adatok szűrése")

 If kerdes <> vbYes Then 'ha a nem-re kattintott kilépek
        MsgBox "Akkor kilépek.", , "Mégsem"
        Exit Sub
 End If

UtolsoA = Range("A" & Rows.Count).End(xlUp).Row  'A oszlop utolsó cella száma

On Error Resume Next
    ' Feltételes szűrés
    Range("I1").Select
    ActiveSheet.Range("$A$1:$I$" & UtolsoA).AutoFilter Field:=9, Criteria1:="<>*_00", Operator:=xlAnd
        
SzurtUtolsoA = Range("A" & Rows.Count).End(xlUp).Row  'A oszlop utolsó cella száma

    Range("A1:I" & SzurtUtolsoA).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ' első sor sortörés és magasság beállítása
    With Rows(1) ' az első sor kiválasztása
        .RowHeight = 45 ' magasság állítás
        .WrapText = True ' sortörés a cellaszövegben
    End With
    
    ' minden oszlop legyen Autofit + 3 pont széles
    For i = 1 To 9
        Columns(i).AutoFit
        Columns(i).ColumnWidth = Columns(i).ColumnWidth + 3
    Next i
    ' újra beállítjuk az első sor magasságát
    ' első sor legyen Autofit + 5 pont magas
    Rows(1).AutoFit
    Rows(1).RowHeight = Rows(1).RowHeight + 5
    
    ' H és I oszlop legyen középre rendezve
    With Columns("H:I") ' a H és I oszlop kiválasztása
        .HorizontalAlignment = xlCenter ' szöveg rendezés középre
    End With
    
ListaUtolsoA = Range("A" & Rows.Count).End(xlUp).Row  'A oszlop utolsó cella száma
    
    ' A legegyszerűbb teljes rácsozás parancs !!!
    Range("A1:I" & ListaUtolsoA).Borders.LineStyle = xlContinuous
    Range("A1:I" & ListaUtolsoA).Borders.Weight = xlThin
    
    Range("A1").Select
Application.Wait (Now + TimeValue("0:00:02")) ' Egy kis szünet
    Call Emailbe_ICk
End Sub


Sub Emailbe_ICk()
' DoWtHen Makró 2026.07.20
' IC-k lista Email generálása
' Copilot szerkesztette

Dim OutApp As Object
Dim OutMail As Object
Dim editor As Object
Dim rg As Range
Dim kerdes As Integer
Dim lang As Long 'Windows nyelv keresése
Dim wordNo As String 'Windows nyelv keresése

lang = Application.LanguageSettings.LanguageID(msoLanguageIDUI) 'Nyelvi kódtábla számát adja vissza

    'kódtáblához igazodva írja ki a Nem szót
    Select Case lang
        Case 1033: wordNo = "No"
        Case 1038: wordNo = "Nem"
        Case 1031: wordNo = "Nein"
        Case Else: wordNo = "No"   ' alapértelmezett
    End Select

kerdes = MsgBox(Space(15) & "Létrehozok egy EMAIL-t," & vbCrLf & "beszúrom a táblázatot és megadom a Címzeteket is! " & vbCrLf & "Mehet  vagy a  " & wordNo & "  gombbal kilépsz??", vbYesNo + vbQuestion, "Adatok szűrése")

 If kerdes <> vbYes Then 'ha a nem-re kattintott kilépek
        MsgBox "Akkor kilépek.", , "Mégsem"
        Exit Sub
 End If

    ' Másolandó tartomány
    Set rg = ActiveSheet.UsedRange
    rg.Copy

    ' Outlook email létrehozása
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    ' Címzettek
    OutMail.To = "valaki@ceg.hu; valaki22@ceg.hu"
    
    ' Email tárgy
    OutMail.Subject = "Programozott IC-k listája  " & Format(Date, "yyyy.mm.dd")

    ' Email megnyitása
    OutMail.Display

    ' Várunk, amíg a WordEditor létrejön
    Do While OutMail.GetInspector.WordEditor Is Nothing
        DoEvents
    Loop

    ' WordEditor elérése
    Set editor = OutMail.GetInspector.WordEditor.Application.Selection

    ' 1) Szöveg beírása
    editor.TypeText "Sziasztok." & vbCrLf & vbCrLf
    editor.TypeText "Küldöm a programozott IC-k táblázatát." & vbCrLf & vbCrLf

    ' 2) Táblázat beillesztése
    editor.Paste
End Sub

```
