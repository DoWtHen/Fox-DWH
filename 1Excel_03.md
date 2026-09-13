```vba

'====================================
'           SEGÉD KÉPLETEK
'              FoxConn
'====================================

Sub Toltheto_Tarhelyek()
' DoWtHen Makró 2026.04.22
' Foxconn segédlet

Dim kerdes As Integer

kerdes = MsgBox("Képleteket írok a H1 cellától!" & vbCrLf & "Mehet??", vbYesNo + vbQuestion, "Adat másolása")

If kerdes = vbYes Then
    Range("H1") = "#"
    Range("I1") = "Tárhely"
    Range("J1") = "Foglalt tárhely"
    Range("K1") = "Üres tárhely"
    Range("I2") = 0
    Range("H2,H3") = "MP"
    'Range("I3").FormulaLocal = "=ÖSSZEFŰZ(I2;""-A"")"
    'Az ősszefűz függvény angolul CONCATENATE, a fűz függvény CONCAT
    Range("I3") = "=CONCAT(I2,""-A"")"
    'Range("J2").FormulaLocal = "=HAHIBA(FKERES(ÖSSZEFŰZ(H2;I2);$A$2:$C$1000;2;HAMIS);""nincs ilyen tárhely"")"
    Range("J2") = "=IFERROR(VLOOKUP(CONCAT(H2,I2),$A$2:$C$1000,2,FALSE),""nincs ilyen tárhely"")"

    'Range("J3").FormulaLocal = "=HAHIBA(FKERES(ÖSSZEFŰZ(H3;I3);$A$2:$C$1000;2;HAMIS);""nincs ilyen tárhely"")"
    Range("J3") = "=IFERROR(VLOOKUP(CONCAT(H3,I3),$A$2:$C$1000,2,FALSE),""nincs ilyen tárhely"")"
    
    'Range("K2").FormulaLocal = "=HAHIBA(FKERES(ÖSSZEFŰZ(H2;I2);$A$2:$C$1000;3;HAMIS);""nincs ilyen tárhely"")"
    Range("K2") = "=IFERROR(VLOOKUP(CONCAT(H2,I2),$A$2:$C$1000,3,FALSE),""nincs ilyen tárhely"")"
    
    'Range("K3").FormulaLocal = "=HAHIBA(FKERES(ÖSSZEFŰZ(H3;I3);$A$2:$C$1000;3;HAMIS);""nincs ilyen tárhely"")"
    Range("K3") = "=IFERROR(VLOOKUP(CONCAT(H3,I3),$A$2:$C$1000,3,FALSE),""nincs ilyen tárhely"")"
Application.Wait (Now + TimeValue("0:00:01")) ' Egy kis szünet

    Range("H1:K3").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Range("I2").Select
    With Selection.Interior
        .Pattern = xlNone
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
Application.Wait (Now + TimeValue("0:00:01")) ' Egy kis szünet
    
    Range("H1:K3").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Range("H1:K1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    Range("K8") = "Teljesen Üres Tárhelyek"
    Range("I2").Select
Else
    MsgBox "Akkor kilépek"
End If
End Sub


Function SzinSzamolas(rng As Range, colorCell As Range) As Long
' DoWtHen Makró 2026.05.30
' Foxconn segédlet
' Szín számoló függvény
' pl.: beírható a cellába is ha a függvény elérhető
' =SzinSzamolas(C2:C23  ;                    C1)
'            tartomány  ;  a színt tartalmazó cella amit számolni kell
' Copilot segítséggel

    Dim c As Range
    Dim cnt As Long
    
    For Each c In rng
        If c.Interior.Color = colorCell.Interior.Color Then
            cnt = cnt + 1
        End If
    Next c
    SzinSzamolas = cnt
End Function


Sub Aranyok()
' DoWtHen Makró 2026.05.30
' Foxconn segédlet
' Arányszámítás a WO kittingeléshez
' 2026.07.02 -> Flexibilis bárhová helyezhető (még mindig az A és C oszlopból számol)

Dim UtolsoA As Long
Dim kerdes As Integer
Dim AktualCella As Range
Dim destRange As Range
Dim Kijeloles As Range
 
UtolsoA = Range("A" & Rows.Count).End(xlUp).Row  'A oszlop utolsó cella száma

Set AktualCella = ActiveCell 'a kijelölt cella ahova az Arányokat beírja

kerdes = MsgBox("A ""kitting lista"" munkalapon összeszámolja," & vbCrLf & "hogy a WO-ra hány %-nyi" & vbCrLf & "alapanyag van kiadva, könyvelve." & vbCrLf & vbCrLf & vbCrLf & "Képleteket írok a(z)  " & Cells(ActiveCell.Row, ActiveCell.Column).Address(False, False) & "  cellától!" & vbCrLf & "Mehet??" & Space(15) & "====", vbYesNo + vbQuestion, "Adat másolása  Arányok")

 If kerdes <> vbYes Then 'ha nem igen kilépek
        MsgBox "Akkor kilépek.", vbInformation, "Mégsem"
        Exit Sub
 End If
    
    AktualCella.Value = "Össz.sor"
    'Range("G2").FormulaLocal = "=DARAB2(A2:A" & UtolsoA & ")" 'magyar verzió
    ActiveCell.Offset(1, 0).Formula = "=COUNTA(A2:A" & UtolsoA & ")"  'egy sorral lejebb

    ActiveCell.Offset(0, 1).Select  'egy oszloppal jobbra
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(0, 0) = "Zöld"  'ugyan oda
    'Range("H2").FormulaLocal = "=SzinSzamolas(C2:C" & UtolsoA & ";H1)" 'magyar verzió
    'ActiveCell.Offset(1, 0).Value = SzinSzamolas(Range("C2:C" & UtolsoA), Range("H1"))  'egy sorral lejebb  ez csak eredményt ír be
    ActiveCell.Offset(1, 0).Value = SzinSzamolas(Range("C2:C" & UtolsoA), ActiveCell.Offset(0, 0))  'egy sorral lejebb  ez csak eredményt ír be
 
    ActiveCell.Offset(0, 1).Select  'egy oszloppal jobbra
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(0, 0) = "Sárga"  'ugyan oda
    'Range("I2").FormulaLocal = "=SzinSzamolas(C2:C" & UtolsoA & ";I1)" 'magyar verzi
    'ActiveCell.Offset(1, 0).Value = SzinSzamolas(Range("C2:C" & UtolsoA), Range("I1")) 'egy sorral lejebb  ez csak eredményt ír be
    ActiveCell.Offset(1, 0).Value = SzinSzamolas(Range("C2:C" & UtolsoA), ActiveCell.Offset(0, 0)) 'egy sorral lejebb  ez csak eredményt ír be
    
    ActiveCell.Offset(0, 1) = "Üres"  'egy oszloppal jobbra
    'ActiveCell.Offset(1, 1) = "=G2-(H2+I2)"  'egy sorral lejebb és egy oszloppal jobbra
    ActiveCell.Offset(1, 1).Formula = "=" & AktualCella.Offset(1, 0).Address(False, False) & "-(" & AktualCella.Offset(1, 1).Address(False, False) & "+" & AktualCella.Offset(1, 2).Address(False, False) & ")"  'kivonás és összeadás eltolt cellákkal

Application.Wait (Now + TimeValue("0:00:01")) ' Egy kis szünet
    
    Application.CutCopyMode = False
    ActiveCell.Offset(2, -1).Select  'két sorral lejebb és egy oszloppal balra
    Selection.Style = "Percent"

    ActiveCell.Formula = "=" & ActiveCell.Offset(-1, 0).Address(False, False) & "/" & AktualCella.Offset(1, 0).Address(True, True)  'osztás eltolt cellákkal
    
Set destRange = Range(ActiveCell, ActiveCell.Offset(0, 2)) 'tartomány megadása aktiv cellához képest

    ActiveCell.AutoFill Destination:=destRange, Type:=xlFillDefault 'aktív cellától a megadott tartományig kijelölés
    AktualCella.Select
    
Set Kijeloles = Range(AktualCella, ActiveCell.Offset(2, 3)) 'tartomány megadása aktiv cellához képest
    
    Kijeloles.Select  'középre igazítás
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    AktualCella.Select
End Sub


Sub Munkalapok_Atnevezese_Munkanapokra()
' DoWtHen Makró 2026.08.01 eredeti fájl
' DoWtHen Makró 2026.09.08
' Foxconn segédlet
' Munkalapok átnevezése csak munkanapokra
' Copilot szerkesztette

    Dim ws As Worksheet
    Dim KezdoDatum As Date
    Dim BeirtSzoveg As String
    Dim parts() As String
    Dim y As Long, m As Long, d As Long
    Dim i As Long

    '--- dátum bekérése ---
    BeirtSzoveg = InputBox( _
        "A ""Műszakjelentés"" munkafüzet MINDEN lapjának átnevezése MUNKANAPOKRA!" & vbCrLf & vbCrLf & _
        "Írd be a kezdő dátumot (pl. 2026.08.01)." & vbCrLf & _
        "Ez lesz az első munkalap neve.", _
        "Átnevezés – Munkanapok", Format(Date, "yyyy.mm.dd"))

    If BeirtSzoveg = "" Then
        MsgBox "Nem adtál meg dátumot. Kilépek.", vbInformation, "Mégsem"
        Exit Sub
    End If

    '--- yyyy.mm.dd feldarabolása ---
    parts = Split(BeirtSzoveg, ".")
    If UBound(parts) <> 2 Then
        MsgBox "Érvénytelen formátum!" & vbCrLf & "Használd így: 2026.08.01", vbCritical
        Exit Sub
    End If

    '--- év, hónap, nap számokká alakítása ---
    y = CLng(parts(0))
    m = CLng(parts(1))
    d = CLng(parts(2))

    '--- dátum összeállítása ---
    On Error Resume Next
    KezdoDatum = DateSerial(y, m, d)
    If Err.Number <> 0 Then
        MsgBox "Érvénytelen dátumérték!", vbCritical
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    '--- munkalapok átnevezése csak munkanapokra ---
    Dim AktDatum As Date
    AktDatum = KezdoDatum

    For i = 1 To ActiveWorkbook.Worksheets.Count
        '--- ha hétvége, léptess tovább hétfőre ---
        Do While Weekday(AktDatum, vbMonday) > 5   ' 6=szombat, 7=vasárnap
            AktDatum = AktDatum + 1
        Loop

        Set ws = ActiveWorkbook.Worksheets(i)

        On Error Resume Next
        ws.Name = Format(AktDatum, "yyyy.mm.dd")

        If Err.Number <> 0 Then
            MsgBox "Nem sikerült átnevezni a(z) " & ws.Name & _
                   " lapot › " & Format(AktDatum, "yyyy.mm.dd")
            Err.Clear
        End If
        On Error GoTo 0

        '--- következő munkanap ---
        AktDatum = AktDatum + 1
    Next i

    MsgBox "Kész! A munkalapok átnevezése munkanapokra megtörtént.", vbInformation
End Sub

```
