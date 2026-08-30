# Fox-DWH Excel:
<img width="1308" height="298" alt="ikonkepek" src="https://github.com/user-attachments/assets/b5f88e6d-ceff-4586-a8ec-261008ee579f" />


## Excel Personal kódok:

```vba
Sub AutoSzurok_KI_BE()
' DoWtHen Makró 2026.07.17
' Auto Szűrők ki-be kapcsolása

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Ha van szűrés, akkor törli
    If ws.FilterMode Then
        ws.ShowAllData
    End If

    ' Ha nincs AutoFilter a lapon, akkor bekapcsolja
    If ws.AutoFilter Is Nothing Then
        ' Itt állítsd be, melyik soron legyen az AutoFilter
        ws.Range("A1").AutoFilter
    End If
End Sub
```

```vba
Option Explicit

Sub info()
' DoWtHen Makró 2024.12.20

Dim Aktualis, Elso_cella, Utolso_cella, AktualisOszlop_Utolso_cella, ScrollFelfele, ScrollLefele, MindenLap_Legfelulre, AutoSzurok_KI_BE, Kozepre_Igazit, ZoomFel, ZoomLe, Zoom100

Elso_cella = "* Első cella:  az A2 cellára ugrik."
Utolso_cella = "* Utolsó cella:  az A oszlop utolsó használt cellája alá ugrik."
AktualisOszlop_Utolso_cella = "* AktuálisOszlop Utolsó cella:  a kijelölt cella oszlopának utolsó használt cellája alá ugrik."
ScrollFelfele = "* Scroll Felfele:  20 sort ugrik felfelé, kijelöli a cellát is."
ScrollLefele = "* Scroll Lefele:   20 sort ugrik lefelé, kijelöli a cellát is."
MindenLap_Legfelulre = "* MindenLap Legfelülre:  minden munkalapot az A1 cellára görget vissza," & vbCrLf & "de a cella kijelölés az aktuális cellán marad."
AutoSzurok_KI_BE = "* AutoSzűrők ki-be:  ki-be kapcsolja az AutoSzűrőket az A1 cellától kezdve."
Kozepre_Igazit = "* Középre Igazít:  A kijelölt cella-tartomány közepére rendezi a tartalmat."
ZoomFel = "* ZoomFel:  10%-kal növeli a táblázat Nagyítását 160%-ig."
Zoom100 = "* 100%:  100%-ra állítja a képméretet."
ZoomLe = "* ZoomLe:  10%-kal csökkenti a táblázat Nagyításást 50%-ig."
Aktualis = "2025.01.30"

MsgBox Space(36) & "Infók a makrókról 1.rész:" & vbCrLf & Elso_cella & vbCrLf & Utolso_cella & vbCrLf & AktualisOszlop_Utolso_cella & vbCrLf & ScrollFelfele & vbCrLf & _
ScrollLefele & vbCrLf & MindenLap_Legfelulre & vbCrLf & AutoSzurok_KI_BE & vbCrLf & Kozepre_Igazit & vbCrLf & ZoomFel & vbCrLf & Zoom100 & vbCrLf & ZoomLe & vbCrLf & Space(90) & Aktualis, , "Információk"
End Sub

Sub info_2()
' DoWtHen Makró 2025.07.15

Dim Aktualis, Biztonsagi_Mentes

Biztonsagi_Mentes = "* Biztonsági Mentés:  Egy megadott mappába létrehoz a fájlról egy másolatot dátum, idő hozzáadásával a fájlnévhez," & vbCrLf & "a kiterjesztést az eredeti fájlból adja hozzá."
Aktualis = "2025.07.15"

MsgBox Space(36) & "Infók a makrókról 2.rész:" & vbCrLf & Biztonsagi_Mentes & vbCrLf & Space(90) & Aktualis, , "Információk"
End Sub


Sub Elso_cella()
' DoWtHen Makró 2024.12.20
' A2 cellára ugrik

    Range("A2").Select
End Sub


Sub Utolso_cella()
' DoWtHen Makró 2024.12.20
' A oszlop utolsó cellájára ugrik

Dim UtolsoA As Long

UtolsoA = Range("A" & Rows.Count).End(xlUp).Row 'az A oszlop utolsó cella száma
    Range("A" & UtolsoA + 1).Select
End Sub


Sub AktualisOszlop_Utolso_cella()
' DoWtHen Makró 2024.12.20
' Aktuális oszlop utolsó cellájára ugrik

Dim Utolsocella As Long

Utolsocella = Cells(Rows.Count, ActiveCell.Column).End(xlUp).Row

    Cells(Utolsocella + 1, ActiveCell.Column).Select
End Sub


Sub AutoSzurok_KI_BE()
' DoWtHen Makró 2024.12.20
' Auto Szűrők ki-be kapcsolása

    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    
'On Error GoTo Hibasor
'    Range("A1").Select
'    Selection.AutoFilter  'ki-be kapcsolja a szűrőket
'    Exit Sub

'Hibasor:
'MsgBox "Nincsenek a munkalapon Szűrők!"
End Sub


Sub ScrollLefele()
' DoWtHen Makró 2024.12.20
' 20 sort ugrik lefelé, kijelöli a cellát.

    ActiveWindow.SmallScroll Down:=20 'ez a sor scrollozza a táblát
    ActiveCell.Offset(20, 0).Select 'ez a sor kijelöli a cellát
End Sub


Sub ScrollFelfele()
' DoWtHen Makró 2024.12.20
' 20 sort ugrik felfelé, kijelöli a cellát.

On Error GoTo Hibasor
    ActiveWindow.SmallScroll Up:=20 'ez a sor scrollozza a táblát
    ActiveCell.Offset(-20, 0).Select 'ez a sor kijelöli a cellát
    Exit Sub
    
Hibasor:
MsgBox "Nem tudok feljebb lépni!"
End Sub


Sub MindenLap_Legfelulre()
' DoWtHen Makró 2024.12.21
' Minden munkalapon az A1 cellára görgeti vissza a táblázatot, a cella kijelölésen nem változtat.
'forrás: https://wellsr.com/vba/2017/excel/vba-scroll-with-scrollrow-and-scrollcolumn/

Dim ws As Worksheet

  For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    ActiveWindow.ScrollColumn = 1 'oszlop
    'ActiveWindow.ScrollRow = 1 'sor
    ActiveWindow.ScrollRow = 5 'sor
  Next ws
End Sub


Sub Kozepre_Igazit()
' DoWtHen Makró 2024.12.21
' A kijelölt cella tartomány közepére rendezi a tartalmat
'forrás: https://excel-bazis.hu/tutorial/kijeloles-kozepere-makroval

  With Selection
    .HorizontalAlignment = xlCenterAcrossSelection
    .VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    .ShrinkToFit = False
    .MergeCells = False
  End With
End Sub


Sub ZoomFel()
' DoWtHen Makró 2025.01.30
'Forrás: https://www.automateexcel.com/vba/zoom/

Dim x As Integer 'variable for loop
Dim OriginalZoom As Integer 'variable for original zoom

OriginalZoom = ActiveWindow.Zoom 'az aktuális Zoom értéke
'MsgBox OriginalZoom

  If OriginalZoom Mod 10 = 0 Then 'elosztom a számot, ha nincs maradék ez fut le
    'MsgBox "nincs maradék"
    If OriginalZoom < 160 Then
        ActiveWindow.Zoom = OriginalZoom + 10
    Else
        ActiveWindow.Zoom = 80
    End If
  Else 'ha van maradék ez fut le
    OriginalZoom = Application.RoundUp(OriginalZoom, -1) 'felfele kerekítem a számot pl.:58-at 60-ra
    'MsgBox KerekZoom
    ActiveWindow.Zoom = OriginalZoom
    If OriginalZoom < 160 Then
        ActiveWindow.Zoom = OriginalZoom + 10
    Else
        ActiveWindow.Zoom = 80
    End If
  End If
End Sub


Sub ZoomLe()
' DoWtHen Makró 2025.01.30
'Forrás: https://www.automateexcel.com/vba/zoom/

Dim x As Integer 'variable for loop
Dim OriginalZoom As Integer 'variable for original zoom

OriginalZoom = ActiveWindow.Zoom 'az aktuális Zoom értéke
'MsgBox OriginalZoom

  If OriginalZoom Mod 10 = 0 Then 'elosztom a számot, ha nincs maradék ez fut le
    'MsgBox "nincs maradék"
      If OriginalZoom > 50 Then
        ActiveWindow.Zoom = OriginalZoom - 10
      Else
        ActiveWindow.Zoom = 100
      End If
  Else 'ha van maradék ez fut le
    OriginalZoom = Application.RoundUp(OriginalZoom, -1) 'felfele kerekítem a számot pl.:58-at 60-ra
    'MsgBox KerekZoom
    ActiveWindow.Zoom = OriginalZoom
      If OriginalZoom > 50 Then
        ActiveWindow.Zoom = OriginalZoom - 10
      Else
        ActiveWindow.Zoom = 100
      End If
  End If
End Sub


Sub Zoom100()
' DoWtHen Makró 2025.01.30

    ActiveWindow.Zoom = 100
End Sub


Sub Biztonsagi_Mentes()
' DoWtHen Makró 2024.12.22
' Biztonsági mentés a C:\bizment mappába

Dim savedate, savetime, fajlneve, kiterjesztes
Dim formattime As String
Dim formatdate As String
Dim vFn As Variant
Dim menteshelye As String

savedate = Date
savetime = Time
formattime = Format(savetime, "hh.MM")
formatdate = Format(savedate, "YYYY.MM.DD")
fajlneve = ActiveWorkbook.Name
vFn = Split(fajlneve, ".") 'a kiterjesztést keresi meg a teljes fájlnévben
kiterjesztes = vFn(UBound(vFn)) 'a kiterjesztés menti a változóba

menteshelye = "C:\Biz_Ment\"

On Error GoTo Hibasor
    ActiveWorkbook.SaveCopyAs Filename:=menteshelye & ActiveWorkbook.Name & " " & formatdate & "-" & formattime & "." & kiterjesztes
    'ThisWorkbook.SaveCopyAs Filename:=menteshelye & ThisWorkbook.Name & " " & formatdate & "-" & formattime & ".xlsx"
    Exit Sub
    
Hibasor:
MsgBox "Nincs ilyen mappa:  " & menteshelye & vbCrLf & "Hozd létre a mappát vagy változtasd meg a makróban a mappa elérési útvonalat.", vbCritical, "Nincs Mentési Mappa"
End Sub
```

```vba
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


kerdes = MsgBox("Képleteket írok a(z)  " & Cells(ActiveCell.Row, ActiveCell.Column).Address(False, False) & "  cellától!" & vbCrLf & "Mehet??" & Space(15) & "====", vbYesNo + vbQuestion, "Adat másolása")

 If kerdes <> vbYes Then 'ha nem igen kilépek
        MsgBox "Akkor kilépek.", , "Mégsem"
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

```
