# Fox-DWH Excel:
## Excel Personal kódok:

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

Dim UtolsoA As Long
Dim kerdes As Integer
 
UtolsoA = Range("A" & Rows.Count).End(xlUp).Row  'A oszlop utolsó cella száma


kerdes = MsgBox("Képleteket írok a G1 cellától!" & vbCrLf & "Mehet??", vbYesNo + vbQuestion, "Adat másolása")

If kerdes = vbYes Then
    Range("G1") = "Össz.sor"
    'Range("G2").FormulaLocal = "=DARAB2(A2:A" & UtolsoA & ")" 'magyar verzió
    Range("G2").Formula = "=COUNTA(A2:A" & UtolsoA & ")"

    Range("H1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H1") = "Zöld"
    'Range("H2").FormulaLocal = "=SzinSzamolas(C2:C" & UtolsoA & ";H1)" 'magyar verzió
    'Range("H2").Formula = "=SzinSzamolas(C2:C" & UtolsoA & ",H1)"
    Range("H2").Value = SzinSzamolas(Range("C2:C" & UtolsoA), Range("H1")) 'ez csak eredményt ír be
 
    Range("I1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I1") = "Sárga"
    'Range("I2").FormulaLocal = "=SzinSzamolas(C2:C" & UtolsoA & ";I1)" 'magyar verzió
    'Range("I2").Formula = "=SzinSzamolas(C2:C" & UtolsoA & ",I1)"
    Range("I2").Value = SzinSzamolas(Range("C2:C" & UtolsoA), Range("I1")) 'ez csak eredményt ír be
    
    Range("J1") = "Üres"
    Range("J2") = "=G2-(H2+I2)"
    
Application.Wait (Now + TimeValue("0:00:01")) ' Egy kis szünet
    
    Application.CutCopyMode = False
    Range("H3").Select
    Selection.Style = "Percent"
    ActiveCell.FormulaR1C1 = "=R[-1]C/R2C7"
    Range("H3").Select
    Selection.AutoFill Destination:=Range("H3:J3"), Type:=xlFillDefault
    
    Range("G1:J3").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    Range("G3").Select
Else
    MsgBox "Akkor kilépek.", , "Mégsem"
End If
End Sub
```
