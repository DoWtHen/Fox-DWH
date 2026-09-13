```vba

'====================================
'       LÁDA SZÁMOLÁS KÉPLETEK
'              FoxConn
'====================================

Sub LadaKepletek()
' DoWtHen Makró 2026.08.21
' Foxconn segédlet
' Láda képletek bemásolása a Kittingelős munkalapra
' Megmutatja hány tekercs van az adott ládában és összesen a WO-hoz
' Copolit szerkesztette

Dim UtolsoC As Long
Dim kerdes As Integer
Dim AktualCella As Range
Dim destRange As Range
Dim Kijeloles As Range
 
UtolsoC = Range("C" & Rows.Count).End(xlUp).Row  'A oszlop utolsó cella száma

    Range("D1").Select 'bár ez a makró Offset-et használ, és az aktuális cellától másolja a szöveget képleteket, mégis megadom a G1 induló cellát a hibák elkerülése miatt.
Set AktualCella = ActiveCell 'a kijelölt cella ahova beír a makró

kerdes = MsgBox("A ""LÁDA tartalom"" munkalapra anyagszám/DBszám" & vbCrLf & "összesítő képleteit másolja be a D1 cellától." & vbCrLf & vbCrLf & vbCrLf & "Ez a makró a  " & Cells(ActiveCell.Row, ActiveCell.Column).Address(False, False) & " cellától kezd bemásolni adatokat!" & vbCrLf & Space(17) & "=====" & vbCrLf & Space(25) & " Mehet?", vbYesNo + vbQuestion, "Adat másolása  LÁDA képletek")

 If kerdes <> vbYes Then 'ha nem igen kilépek
        MsgBox "Akkor kilépek.", vbCritical, "Mégsem"
        Exit Sub
 End If
    
    AktualCella.Value = "Hány DB SU van ebből az anyagból a ládában"
    ActiveCell.ColumnWidth = 20
    ActiveCell.WrapText = True
    'ActiveCell.Offset(1, 0).FormulaLocal = "=HA(B2="""";"""";DARAB2(C3:INDEX(C:C;HOL.VAN("" * "";B3:B$5026;0)+SOR(C3)-1)))"  'magyar verzió egy sorral lejebb
    ActiveCell.Offset(1, 0).Formula = "=IF(B2="""","""",COUNTA(C3:INDEX(C:C,MATCH(""*"",B3:B$5026,0)+ROW(C3)-1)))"  'egy sorral lejebb
    ActiveCell.Offset(0, 1).Select  'egy oszloppal jobbra
    
    ActiveCell.Offset(0, 0) = "Összesen ennyi tekercs van a WO-ra kiadva ebből az anyagból"  'ugyan oda
    ActiveCell.ColumnWidth = 27.1
    ActiveCell.WrapText = True
    'ActiveCell.Offset(1, 0).FormulaLocal = "=HA(B2="""";"""";SZUMHA(B:B;B2;D:D))"  'magyar verzió egy sorral lejebb
    ActiveCell.Offset(1, 0).Formula = "=IF(B2="""","""",SUMIF(B:B,B2,D:D))"  'egy sorral lejebb
    ActiveCell.Offset(0, 1).Select  'egy oszloppal jobbra
    
    ActiveCell.Offset(0, 0) = "Hány ládában van ez az anyag?"  'ugyan oda
    ActiveCell.ColumnWidth = 15
    ActiveCell.WrapText = True
    'ActiveCell.Offset(1, 0).FormulaLocal = "=HA(B2="""";"""";DARABHA(B:B;B2))"  'magyar verzió egy sorral lejebb
    ActiveCell.Offset(1, 0).Formula = "=IF(B2="""","""",COUNTIF(B:B,B2))"  'egy sorral lejebb
    
    ActiveCell.Offset(1, -2).Select  'egysorral lejebb és balra kettőt
    Range(ActiveCell, ActiveCell.Offset(0, 2)).Select  'kijelöli az első + két cellát
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(UtolsoC, ActiveCell.Column + 2)), Type:=xlFillDefault  'lemásolja az utolsó celláig
    
    Range("G1").Value = ".": Range("G1").Select  'egysorban de ez két művelet
End Sub


Sub LADA_PN_Lista()
' DoWtHen Makró 2026.08.22
' Foxconn segédlet
' Anyagszám rendezése abc-be, Láda számok hozzáerendelés az anyagszámokhoz
' Melyik ládákban van az adott anyag
' Copolit szerkesztette

    Dim lastRow As Long
    Dim lastPN As Long
    Dim i As Long, j As Long
    Dim pn As String
    Dim wo As String
    Dim woList As String
    Dim k As Long
    
' Növekvő sorrendbe rakva anyagokat felsorolja melyik ládákban találhatóak
    '=== MEGERŐSÍTÉS ===
    If MsgBox("A ""Láda tartalom"" munkalapon" & vbCrLf & "növekvő sorrendbe rakva az anyagokat" & vbCrLf & "felsorolja melyik ládákban találhatóak." & vbCrLf & vbCrLf & vbCrLf & "Ez a makró a  H1  cellától kezd bemásolni adatokat!" & vbCrLf & Space(17) & "=====" & vbCrLf & Space(25) & " Mehet?", vbQuestion + vbYesNo, "Adat másolás  LÁDA PN Lista") = vbNo Then
        MsgBox "Akkor kilépek.", vbCritical, "Mégsem"
        Exit Sub
    End If
    
    lastRow = Cells(Rows.Count, "B").End(xlUp).Row
    
    '=== LISTA KÉSZÍTÉSE ===
    Range("B2:B" & lastRow).Copy
    Range("H2").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ActiveSheet.Range("$H$2:$H" & lastRow).RemoveDuplicates Columns:=1, Header:=xlNo

    '=== RENDEZÉS ===
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("H2:H" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        .SetRange Range("H2:H" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '=== ÜRES CELLÁK TÖRLÉSE (EZ KELL A HIBA MEGSZŰNÉSÉHEZ!) ===
    On Error Resume Next
    Range("H2:H" & lastRow).SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    On Error GoTo 0
    
    '=== ÚJ lastPN meghatározása ===
    lastPN = Cells(Rows.Count, "H").End(xlUp).Row
    
Application.Wait (Now + TimeValue("0:00:02")) 'egy kis szünet

    '=== ÖSSZES TEKERCS DBSZÁM ===
    Range("I2").Select
    'ActiveCell.FormulaLocal = "=HA(H2="""";"""";SZUMHA(B:B;H2;D:D))"  'magyar verzió
    ActiveCell.Formula = "=IF(H2="""","""",SUMIF(B:B,H2,D:D))"  'angol verzió
    Range("I2").Select
    Range("I2").AutoFill Destination:=Range("I2:I" & lastPN), Type:=xlFillDefault 'lemásolja az utolsó celláig
    
    '=== PN LÁDA LISTA ===
    For i = 2 To lastPN
        pn = Cells(i, "H").Value
        woList = ""

        For j = 2 To lastRow

            If Cells(j, "B").Value = pn Then

                ' WO visszakeresése felfelé
                k = j
                Do While k > 1 And Cells(k, "A").Value = ""
                    k = k - 1
                Loop
                wo = Cells(k, "A").Value

                ' WO hozzáadása a listához
                If InStr(woList, wo) = 0 Then
                    If woList = "" Then
                        woList = wo
                    Else
                        woList = woList & ",  " & wo
                    End If
                End If

            End If
        Next j

        Cells(i, "J").Value = woList
    Next i
    
    ' === FEJLÉC ===
    Range("H1") = "Anyagszám"
    Range("I1") = "Összes tekercs száma"
    Range("J1") = "Ezekben a Ládákban találod"
    Range("H1:J1").Select
    
    Selection.Font.Bold = True
    Selection.Font.Size = 18
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Rows(1).RowHeight = 49.5
    Columns("H:H").ColumnWidth = 18.5
    Columns("J:J").ColumnWidth = 55
    Columns("I").ColumnWidth = 8.2
    
    Range("I1").Select
    With Selection
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Font.Bold = False
        .WrapText = True
    End With
    With Columns("I")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Range("G1").Value = ".": Range("G1").Select  'egysorban de ez két művelet
End Sub


Sub Kitting_Lista_DBszam()
' DoWtHen Makró 2026.08.29
' Foxconn segédlet
' WO-ra kiadott mennyiségek beillesztése értékként és SZínezi a cellát
' Copolit segítet

Dim UtolsoA As Long
Dim rng As Range
Dim c As Range
Dim vanLap As Boolean

If MsgBox("A ""kitting lista"" munkalapon az H oszloptól bemásolt LX02-es listából átírja a C oszlopba a WO-ra kikönyvelt DBszámokat." & vbCrLf & "A tételeket szinezi zöld, sárga, piros színnel." & vbCrLf & vbCrLf & vbCrLf & "Ez a makró a  C2  cellától kezd bemásolni adatokat!" & vbCrLf & Space(17) & "=====" & vbCrLf & Space(25) & " Mehet?", vbQuestion + vbYesNo, "Adat másolás  Kitting Lista DBszám") = vbNo Then
    MsgBox "Akkor kilépek", vbCritical, "Mégsem"
    Exit Sub
End If

If ActiveSheet.Name <> "kitting lista" Then
    MsgBox "Nem a Kitting Lista lapon vagy!", vbCritical, "Nem jó munkalap!"
    Exit Sub
End If

UtolsoA = Range("A" & Rows.Count).End(xlUp).Row  'A oszlop utolsó cella száma

    Range("C2").Select
    'ActiveCell.FormulaLocal = "=XKERES($A2;$H:$H;$J:$J;"""")"  'magyar verzió
    ActiveCell.Formula = "=XLOOKUP($A2,$H:$H,$J:$J,"""")"  'angol verzió
    'Range("C2").Select
    Range("C2").AutoFill Destination:=Range("C2:C" & UtolsoA), Type:=xlFillDefault 'lemásolja az utolsó celláig
    Range("C2:C" & UtolsoA).Value = Range("C2:C" & UtolsoA).Value 'csak értékekre váltás

Set rng = Range("C2:C" & UtolsoA) 'C2-től indulunk, és lefelé megyünk a C oszlopban

    For Each c In rng
        If c.Value = "" Then 'ha üres piros
            c.Interior.Color = RGB(255, 0, 0)
        ElseIf c.Value >= c.Offset(0, -1).Value Then 'ha nagyobb vagy egyenlő zöld
            c.Interior.Color = RGB(0, 176, 80)
        Else 'ha kisebb sárga
            c.Interior.Color = RGB(255, 255, 0)
        End If
    Next c
End Sub

```
