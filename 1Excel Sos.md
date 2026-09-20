
```vbs
Sub LADA_PN_Lista()
' DoWtHen Makró 2026.08.22
' v2 DoWtHen Makró 2026.09.17
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
    Dim db As Long
    Dim r As Long
    
' Növekvő sorrendbe rakva anyagokat felsorolja melyik ládákban találhatóak
    '=== MEGERŐSÍTÉS ===
    If MsgBox("A ""Láda tartalom"" munkalapon" & vbCrLf & "növekvő sorrendbe rakva az anyagokat" & vbCrLf & "felsorolja melyik ládákban találhatóak." & vbCrLf & vbCrLf & _
              "Ez a makró a  H1  cellától kezd bemásolni adatokat!" & vbCrLf & Space(17) & "=====" & vbCrLf & Space(25) & " Mehet?", vbQuestion + vbYesNo, "Adat másolás  LÁDA PN Lista") = vbNo Then
        MsgBox "Akkor kilépek.", vbCritical, "Mégsem"
        Exit Sub
    End If
    
    '=== ALAP ADATOK ===
    lastRow = Cells(Rows.Count, "B").End(xlUp).row
    
    '=== LISTA KÉSZÍTÉSE ===
    Range("H2:H" & lastRow).Value = Range("B2:B" & lastRow).Value
    Range("H2:H" & lastRow).RemoveDuplicates Columns:=1, Header:=xlNo
    
    '=== RENDEZÉS ===
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("H2:H" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange Range("H2:H" & lastRow)
        .Header = xlNo
        .Apply
    End With
    
    '=== ÜRES CELLÁK TÖRLÉSE ===
    On Error Resume Next
    Range("H2:H" & lastRow).SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    On Error GoTo 0
    
    '=== ÚJ lastPN meghatározása ===
    lastPN = Cells(Rows.Count, "H").End(xlUp).row
    
Application.Wait (Now + TimeValue("0:00:02")) 'egy kis szünet

    '=== ÖSSZES TEKERCS DBSZÁM ===
    Range("I2").Select
    'ActiveCell.FormulaLocal = "=HA(H2="""";"""";SZUMHA(B:B;H2;D:D))"  'magyar verzió
    ActiveCell.Formula = "=IF(H2="""","""",SUMIF(B:B,H2,D:D))"  'angol verzió
    Range("I2").Select
  '  Range("I2").AutoFill Destination:=Range("I2:I" & lastPN), Type:=xlFillDefault 'lemásolja az utolsó celláig
    
    '=== PN LÁDA LISTA ===
    For i = 2 To lastPN
        
        pn = Cells(i, "H").Value
        woList = ""
        
        For j = 2 To lastRow
            
            If Cells(j, "B").Value = pn Then
                
                '=== LÁDA SOR MEGKERESÉSE ===
                k = j
                Do While k > 1 And Cells(k, "A").Value = ""
                    k = k - 1
                Loop
                wo = Cells(k, "A").Value   ' EZ A LÁDA
                
                '=== CSAK HA MÉG NINCS HOZZÁADVA ===
                If InStr(woList, wo) = 0 Then
                    
                '=== SU DARABSZÁM SZÁMOLÁSA ===
                If InStr(woList, wo) = 0 Then
                
                    db = 0
                    r = j   ' ITT: az ANYAG sorából indulunk
                
                    Do While r <= lastRow _
                        And (r = j Or (Cells(r, "A").Value = "" And Cells(r, "B").Value = ""))
                
                        If Cells(r, "C").Value <> "" Then
                            db = db + 1
                        End If
                
                        r = r + 1
                    Loop
                
                    If woList = "" Then
                        woList = wo & " (" & db & " db)"
                    Else
                        woList = woList & ",  " & wo & " (" & db & " db)"
                    End If
                
                  End If
                End If
            End If
            
        Next j
        
        Cells(i, "J").Value = woList
        
    Next i
    
    '=== FEJLÉC ===
    Range("H1") = "Anyagszám"
    Range("I1") = "Összes tekercs száma"
    Range("J1") = "Ezekben a Ládákban találod"
    
    With Range("H1:J1")
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Rows(1).RowHeight = 49.5
    Columns("H:H").ColumnWidth = 18.5
    Columns("J:J").ColumnWidth = 55
    Columns("I:I").ColumnWidth = 8.2
    
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
```
