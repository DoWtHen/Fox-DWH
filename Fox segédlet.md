```python

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
    Range("I3").FormulaLocal = "=ÖSSZEFŰZ(I2;""-A"")"
    Range("J2").FormulaLocal = "=HAHIBA(FKERES(ÖSSZEFŰZ(H2;I2);$A$2:$C$675;2;HAMIS);""nincs ilyen tárhely"")"
    Range("J3").FormulaLocal = "=HAHIBA(FKERES(ÖSSZEFŰZ(H3;I3);$A$2:$C$675;2;HAMIS);""nincs ilyen tárhely"")"
    Range("K2").FormulaLocal = "=HAHIBA(FKERES(ÖSSZEFŰZ(H2;I2);$A$2:$C$675;3;HAMIS);""nincs ilyen tárhely"")"
    Range("K3").FormulaLocal = "=HAHIBA(FKERES(ÖSSZEFŰZ(H3;I3);$A$2:$C$675;3;HAMIS);""nincs ilyen tárhely"")"
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
```
