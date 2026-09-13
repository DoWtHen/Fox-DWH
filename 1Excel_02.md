```vba

'==================================
'            DoWtHen

' === PRÓBA FÜGGVÉNYEK, MAKRÓK ===
'==================================
Public Highlighter As clsHighlight

Sub ToggleRowHighlight()
' DoWtHen Makró 2026.04.13
' Sorok kiemelése sárga színnel kapcsoló makró része
' Copilot szerkesztette

    'Ha még nincs példány, hozzuk létre
    If Highlighter Is Nothing Then
        Set Highlighter = New clsHighlight
        Set Highlighter.App = Application
        Highlighter.HighlightEnabled = False 'induláskor legyen kikapcsolva, most úgyis váltunk
    End If

    Highlighter.HighlightEnabled = Not Highlighter.HighlightEnabled

    If Highlighter.HighlightEnabled Then
        MsgBox "Sor kiemelés: BEKAPCSOLVA", vbInformation
    Else
        MsgBox "Sor kiemelés: KIKAPCSOLVA", vbExclamation
    End If
End Sub


Function Toldalek(ertek As Variant) As String
' DoWtHen makró
' 2026.08.18
' Toldalék hozzáadása számhoz dátumhoz függvény
' A képlet után Copolit szerkesztette

    Dim nap As Long
    Dim utolso2 As Long
    Dim szoveg As String

    ' Ha dátum
    If IsDate(ertek) Then
        szoveg = Format(ertek, "yyyy.mm.dd")
        nap = Day(ertek)
        utolso2 = nap Mod 100

    ' Ha szám
    ElseIf IsNumeric(ertek) Then
        szoveg = CStr(ertek)
        utolso2 = CLng(ertek) Mod 100

    Else
        Toldalek = "#HIBA"
        Exit Function
    End If

    ' Toldalék meghatározása
    If utolso2 = 12 Or utolso2 = 22 Then
        Toldalek = szoveg & ".-e"
        Exit Function
    End If

    Select Case utolso2
        Case 2, 3, 6, 8, 13, 16, 18, 20, 23, 26, 28, 30
            Toldalek = szoveg & ".-a"
        Case Else
            Toldalek = szoveg & ".-e"
    End Select
End Function


Sub Minden_Tagolas_Osszecsuk()
' DoWtHen makró
' 2026.08.18
' Tagolás Sorok összecsukása

    Dim r As Range

    For Each r In ActiveSheet.UsedRange.Rows 'a ciklus minden soron ellenőrzi hogy van-e ott Tagolás szum kocka -/+
        ' Csak olyan sor, ahol ténylegesen summary (itt te döntöd el: pl. 1-es szint)
        If r.OutlineLevel = 1 Then
            On Error Resume Next
            r.ShowDetail = False
            On Error GoTo 0
        End If
    Next r
        Range("D37").Select
End Sub


Sub Tagolas_Kinyit()
' DoWtHen makró
' 2026.08.18
' Tagolás Sorok kinyitása

    ActiveSheet.Outline.ShowLevels RowLevels:=8 'bármilyen szám ami nagyobb mint a táblázat tagok száma
    Range("F2").Select
End Sub

```
