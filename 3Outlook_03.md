```vba
'====================================
'       VÁLASZ SZÖVEG SABLONOK
'         FoxConn és egyebek
'====================================

Sub ValaszSzoveg_I()
' DoWtHen Makró 2026.10.01
' Válasz Szöveg Sablonok
' IC tesztesek

Dim objDoc As Object
Dim objSel As Object

On Error Resume Next
    ' 1. ESET: Ha a levél külön ablakban van megnyitva
  If Not Application.ActiveInspector Is Nothing Then
    Set objDoc = Application.ActiveInspector.WordEditor
    ' 2. ESET: Ha a levél a főablakba van beágyazva (Inline Response)
  ElseIf Not Application.ActiveExplorer Is Nothing Then
    Set objDoc = Application.ActiveExplorer.ActiveInlineResponseWordEditor
  End If
    ' Szöveg beillesztése, ha sikerült elérni a szerkesztőt
  If Not objDoc Is Nothing Then
    Set objSel = objDoc.Windows(1).Selection
    objSel.TypeText "Sziasztok." & vbCrLf & vbCrLf & "Átvettük, rendezésig elzártuk."
  Else
    MsgBox "Nem található aktív e-mail szerkesztő! Győződj meg róla, hogy épp írsz egy levelet.", vbExclamation, "Hiba"
  End If
On Error GoTo 0
End Sub


Sub ValaszSzoveg_II()
' DoWtHen Makró 2026.10.01
' Válasz Szöbeg Sablonok
' DeBug

Dim objDoc As Object
Dim objSel As Object

On Error Resume Next
    ' 1. ESET: Ha a levél külön ablakban van megnyitva
  If Not Application.ActiveInspector Is Nothing Then
    Set objDoc = Application.ActiveInspector.WordEditor
    ' 2. ESET: Ha a levél a főablakba van beágyazva (Inline Response)
  ElseIf Not Application.ActiveExplorer Is Nothing Then
    Set objDoc = Application.ActiveExplorer.ActiveInlineResponseWordEditor
  End If
    ' Szöveg beillesztése, ha sikerült elérni a szerkesztőt
  If Not objDoc Is Nothing Then
    Set objSel = objDoc.Windows(1).Selection
    objSel.TypeText "Sziasztok." & vbCrLf & vbCrLf & "Könyvelve, kiadtam."
  Else
    MsgBox "Nem található aktív e-mail szerkesztő! Győződj meg róla, hogy épp írsz egy levelet.", vbExclamation, "Hiba"
  End If
On Error GoTo 0
End Sub


Sub ValaszSzoveg_III()
' DoWtHen Makró 2026.10.01
' Válasz Szöveg Sablonok
' IQAC

Dim objDoc As Object
Dim objSel As Object

On Error Resume Next
    ' 1. ESET: Ha a levél külön ablakban van megnyitva
  If Not Application.ActiveInspector Is Nothing Then
    Set objDoc = Application.ActiveInspector.WordEditor
    ' 2. ESET: Ha a levél a főablakba van beágyazva (Inline Response)
  ElseIf Not Application.ActiveExplorer Is Nothing Then
    Set objDoc = Application.ActiveExplorer.ActiveInlineResponseWordEditor
  End If
    ' Szöveg beillesztése, ha sikerült elérni a szerkesztőt
  If Not objDoc Is Nothing Then
    Set objSel = objDoc.Windows(1).Selection
    objSel.TypeText "Sziasztok." & vbCrLf & vbCrLf & "Elhoztuk, könyvelés alatt."
  Else
    MsgBox "Nem található aktív e-mail szerkesztő! Győződj meg róla, hogy épp írsz egy levelet.", vbExclamation, "Hiba"
  End If
On Error GoTo 0
End Sub

```
