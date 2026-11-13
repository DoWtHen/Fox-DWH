# Word makrók
## Doksi_Sablon modul

```vba
Sub WO_Doksi_sablon()
' DoWtHen Makró 2026.11.13
' Foxconn segédlet
' WO Doksi Makró

  If MsgBox(Space(3) & "Szöveget illesztek be a lapra!" & vbCrLf & _
    Space(3) & "--------------------------------" & vbCrLf & vbCrLf & _
    Space(4) & "Biztosan futtatod a makrót?", vbCritical + vbYesNo, "Megerősítés") = vbNo Then
    MsgBox "Akkor kilépek.", vbInformation, "Mégsem"
    Exit Sub
  End If
  
    Selection.TypeText Text:="WO 2870103"
    Selection.TypeParagraph
    Selection.TypeText Text:="Nagyérték" & ChrW(369) & " anyagok,"
    Selection.TypeParagraph
    Selection.TypeText Text:="KIADVA"
    
  If Selection.PageSetup.Orientation = wdOrientPortrait Then
    Selection.PageSetup.Orientation = wdOrientLandscape 'fekvő lap
    Selection.PageSetup.VerticalAlignment = wdAlignVerticalCenter 'középre igazított szöveg
  Else
    Selection.PageSetup.Orientation = wdOrientPortrait
  End If
    
    Selection.WholeStory
    Selection.Font.Size = 82
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Bold = wdToggle
    Selection.Font.Size = 100
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.WholeStory
    Selection.EndKey Unit:=wdStory 'utolsó szó végére ugrik, a dolsi legvégére
End Sub


Sub Torles()
' DoWtHen Makró 2026.11.13
' Dokumentum teljes törlése Makró

  If MsgBox(Space(3) & "A makró törli a teljes dokumentumot!" & vbCrLf & _
    Space(3) & "-----------------------------------------" & vbCrLf & vbCrLf & _
    Space(10) & "Biztosan futtatod a makrót?", vbCritical + vbYesNo, "TÖRLÉS Megerősítés") = vbNo Then
    MsgBox "Akkor kilépek.", vbInformation, "Mégsem"
    Exit Sub
  End If

    Selection.WholeStory

  With ActiveDocument.PageSetup
    .Orientation = wdOrientPortrait
    .VerticalAlignment = wdAlignVerticalTop
  End With
    
    Selection.TypeBackspace
    Selection.Style = ActiveDocument.Styles("Normál")
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
```
