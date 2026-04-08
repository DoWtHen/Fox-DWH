# Fox-DWH Word:
## Excel Personal kódok:

```vba
Sub negyzetmeter()
      ' négyzetméter Makró
      Selection.TypeText Text:="m"
      Selection.InsertSymbol Font:="+Szövegtörzs", CharacterNumber:=178, Unicode:=True
      Selection.TypeText Text:=" "
End Sub


Sub kobmeter()
      ' köbméter Makró
      Selection.TypeText Text:="m"
      Selection.InsertSymbol Font:="+Szövegtörzs", CharacterNumber:=179, Unicode:=True
      Selection.TypeText Text:=" "
End Sub


Sub celsiusfok()
      ' Celsiusfok Makró
      Selection.TypeText Text:="C"
      Selection.InsertSymbol Font:="+Szövegtörzs", CharacterNumber:=176, Unicode:=True
      Selection.TypeText Text:=" "
End Sub
```
