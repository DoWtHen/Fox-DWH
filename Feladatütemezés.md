# Feladatütemezés
### Feladat létrehozása
Az **Indítás** fülön állítsd be; minden nap 10:00-kor <br>
                                        18:00-kor <br>
                                        02:00-kor<br><br>

A **Műveletek** fülön kattints az **Új...** gombra
**Program**-hoz írd be:
```powershell
powershell.exe
```

**Argumentumhoz** írd be: 
```powershell
-c (New-Object Media.SoundPlayer 'C:\Windows\Media\Alarm03.wav').PlaySync()
```
```powershell
-command "Start-Process 'e:\Programok\Automatikus kijelentkezés\figyelem.png'"
```
