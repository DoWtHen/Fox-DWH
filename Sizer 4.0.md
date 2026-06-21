# Sizer 4.0 Makrók - www.brianapps.net
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2026.06.21

### Makrók az ablakok elrendezéséhez.
Használatuknál be kell írni a program nevét(```program_neve```)
 kiterjesztés nélkül pl.:``` notepad ```.<br>
A leírást csak másold be a macro létrehozása ablakba, nevezd el és próbálgasd, módosítsd.

<img width="386" height="383" alt="Macro" src="https://github.com/user-attachments/assets/cdae53af-8b7c-4a66-a6a8-4c1c47848d26" />

# 4 ablak elhelyezése:
3 ablak bal oldalt egymás alatt, a 4. ablak a képernyő közepe körül lesz.<br>
Minden ablak ``` w = 385 h = 460 ``` méretű lesz.

<img width="1048" height="751" alt="4ablak" src="https://github.com/user-attachments/assets/514d69be-0e74-4879-a378-7e072f4009d0" />

```xml
resize : proc[program_neve]*    l = 50 * 
(($_index-2)*($_index-3)*($_index-4)/-6) + 50 * 
(($_index-1)*($_index-3)*($_index-4)/2)  + 50 * 
(($_index-1)*($_index-2)*($_index-4)/-2) + 1580 * 

(($_index-1)*($_index-2)*($_index-3)/6) t = 10 * 
(($_index-2)*($_index-3)*($_index-4)/-6)  + 550 * 
(($_index-1)*($_index-3)*($_index-4)/2)   + 1020 * 
(($_index-1)*($_index-2)*($_index-4)/-2)  + 600 * 

(($_index-1)*($_index-2)*($_index-3)/6) w = 385 h = 460
```

# 3 ablak elhelyezése:
2 ablak bal oldalt egymás alatt, a 3. ablak jobb felül lesz.<br>
Minden ablak ``` w = 1000 h = 900 ``` méretű lesz.

<img width="1024" height="900" alt="3ablak" src="https://github.com/user-attachments/assets/06f8018b-5bea-41c0-80e7-38c2ae41593d" />


```xml
resize : proc[program_neve]* l = 50
* (($_index-2)*($_index-3)/2) + 50
* (($_index-1)*($_index-3)/-1) + 1050 

* (($_index-1)*($_index-2)/2) t = 0 
* (($_index-2)*($_index-3)/2) + 905
* (($_index-1)*($_index-3)/-1) + 0 

* (($_index-1)*($_index-2)/2) w = 1000 h = 900
```

# 2 ablak elhelyezése:
Bal oldalt egy ablak, jobb oldalt egy ablak lesz.<br>
Minden ablak ``` w = 1000 h = 900 ``` méretű lesz.

<img width="1471" height="455" alt="2ablak" src="https://github.com/user-attachments/assets/6d46f0c4-4d3a-4a19-bdf7-15181946a92d" />

```xml
resize : proc[program_neve]* l = (($_index - 1) % 2) * (0.5 * w:workarea) t = 0 w = 1000 h = 900
```

# 2 vagy több ablak elhelyezése fűggőlegesen:
Bal oldalt kezdve arányosan eltolva teszi a következő ablakot egymás mellé.<br>
Minden ablak ``` w = 1000 h = 900 ``` méretű lesz.

<img width="1919" height="464" alt="2vagytobb_ablak" src="https://github.com/user-attachments/assets/56dc528b-ec7a-4024-ba08-223ec35c63dd" />

```xml
resize : proc[program_neve]* l = 50 * (2 - $_index) + 1095 * ($_index - 1) t = 10 w = 1000 h = 900
```
