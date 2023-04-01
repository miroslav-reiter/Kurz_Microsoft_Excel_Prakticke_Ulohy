
# 📊 Materiály ku kurzu Microsoft Excel - Praktické Úlohy


[Kurz Microsoft Excel - Praktické Úlohy](https://www.it-academy.sk/kurz/microsoft-excel-prakticke-ulohy/) je vhodný pre teba pokiaľ chceš pomocou Excelu riešiť praktické úlohy každodenného života. Všetko, čo potrebujete o Exceli vedieť pre riešenie svojich úloh, sa naučíte na praktických príkladoch. Dozvieš sa, že sa dá Excel využiť nielen na tabuľky a grafy, ale aj na hľadanie najlepšieho (optimálneho) riešenie problémov v praktickom živote. Na kurze nájdeš návod, ako riešiť pomocou Excelu svoj praktický problém. Ukážeme si od tých najjednoduchších príkladov a postupne sa prepracujeme k zložitejším. Tento kurz sa realizuje na mieru požiadavkám zákazníka.

## ❓ Čo je to Microsoft Excel
Microsoft Excel je **tabuľkový procesor od** spoločnosti Microsoft navrhnutý pre operačný systém **Microsoft Windows**, **Mac Os**, **Android** a **iO**S. Je súčasťou **kancelárskeho balíka Microsoft Office** spolu s aplikáciami Microsoft Word, Microsoft PowerPoint, Microsoft Outlook, Microsoft Access atď.

## 🙋 Verzie a edície Microsoft Excel
Najaktuálnešia/najnovšia verzia je **Microsoft Excel 365 (Office 365)**. Na trhu sú aj standalone verzie: 2000, 2002, 2003, 2007, 2010, 2013, 2016, 2019

**TIP:** Verzie zistíme na Karte Domov (Home) > Konto (Account) > Čo je Excel
![verzia](https://user-images.githubusercontent.com/24510943/212565132-3a9892b7-d660-4e8e-b883-45794a06fc50.png)


## ⚓ Odkazy na kurzy
[Prezenčné Kurzy Microsoft Excel](https://www.it-academy.sk/kategoria/kancelarske-baliky/kurzy-excel/)  
[Online Kurzy Microsoft Excel](https://www.vita.sk/?s=excel)  

## 📁 Súbory a Materiály
Dostupné na GitHube alebo na kurze od lektora

## 🧰 Stránky a nástroje na precvičovanie Microsoft Excel
1. [Microsoft 365](https://www.microsoft.com/sk-sk/microsoft-365/excel)
2. [ASAP Utilities](http://www.asap-utilities.com/excel-tips-shortcuts.php)
3. [Microsoft Excel Alza Návod](https://www.alza.sk/microsoft-excel-navod)
4. [FinStat Firmy s najväčšími tržbami](https://finstat.sk/databaza-financnych-udajov?sort=sales-desc&years=2020)
5. [FinStat Najziskovejšie Firmy](https://finstat.sk/databaza-financnych-udajov?sort=profit-desc&years=2020)
6. [FinStat Najväčší zamestnávatelia](https://finstat.sk/databaza-firiem-organizacii?sort=empl-desc)
7. [FinStat Najväčšie univerzity](https://finstat.sk/databaza-neziskoviek?sort=revenue-desc&tab=revenue&legalform=382)

## 📔 Dokumentácia Microsoft Excel a Guidelines
1. [Microsoft Excel help & learning](https://support.microsoft.com/en-us/excel)
2. [Premium templates](https://templates.office.com/en-us/premium-templates)
3. [Analyze Data in Microsoft Excel](https://support.microsoft.com/en-us/office/analyze-data-in-excel-3223aab8-f543-4fda-85ed-76bb0295ffc4)
4. [Microsoft Excel functions (alphabetical)](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188)
5. [The Ultimate Guide to Using Microsoft Excel](https://blog.hubspot.com/marketing/microsoft-excel)
6. [Microsoft Excel Statistics](https://support.microsoft.com/en-us/office/check-workbook-statistics-afa12d4b-9584-4826-99a8-33228467e006)

## 📈 YouTube video záznamy z kurzy a prednášok Playlist (Kancelárske Balíky)
[YouTube kanál IT Academy](https://www.youtube.com/watch?v=6nbo18YVf5g&list=PLIu_ZdHo7Pk-rY_6wVj108Dmff67eQWRG)

## 📎Obsah Kurzu
### 📑 Microsoft Excel má 3 typy Hárkov (Sheets)
1. **Pracovný Hárok (Worksheet) (Shift + F11)**
2. Makro Hárok (Macro Sheet) (Ctrl + F11)
3. Grafový Hárok (Graph Sheet) (F11)

![harky](https://user-images.githubusercontent.com/24510943/212564384-aa4f4b9a-1b41-419b-b67a-6b5dfa0053cc.png)

### 🔥 Duplikácia a Kopírovanie Formátu  
* Hromadné Vkladanie, Generovanie Hodnôt (Ctrl + Enter)  
* Kopírovanie Formátu Metlička (2-klik na metlu)  

### 📋 Tabuľky a Rýchla Analýza Dát
* Vytvorenie Tabuľky (Ctrl + T, Ctrl + Shift + L)

* Rýchla Analýza Dát/Quick Analysis (Ctrl + Q)  

**Ako nepomenovávať:**
1. Žiadne neviditeľné symboly t.j. bez medzier/tabov
2. Nezačínaš číslom
3. Neštandardné znaky € / * @ $ ^ & # + - 
4. Bez diakritiky
5. Nie generické názvy tabulka1

**Ako pomenovať:**
1. **Maďarská notácia/zápis**
> tab
> t
> dim
> d
> tMzdyZamestnanciZima2023

2. **Ťavia notácia/zápis**
> klientiLeto2023

3. **Podčiarkovniková notácia/zápis** 
> klienti_leto_2023

**TIP**: KROLA

## 💡 Snippety
### 🗔 Zobraz Prehľadové Okno s Hárkami 
```vb
Sub WbTab()
' Zobraz Prehľadové Okno s Hárkami (Taby)
    Application.CommandBars("Workbook tabs").ShowPopup
End Sub
```

### Funkcie a Vzorce (Formulas)
#### MEDIAN - Štatistická Funkcia - Stredná hodnota  
Medián čísel v rozsahu buniek. Medián je stredná hodnota zoradeného rozsahu čísel
```
=MEDIAN(A2:A7)	
```

####  POWER - Matematická Funkcia - Umocnenie čísla 
```
=POWER(5,2)	Vypočíta druhú mocninu čísla 5 (25)
```
```
=5^3	Vypočíta tretiu mocninu čísla 5 (125
```

#### REPT - Matematická Funkcia - Opakovanie znakov v bunke
```
=REPT(".";6)	Opakovanie obdobia (.) 6-krát (......)  
```
```
=REPT("-";4)	Opakovanie pomlčky (-) 4-krát (----)    
```

#### AND, OR, NOT, IF - Logické Funkcie - Spájanie Funkcií
```
=AND(A2>A3; A2<A4)	Je číslo 15 väčšie ako 9 a menšie ako 8? (FALSE)  
```
```
=OR(A2>A3; A2<A4)	Je číslo 15 väčšie ako 9 alebo menšie ako 8? (TRUE)  
```
```
=NOT(A2+A3=24)	Nie je súčet 15 plus 9 rovný 24? (FALSE)  
```

```
=IF(A2=15; "OK"; "Nie OK")	Ak sa hodnota v bunke A2 rovná 15, vráť hodnotu "OK". (OK)  
```
```
=IF(AND(A2>A3; A2<A4); "OK"; "Nie OK")	Ak je číslo 15 väčšie ako 9 a menšie ako 8, vráť hodnotu "OK". (Nie OK)  
```
```
=IF(OR(A2>A3; A2<A4); "OK"; "Nie OK")	Ak je číslo 15 väčšie ako 9 alebo menšie ako 8, vráť hodnotu "OK". (OK)  
```
```
=IF(A3>89;"A";IF(A3>79;"B";IF(A3>69;"C";IF(A3>59;"D";"F"))))  
```
#### Kombinácia Funkcií na Podmienené Formátovanie podľa výsledkov podielu/percent z obratu/predaja
```
=IF(F9>0;REPT("●";INT(F9*100));"")
```
```
=IF(F10<0;REPT("●";INT(-F10*100));"
```

#### Vyhľadávacie Funkcie  
Typ zhody  
A. Presne (exact match): 0, False   
B. Približne: 1, True, Nič 

0 nie je nič Null   

```
=IFNA(VLOOKUP(TRIM(C15);B7:C13;2;0); "Nepracuje u nás")  
```
```
="Q"&VLOOKUP(B25;$E$25:$G$28;3;1)  
```

#### Čistenie Dát
TRIM - Odstráňovanie medzier/Neviditeľné symboly  
CLEAN - Odstráňovanie netlačiteľných symboly  
VALUE - Konverzia Textu na Číselnú Hodnotu  
```
ABC(VALUE(CLEAN(TRIM(F15))))  
```

## Typy Súborov/Rozšírení Microsoft Excel
1. **XLSX (Textové)**  
2. XLSM (Textové)  
3. XLS (Binárne)  
4. **XLSB (Binárne)**  


## Funkcia pre sum na základe farby SumFarba (Podmienené Formátovanie cez VBA)
```vb
' Farba Vyplne/Bunky
' Rozsah Stlpec, Oblast, Tabulka
Function sum_farba(Farba As Range, Rozsah As Range)
    Dim X As Double
    Dim Y As Double
    Dim i As Variant 'Object
    
    Y = Farba.Interior.ColorIndex
    
    For Each i In Rozsah
        If i.Interior.ColorIndex = Y Then
            X = WorksheetFunction.Sum(i, X)
        End If
    Next i
    
    sum_farba = X
End Function
```
## Funkcia pre min na základe farby MinFarba (Podmienené Formátovanie cez VBA)
```vb
' Farba Vyplne/Bunky
' Rozsah Stlpec, Oblast, Tabulka
Function min_farba(Farba As Range, Rozsah As Range)
    Dim X As Double
    Dim Y As Double
    Dim i As Variant 'Object
    
    Y = Farba.Interior.ColorIndex
    
    For Each i In Rozsah
        If i.Interior.ColorIndex = Y Then
            X = WorksheetFunction.Min(i, X)
        End If
    Next i
    
    min_farba = X
End Function
```
## Subrutina s udalosťou typu zmena výberu na hárku pre zvýraznenie celého stĺpca a riadku podľa vybranej bunky
```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' ColorIndex property (Excel Graph)
    ' https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex
    ' 1 - cierna, 2 - biela, 3 - cervena, 4 - Zelena,
    ' 5 - Modra, 6 - zlta, 7 - magenta, 8 - cyan, 9 - bordova
    Cells.Interior.ColorIndex = xlColorIndexNone
    Target.EntireColumn.Interior.ColorIndex = 6
    Target.EntireRow.Interior.ColorIndex = 6
    Target.Interior.ColorIndex = xlColorIndexNone
End Sub
```

## ⬇️ Subrutina na vzostupne zoradenie hárkov (0-9A-Z)
```vb
' Subrutina na vzostupne zoradenie hárkov (0-9A-Z)
Sub zorad_harky_vzostupne()
	' Vypnutie aktualizacie udajov na obrazovke, vyrazne zrýchluje spracovanie makier/subrutín
	Application.ScreenUpdating = False
	' Premenná, kde ulozime celkovy pocet vsetkych harkov (pracovnych, grafovych aj makro harkov)
	Dim ShCount As Integer: ShCount = Sheets.Count  
	' Pomocné premenné i a j pre cyklus, pre preiterovanie celého Excel súboru/zošita (Workbook)
	Dim i As Integer 
	Dim j As Integer

	For i = 1 To ShCount - 1
		For j = i + 1 To ShCount
			If UCase(Sheets(j).Name) < UCase(Sheets(i).Name) Then
				Sheets(j).Move before:=Sheets(i)
			End If
		Next j
	Next i

	' Zapnutie aktualizacie udajov na obrazovke
	Application.ScreenUpdating = True
End Sub
```

## ⬆️ Subrutina na vzostupne zoradenie hárkov (0-9A-Z)
```vb
' Subrutina na zostupne zoradenie hárkov (Z-A9-0)
Sub zorad_harky_zostupne()
	Application.ScreenUpdating = False
	' Premenná, kde ulozime celkovy pocet vsetkych harkov (pracovnych, grafovych aj makro harkov)
	Dim ShCount As Integer: ShCount = Sheets.Count  
	' Pomocné premenné i a j pre cyklus, pre preiterovanie celého Excel súboru/zošita (Workbook)
	Dim i As Integer 
	Dim j As Integer
	
	For i = 1 To ShCount - 1
		For j = i + 1 To ShCount
			If UCase(Sheets(j).Name) > UCase(Sheets(i).Name) Then
				Sheets(j).Move before:=Sheets(i)
			End If
		Next j
	Next i
	
	Application.ScreenUpdating = True
End Sub
```

## ↕️ Subrutina na zoradenie hárkov podľa vstupu používateľa
```vb
' Subrutina na zoradenie hárkov podľa vstupu používateľa, 
' Máš na výber vzostupne - Yes, zostupne - No
Sub zorad_harky_podla_vstupu()
	Application.ScreenUpdating = False
	Dim ShCount As Integer, i As Integer, j As Integer
	Dim SortOrder As VbMsgBoxResult
	SortOrder = MsgBox("Vyberte Yes pre vzostupné poradie a No pre zostupné poradie", vbYesNoCancel)
	ShCount = Sheets.Count
	
	For i = 1 To ShCount - 1
		For j = i + 1 To ShCount
			If SortOrder = vbYes Then
				If UCase(Sheets(j).Name) < UCase(Sheets(i).Name) Then
					Sheets(j).Move before:=Sheets(i)
				End If
			ElseIf SortOrder = vbNo Then
				If UCase(Sheets(j).Name) > UCase(Sheets(i).Name) Then
				Sheets(j).Move before:=Sheets(i)
				End If
			End If
		Next j
	Next i
	
	Application.ScreenUpdating = True
End Sub
```
