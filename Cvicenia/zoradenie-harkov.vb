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
