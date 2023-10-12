REM  *****  BASIC  *****

Sub Main

	Dim oSheet As Object
	Dim oCol,row, media, oRow, oRow1, oCell As Long
	Dim oMedidor As Integer
	Dim mssg, mssg1, hoja1, hoja2, hoja3, hoja4, hoja5, hoja6 As String
	Dim oSheet1, oSheet2, oSheet3, oSheet4, oSheet5, oSheet6 As Object
	Dim oString1, oString2, oString3, oString4, oString5, oString6 As Object

	media = 80.0
	hoja1 = "September 2nd -8th"
	hoja2 = "Call Efficiency 2.0"
	hoja3 = "Call Handling 2.0"
	hoja4 = "Call Procedures 2.0"
	hoja5 = "Customer Experience 2.0"
	hoja6 = "Knowledge Score 2.0"
	
	checkFirst(media, hoja1, hoja2, hoja3, hoja4, hoja5, hoja6)
	lastStep ()

End Sub

'sheet(3) Customer exp

Function checkFirst(media, hoja1, hoja2, hoja3, hoja4, hoja5, hoja6)
	'verifica cual medidor es menor the 80 informe a llama a la funcion de cada medidor
		
	oSheet = ThisComponent.Sheets.getByName(hoja1)'sheet1
	
	for oCol = 2 to 6
	oRow = 3
		while oRow < 57
			oCell = oSheet.getCellByPosition(oCol,oRow).value
'			if (oCell <> 0 and oCell < 80 ) then
'				checkSecond(oRow)
'				oRow = oRow + 1
'			else
'				oRow = oRow + 1
'			end if
			Select Case oCol
			   Case 2
				if (oCell <> 0 and oCell < media ) then
					checkSecond(oRow,media, hoja2)
					oRow = oRow + 1
				else
					oRow = oRow + 1
				end if
			  Case 3
				if (oCell <> 0 and oCell < media ) then
					checkThird(oRow,media, hoja3)
					oRow = oRow + 1
				else
					oRow = oRow + 1
				end if
			  Case 4
				if (oCell <> 0 and oCell < media ) then
					checkFourth(oRow,media, hoja4)
					oRow = oRow + 1
				else
					oRow = oRow + 1
				end if
			  Case 5
				if (oCell <> 0 and oCell < media ) then
					checkFifth(oRow,media, hoja5)
					oRow = oRow + 1
				else
					oRow = oRow + 1
				end if
			  Case 6
				if (oCell <> 0 and oCell < media ) then
					checkSixth(oRow,media, hoja6)
					oRow = oRow + 1
				else
					oRow = oRow + 1
				end if
			End Select
		Wend
	Next oCol
End Function

Function checkSecond(row,media, hoja2)
	'verifica cual medidor es menor the 80 en call efeciency y devuelve mssg por cada uno

	oSheet = ThisComponent.Sheets.getByName(hoja2)

	if (oSheet.getCellByPosition(1,row).string <> "#DIV/0!" and oSheet.getCellByPosition(1,row).value < media) then 
		oSheet.getCellByPosition(6,row).string = oSheet.getCellByPosition(6,row).string & "Agent Readiness, "
	end if
	if (oSheet.getCellByPosition(2,row).string <> "#DIV/0!" and oSheet.getCellByPosition(2,row).value < media) then 
		oSheet.getCellByPosition(6,row).string = oSheet.getCellByPosition(6,row).string & "Silence Time, "
	end if
	if (oSheet.getCellByPosition(3,row).string <> "#DIV/0!" and oSheet.getCellByPosition(3,row).value < media) then
		oSheet.getCellByPosition(6,row).string = oSheet.getCellByPosition(6,row).string & " Hold Time, "
	end if
	if (oSheet.getCellByPosition(4,row).string <> "#DIV/0!" and oSheet.getCellByPosition(4,row).value < media) then
		oSheet.getCellByPosition(6,row).string = oSheet.getCellByPosition(6,row).string & " AHT CC, "
	end if
	if (oSheet.getCellByPosition(5,row).string <> "#DIV/0!" and oSheet.getCellByPosition(5,row).value < media) then
		oSheet.getCellByPosition(6,row).string = oSheet.getCellByPosition(6,row).string & " AHT FRC,"
	end if
		
end Function

Function checkThird(row,media, hoja3)
	'verifica cual medidor es menor the 80 en call efeciency y devuelve mssg por cada uno

	oSheet = ThisComponent.Sheets.getByName(hoja3)

	if (oSheet.getCellByPosition(1,row).string <> "#DIV/0!" and oSheet.getCellByPosition(1,row).value < media) then 
		oSheet.getCellByPosition(6,row).string = oSheet.getCellByPosition(6,row).string & "Professinalism, "
	end if
	if (oSheet.getCellByPosition(2,row).string <> "#DIV/0!" and oSheet.getCellByPosition(2,row).value < media) then 
		oSheet.getCellByPosition(6,row).string = oSheet.getCellByPosition(6,row).string & "ownership, "
	end if
	if (oSheet.getCellByPosition(3,row).string <> "#DIV/0!" and oSheet.getCellByPosition(3,row).value < media) then
		oSheet.getCellByPosition(6,row).string = oSheet.getCellByPosition(6,row).string & " Understandability, "
	end if
	if (oSheet.getCellByPosition(4,row).string <> "#DIV/0!" and oSheet.getCellByPosition(4,row).value < media) then
		oSheet.getCellByPosition(6,row).string = oSheet.getCellByPosition(6,row).string & " Listening, "
	end if
	if (oSheet.getCellByPosition(5,row).string <> "#DIV/0!" and oSheet.getCellByPosition(5,row).value < media) then
		oSheet.getCellByPosition(6,row).string = oSheet.getCellByPosition(6,row).string & " Call Wrap Time, "
	end if
		
end Function

Function checkFourth(row,media, hoja4)
	'verifica cual medidor es menor the 80 en call efeciency y devuelve mssg por cada uno

	oSheet = ThisComponent.Sheets.getByName(hoja4)

	if (oSheet.getCellByPosition(1,row).string <> "#DIV/0!" and oSheet.getCellByPosition(1,row).value < media) then 
		oSheet.getCellByPosition(5,row).string = oSheet.getCellByPosition(5,row).string & "Call Introduction,  "
	end if
	if (oSheet.getCellByPosition(2,row).string <> "#DIV/0!" and oSheet.getCellByPosition(2,row).value < media) then
		oSheet.getCellByPosition(5,row).string = oSheet.getCellByPosition(5,row).string & " Verification, "
	end if
	if (oSheet.getCellByPosition(3,row).string <> "#DIV/0!" and oSheet.getCellByPosition(3,row).value < media) then
		oSheet.getCellByPosition(5,row).string = oSheet.getCellByPosition(5,row).string & " Hold Language, "
	end if
	if (oSheet.getCellByPosition(4,row).string <> "#DIV/0!" and oSheet.getCellByPosition(4,row).value < media) then
		oSheet.getCellByPosition(5,row).string = oSheet.getCellByPosition(5,row).string & " Call Closing, "
	end if
		
end Function

Function checkfifth(row,media, hoja5)
	'verifica cual medidor es menor the 80 en call efeciency y devuelve mssg por cada uno

	oSheet = ThisComponent.Sheets.getByName(hoja5)

	if (oSheet.getCellByPosition(1,row).string <> "#DIV/0!" and oSheet.getCellByPosition(1,row).value < media) then 
		oSheet.getCellByPosition(4,row).string = oSheet.getCellByPosition(4,row).string & "Customer Sentiment, "
	end if
	if (oSheet.getCellByPosition(2,row).string <> "#DIV/0!" and oSheet.getCellByPosition(2,row).value < media) then
		oSheet.getCellByPosition(4,row).string = oSheet.getCellByPosition(4,row).string & " Customer Effort, "
	end if
	if (oSheet.getCellByPosition(3,row).string <> "#DIV/0!" and oSheet.getCellByPosition(3,row).value < media) then
		oSheet.getCellByPosition(4,row).string = oSheet.getCellByPosition(4,row).string & " Inquiry Resolution, "
	end if
		
end Function

Function checksixth(row,media, hoja6)
	'verifica cual medidor es menor the 80 en call efeciency y devuelve mssg por cada uno

	oSheet = ThisComponent.Sheets.getByName(hoja6)

	if (oSheet.getCellByPosition(1,row).string <> "#DIV/0!" and oSheet.getCellByPosition(1,row).value < media) then 
		oSheet.getCellByPosition(3,row).string = oSheet.getCellByPosition(3,row).string & "Average, "
	end if
	if (oSheet.getCellByPosition(2,row).string <> "#DIV/0!" and oSheet.getCellByPosition(2,row).value < media) then
		oSheet.getCellByPosition(3,row).string = oSheet.getCellByPosition(3,row).string & " Empathy "
	end if
		
end Function

Function lastStep ()

	
	
	for oRow1 = 3 to 56
		oSheet1 = ThisComponent.Sheets.getByName("September 2nd -8th")
		oSheet2 = ThisComponent.Sheets.getByName("Call Efficiency 2.0")
		oSheet3 = ThisComponent.Sheets.getByName("Call Handling 2.0")
		oSheet4 = ThisComponent.Sheets.getByName("Call Procedures 2.0")
		oSheet5 = ThisComponent.Sheets.getByName("Customer Experience 2.0")
		oSheet6 = ThisComponent.Sheets.getByName("Knowledge Score 2.0")
		oString1 = oSheet1.getCellByPosition(9,oRow1)
		oString2 = oSheet2.getCellByPosition(6,oRow1)
		oString3 = oSheet3.getCellByPosition(6,oRow1)
		oString4 = oSheet4.getCellByPosition(5,oRow1)
		oString5 = oSheet5.getCellByPosition(4,oRow1)
		oString6 = oSheet6.getCellByPosition(3,oRow1)
		
		'oSheet1.getCellByPosition(10,3).string = oSheet2.RangeAddress.Sheet
		oString1.string = oString2.string & oString3.string & oString4.string & oString5.string & oString6.string
	next oRow1

end Function