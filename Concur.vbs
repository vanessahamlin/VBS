'Open File
'Read Cell with BU
'if statement
	'if it's BU matches then put on spreadsheet and save it.
	'Business Unit' ' HCC - GLASSEAL' NCC REYNOSA' AEROSPACE REYNOSA' ASM (SPECIALTY MOTORS) REYNOSA' MICROPOISE' HUGHES TREITLER REYNOSA' POWER INSTRUMENTS REYNOSA' TIP' Solid State Controls' POWER INSTRUMENTS' SEACON PHOENIX' EDAX' ZYGO' AMERON' CPD - NESQ' PRECITECH' VISION RESEARCH' HUGHES TREITLER' TCI' AERO - PDS' SOUTHERN AEROPARTS' CHANDLER' AMT' COINING' CAMECA INSTRUMENTS' AEROSPACE' APT' CORPORATE' MICROPOISE OH' O'BRIEN - STL' REICHERT' CRYSTAL ENGINEERING' PAI - PROCESS' TMC' PROGRAMMABLE POWER' DRAKE AIR' ATLAS' HCC - HERMETICS' PI' HCC' HAYDON KERK - NH' US GAUGE' READING ALLOYS' PI Support Services' MIL AERO' CSI' O'BRIEN - CAR' VIS - REY' SOLID STATE CANADA' AVICENNA' LAMB' HAYDON KERK - CT' VIS - GRAND JUNCTION' PRESTOLITE POWER' B & S AIRCRAFT' TSE' PAI - SUPPORT SERVICES' TIP - PITTMAN' DunkerMotoren' ADVANCED INDUSTRIES' HIGH STANDARD AVIATION' HAMILTON PRECISION METALS' NEWAGE TESTING' PAI CANADA' TIP - ROCK CREEK' SMP - 84' DUNKERMOTEREN' ATLAS WEATHERING' MICROPOISE MI' SPECTRO' HCC - SEALTRON' TIP - SAUGERTIES' LAMB HQ' PowerVar Mexico' CHEMICAL PRODUCTS' TESEQ' POWERVAR' LAND' SMP - CT' IFI' PAI' SOLID STATE HDR' ACI' AMPTEK' PETROLAB' Solidstate Controls Mexico' HCC INTERCONNECTS MEXICO' CORPORATE EO' LAMB COMMERCIAL' AVTECH' TIP Rock Creek' EDAX - DRAPER' VIS - REYNOSA' 
	
	

Dim objRootDSE, adoConnection, adoCommand, strQuery, strText
Dim adoRecordset, strDNSDomain, objShell, lngBiasKey
Dim lngBias, k, strDN, dtmDate, objDate
Dim strBase, strFilter, strAttributes, lngHigh, lngLow
Dim unusedRow

Const xlCellTypeLastCell = 11
Const xlAscending = 1
Const xlYes = 1
Const ForAppending = 8

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Concur.xlsx")
Set objWorksheet = objWorkbook.Worksheets("page")
objWorksheet.Activate

intRow = 5


Do until objExcel.Cells(intRow, 13).Value = ""
	strStartDate = objExcel.Cells(intRow, 1).Value 
	strTravelRuleClass = objExcel.Cells(intRow, 2).Value 
	strEmplID = objExcel.Cells(intRow, 3).Value 
	strEmployeeName = objExcel.Cells(intRow, 4).Value 
	strLogonID = objExcel.Cells(intRow, 5).Value 
	strEmailAddress = objExcel.Cells(intRow, 6).Value 
	strLocale = objExcel.Cells(intRow, 7).Value 
	strCountry = objExcel.Cells(intRow, 8).Value 
	strStateProvinceRegion = objExcel.Cells(intRow, 9).Value 
	strLedgerCode = objExcel.Cells(intRow, 10).Value 
	strCurr = objExcel.Cells(intRow, 11).Value 
	strBU = objExcel.Cells(intRow, 12).Value 
	strBusinessUnit = objExcel.Cells(intRow, 13).Value 
	strBusinessUnit = UCase(strBusinessUnit)
	strPL = objExcel.Cells(intRow, 14).Value 
	strPLName = objExcel.Cells(intRow, 15).Value 
	strPaygroup = objExcel.Cells(intRow, 16).Value 
	strPaygroupName = objExcel.Cells(intRow, 17).Value 
	strSite = objExcel.Cells(intRow, 18).Value 
	strLocation = objExcel.Cells(intRow, 19).Value 
	strFinanceMngr = objExcel.Cells(intRow, 20).Value 
	strManager = objExcel.Cells(intRow, 21).Value 
	strHierarchy = objExcel.Cells(intRow, 22).Value 
	strMngrID = objExcel.Cells(intRow, 23).Value 
	strIsTestUser = objExcel.Cells(intRow, 24).Value 
	strTravFlag = objExcel.Cells(intRow, 25).Value 
	strExpFlag = objExcel.Cells(intRow, 26).Value 
	strApprover = objExcel.Cells(intRow, 27).Value 
	strLastName = objExcel.Cells(intRow, 28).Value 
	strFirstName = objExcel.Cells(intRow, 29).Value 
	strFieldMngr = objExcel.Cells(intRow, 30).Value 
	strCountEmployeeName = objExcel.Cells(intRow, 31).Value 
	

	
	
	'wscript.echo strBusinessUnit
		For Each objWorksheet in objWorkbook.Worksheets
			If objWorksheet.Name = strBusinessUnit Then
				x = 1
				Exit For
			End If
		Next

		If x = 1 Then
			'Wscript.Echo "The specified worksheet was found."
			objExcel.Sheets(strBusinessUnit).Activate
			Set objRange = objWorksheet.UsedRange
			objRange.SpecialCells(xlCellTypeLastCell).Activate

			intNewRow = objExcel.ActiveCell.Row + 1
			objExcel.Cells(intNewRow, 1).Value = strStartDate
			objExcel.Cells(intNewRow, 2).Value = strTravelRuleClass
			objExcel.Cells(intNewRow, 3).Value = strEmplID
			objExcel.Cells(intNewRow, 4).Value = strEmployeeName
			objExcel.Cells(intNewRow, 5).Value = strLogonID
			objExcel.Cells(intNewRow, 6).Value = strEmailAddress
			objExcel.Cells(intNewRow, 7).Value = strLocale
			objExcel.Cells(intNewRow, 8).Value = strCountry
			objExcel.Cells(intNewRow, 9).Value = strStateProvinceRegion
			objExcel.Cells(intNewRow, 10).Value = strLedgerCode
			objExcel.Cells(intNewRow, 11).Value = strCurr
			objExcel.Cells(intNewRow, 12).Value = strBU
			objExcel.Cells(intNewRow, 13).Value = strBusinessUnit
			objExcel.Cells(intNewRow, 14).Value = strPL
			objExcel.Cells(intNewRow, 15).Value = strPLName
			objExcel.Cells(intNewRow, 16).Value = strPaygroup
			objExcel.Cells(intNewRow, 17).Value = strPaygroupName
			objExcel.Cells(intNewRow, 18).Value = strSite
			objExcel.Cells(intNewRow, 19).Value = strLocation
			objExcel.Cells(intNewRow, 20).Value = strFinanceMngr
			objExcel.Cells(intNewRow, 21).Value = strManager
			objExcel.Cells(intNewRow, 22).Value = strHierarchy
			objExcel.Cells(intNewRow, 23).Value = strMngrID
			objExcel.Cells(intNewRow, 24).Value = strIsTestUser
			objExcel.Cells(intNewRow, 25).Value = strTravFlag
			objExcel.Cells(intNewRow, 26).Value = strExpFlag
			objExcel.Cells(intNewRow, 27).Value = strApprover
			objExcel.Cells(intNewRow, 28).Value = strLastName
			objExcel.Cells(intNewRow, 29).Value = strFirstName
			objExcel.Cells(intNewRow, 30).Value = strFieldMngr
			objExcel.Cells(intNewRow, 31).Value = strCountEmployeeName
			objExcel.Sheets("page").Rows(1).Copy objExcel.Sheets(strBusinessUnit).Rows(1)
			
		Else
			'Wscript.Echo "The specified worksheet was not found."
			set objWorksheet = objExcel.Sheets.Add
			objWorksheet.Name = strBusinessUnit
			Set objRange = objWorksheet.UsedRange
			objRange.SpecialCells(xlCellTypeLastCell).Activate

			intNewRow = objExcel.ActiveCell.Row + 1		
			objExcel.Cells(intNewRow, 1).Value = strStartDate
			objExcel.Cells(intNewRow, 2).Value = strTravelRuleClass
			objExcel.Cells(intNewRow, 3).Value = strEmplID
			objExcel.Cells(intNewRow, 4).Value = strEmployeeName
			objExcel.Cells(intNewRow, 5).Value = strLogonID
			objExcel.Cells(intNewRow, 6).Value = strEmailAddress
			objExcel.Cells(intNewRow, 7).Value = strLocale
			objExcel.Cells(intNewRow, 8).Value = strCountry
			objExcel.Cells(intNewRow, 9).Value = strStateProvinceRegion
			objExcel.Cells(intNewRow, 10).Value = strLedgerCode
			objExcel.Cells(intNewRow, 11).Value = strCurr
			objExcel.Cells(intNewRow, 12).Value = strBU
			objExcel.Cells(intNewRow, 13).Value = strBusinessUnit
			objExcel.Cells(intNewRow, 14).Value = strPL
			objExcel.Cells(intNewRow, 15).Value = strPLName
			objExcel.Cells(intNewRow, 16).Value = strPaygroup
			objExcel.Cells(intNewRow, 17).Value = strPaygroupName
			objExcel.Cells(intNewRow, 18).Value = strSite
			objExcel.Cells(intNewRow, 19).Value = strLocation
			objExcel.Cells(intNewRow, 20).Value = strFinanceMngr
			objExcel.Cells(intNewRow, 21).Value = strManager
			objExcel.Cells(intNewRow, 22).Value = strHierarchy
			objExcel.Cells(intNewRow, 23).Value = strMngrID
			objExcel.Cells(intNewRow, 24).Value = strIsTestUser
			objExcel.Cells(intNewRow, 25).Value = strTravFlag
			objExcel.Cells(intNewRow, 26).Value = strExpFlag
			objExcel.Cells(intNewRow, 27).Value = strApprover
			objExcel.Cells(intNewRow, 28).Value = strLastName
			objExcel.Cells(intNewRow, 29).Value = strFirstName
			objExcel.Cells(intNewRow, 30).Value = strFieldMngr
			objExcel.Cells(intNewRow, 31).Value = strCountEmployeeName
			objExcel.Sheets("page").Rows(1).Copy objExcel.Sheets(strBusinessUnit).Rows(1)
			
		End If
		objExcel.Sheets("page").Activate
		intRow = intRow + 1	
		x = 0
	
Loop
For Each objWorksheet in objWorkbook.Worksheets
    'Wscript.Echo objWorksheet.Name
	strWSName = objWorksheet.Name
			
	Set objWorksheet = objWorkbook.Worksheets(strWSName)
	'Set objWorksheet = objWorkbook.Worksheets("Sheet1")
	objWorksheet.Copy
	Set wb = objExcel.ActiveWorkbook 
	wb.SaveAs "C:\Scripts\"&strWSName&".xlsx" 
	wb.Close 

	'objExcel.Quit
Next

wscript.echo "FINISHED"
		