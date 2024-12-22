;ExcelRead.pb

EnableExplicit

XIncludeFile "..\Excel.pb"

Procedure main()
	Protected.IDispatch excel, workBooks, workBook, workSheet, rowsRange, colsRange, usedRange
	Protected.VARIANT vRowIndex, vColIndex, vNone, vCellValue, vCellRange, vBool
	Protected.l rowsCount, colsCount, iRow, iCol
	Protected.s file
	
	COM_Init()
	
	COM_VarNone(@vNone)
	
	file = OpenFileRequester("Open", "", "", 0)
	If file = "" : ProcedureReturn : EndIf

	;Create Excel application
	excel = Excel_Application()
	If excel = 0
		Debug "Error failed to create Excel application"
		ProcedureReturn
	EndIf
	Excel_Application_Put_Visible(excel, #VARIANT_TRUE)
	Excel_Application_Put_DisplayAlerts(excel, #VARIANT_FALSE)
		
	;Open workbook
	workBooks = Excel_Application_Get_Workbooks(excel)
	workBook = Excel_Workbooks_Open(workBooks, file, @vNone, @vNone, @vNone, @vNone, @vNone, 
		@vNone, @vNone, @vNone, @vNone, @vNone, @vNone, @vNone, @vNone, @vNone)
	If workBook = 0
		Debug "Failed to open workbook"
		Debug Excel_GetLastErrorDescription()
		ProcedureReturn
	EndIf
	
	;Read active sheet
	workSheet = Excel_Application_Get_ActiveSheet(excel)
	If workSheet = 0 : ProcedureReturn : EndIf
	
	usedRange = Excel_Worksheet_Get_UsedRange(workSheet)
	rowsRange = Excel_Range_Get_Rows(usedRange)
	colsRange = Excel_Range_Get_Columns(usedRange)
	
	rowsCount = Excel_Range_Get_Count(rowsRange)
	colsCount = Excel_Range_Get_Count(colsRange)
	rowsRange\Release()
	colsRange\Release()
	
	For iRow = 1 To rowsCount
		For iCol = 1 To colsCount
			Excel_Range_Get_Item(usedRange, COM_VarLong(@vRowIndex, iRow), COM_VarLong(@vColIndex, iCol), @vCellRange)
			If vCellRange\pdispVal
				Excel_Range_Get_Value2(vCellRange\pdispVal, @vCellValue)
				If vCellValue\vt <> #VT_EMPTY
					Debug COM_VarToString(@vCellValue, #True)
				EndIf
				
				vCellRange\pdispVal\Release()
			EndIf
		Next 
	Next 
	
	usedRange\Release()
	
	Excel_Workbook_Close(workBook, COM_VarBool(@vBool, #False), @vNone, @vNone)
	workBook\Release()
	workBooks\Release()

	Excel_Application_Quit(excel)
	excel\Release()
EndProcedure

main()
