;ExcelWrite.pb

EnableExplicit

XIncludeFile "..\Excel.pb"

Procedure main()
	Protected.IDispatch excel, workBooks, workBook, cellsRange
	Protected.VARIANT vRowIndex, vColIndex, vCellRange, vStr, vBool, vNone, vLong
	Protected.l iCol, iRow
	Protected.s file
	
	COM_Init()
	
	COM_VarNone(@vNone)

	;Create Excel application
	excel = Excel_Application()
	If excel = 0
		Debug "Error failed to create Excel application"
		ProcedureReturn
	EndIf
	Excel_Application_Put_Visible(excel, #VARIANT_TRUE)
	Excel_Application_Put_DisplayAlerts(excel, #VARIANT_FALSE)
	
	;Add a workbook
	workBooks = Excel_Application_Get_Workbooks(excel)
	workBook = Excel_Workbooks_Add(workBooks, @vNone)

	;Fill 2 rows of 10 columns
	cellsRange = Excel_Application_Get_Cells(excel)
	For iRow = 1 To 2
		For iCol = 1 To 10
			Excel_Range_Get_Item(cellsRange, COM_VarLong(@vRowIndex, iRow), COM_VarLong(@vColIndex, iCol), @vCellRange)
			If vCellRange\pdispVal = 0 ;Error
				Continue
			EndIf	
			
			Excel_Range_Put_Value2(vCellRange\pdispVal, COM_VarString(@vStr, "Item" + Str(iRow) + "-" + Str(iCol)))

			COM_VarClear(@vStr)
			COM_VarClear(@vCellRange) ; = vCellRange\pdispVal\Release()
		Next
	Next 
	
	file = SaveFileRequester("Save", "", "", 0)
	If file
		Excel_Workbook_SaveAs(workBook, COM_VarString(@vStr, file), @vNone, @vNone, @vNone, @vNone, @vNone, 
			#xlNoChange, @vNone, @vNone, @vNone, @vNone, @vNone)
		COM_VarClear(@vstr)
		
		If Excel_GetLastError() <> 0
			Debug "Failed to save file:"
			Debug Excel_GetLastErrorDescription()
			
		Else
			Debug "File " + file + " saved succesfully"
		EndIf 
	EndIf 
	
	Excel_Workbook_Close(workBook, COM_VarBool(@vBool, #False), @vNone, @vNone)
	workBook\Release()
	workBooks\Release()
	cellsRange\Release()

	Excel_Application_Quit(excel)
	excel\Release()
EndProcedure

main()