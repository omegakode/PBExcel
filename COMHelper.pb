;COMHelper.pb

EnableExplicit

XIncludeFile "COMHelper.pbi"

DataSection
	IID_NULL:
	Data.l $00000000
	Data.w $0000, $0000
	Data.b $00, $00, $00, $00, $00, $00, $00, $00
EndDataSection

;- enum DISPID
#DISPID_UNKNOWN	= -1
#DISPID_VALUE	= 0
#DISPID_PROPERTYPUT	= -3
#DISPID_NEWENUM	= -4
#DISPID_EVALUATE = -5
#DISPID_CONSTRUCTOR	= -6
#DISPID_DESTRUCTOR = -7
#DISPID_COLLECT = -8

#VARIANT_ALPHABOOL = $0002

;- COM_NAMES1
Structure COM_NAMES1
	name.i
EndStructure

;- COM_VARIANT_ARRAY
Structure COM_VARIANT_ARRAY
	var.VARIANT[0]
EndStructure

;- DECLARES
DeclareDLL.s COM_PeekBstr(bstr.i, free.b)
DeclareDLL.s COM_PeekVarString(*v.VARIANT, free.b)

;-
ProcedureDLL.l COM_Init()
	ProcedureReturn CoInitialize_(0)
EndProcedure

;-
;- ERROR
Procedure.l COM_ErrClear(*err.COM_INVOKE_ERROR)
	*err\scode = #S_OK
	*err\argError = 0
	*err\description = ""
EndProcedure

;-
;- DISPATCH
ProcedureDLL.l COM_GetDispID(disp.IDispatch, name.i)	
	Protected.COM_NAMES1 names
	Protected.l dispids
	
	If disp = 0 : ProcedureReturn #DISPID_UNKNOWN : EndIf
	
	names\name = name
	disp\GetIDsOfNames(?IID_NULL, @names, 1, #LOCALE_USER_DEFAULT, @dispids)
	
	ProcedureReturn dispids
EndProcedure

;-
ProcedureDLL.l COM_CallFunction(disp.IDispatch, func.i, *args.COM_VARIANT_ARRAY, argsCount.l, *result.VARIANT, *err.COM_INVOKE_ERROR)
	Protected.DISPPARAMS dp
	Protected.l iArg
	
	COM_ErrClear(*err)
	
	If *result : VariantInit_(*result) : EndIf
	
	If disp = 0
		*err\scode = #E_POINTER
		ProcedureReturn *err\scode
	EndIf

	dp\cArgs = argsCount
	dp\rgvarg = *args
	
	*err\scode = disp\Invoke(COM_GetDispID(disp, func), ?IID_NULL, #LOCALE_USER_DEFAULT, #DISPATCH_METHOD, 
		@dp, *result, @*err\exc, @*err\argError)
	If *err\scode <> #S_OK
		*err\description = COM_PeekBstr(*err\exc\bstrDescription, #True)
		*err\exc\bstrDescription = 0
	EndIf
			
	ProcedureReturn *err\scode
EndProcedure

ProcedureDLL.l COM_CallFunctionLong(disp.IDispatch, func.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_CallFunction(disp, func, *args, argsCount, @result, *err) = #S_OK
		ProcedureReturn result\lVal
	EndIf
EndProcedure

ProcedureDLL.w COM_CallFunctionWord(disp.IDispatch, func.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_CallFunction(disp, func, *args, argsCount, @result, *err) = #S_OK
		ProcedureReturn result\iVal
	EndIf
EndProcedure

ProcedureDLL.b COM_CallFunctionByte(disp.IDispatch, func.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_CallFunction(disp, func, *args, argsCount, @result, *err) = #S_OK
		ProcedureReturn result\bVal
	EndIf
EndProcedure

ProcedureDLL.f COM_CallFunctionFloat(disp.IDispatch, func.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_CallFunction(disp, func, *args, argsCount, @result, *err) = #S_OK
		ProcedureReturn result\fltVal
	EndIf
EndProcedure

ProcedureDLL.d COM_CallFunctionDouble(disp.IDispatch, func.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_CallFunction(disp, func, *args, argsCount, @result, *err) = #S_OK
		ProcedureReturn result\dblVal
	EndIf
EndProcedure

ProcedureDLL.w COM_CallFunctionBool(disp.IDispatch, func.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_CallFunction(disp, func, *args, argsCount, @result, *err) = #S_OK
		ProcedureReturn result\boolVal
	EndIf
EndProcedure

ProcedureDLL.s COM_CallFunctionString(disp.IDispatch, func.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	Protected.s ret
	
	If COM_CallFunction(disp, func, *args, argsCount, @result, *err) = #S_OK
		ProcedureReturn COM_PeekVarString(@result, #True)
	EndIf
EndProcedure

ProcedureDLL.i COM_CallFunctionDispatch(disp.IDispatch, func.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_CallFunction(disp, func, *args, argsCount, @result, *err) = #S_OK
		ProcedureReturn result\pdispVal
	EndIf
EndProcedure

ProcedureDLL.i COM_CallFunctionUnknown(disp.IDispatch, func.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_CallFunction(disp, func, *args, argsCount, @result, *err) = #S_OK
		ProcedureReturn result\punkVal
	EndIf
EndProcedure

ProcedureDLL.i COM_CallFunctionSafeArray(disp.IDispatch, func.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_CallFunction(disp, func, *args, argsCount, @result, *err) = #S_OK
		ProcedureReturn result\parray
	EndIf
EndProcedure

;-
ProcedureDLL.l COM_PutProperty(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.DISPPARAMS dp
	Protected.l dispid

	COM_ErrClear(*err)
	
	If disp = 0
		*err\scode = #E_POINTER
		ProcedureReturn *err\scode
	EndIf
	
	dispid = #DISPID_PROPERTYPUT
	dp\rgvarg = *args
	dp\cArgs = argsCount
	dp\cNamedArgs = 1
	dp\rgdispidNamedArgs = @dispid
	
	*err\scode = disp\Invoke(COM_GetDispID(disp, prop), ?IID_NULL, #LOCALE_USER_DEFAULT, #DISPATCH_PROPERTYPUT, @dp, #Null, @*err\exc, @*err\argError)

	ProcedureReturn *err\scode
EndProcedure

ProcedureDLL.l COM_PutPropertyLong(disp.IDispatch, prop.i, lVal.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT v
	
	v\vt = #VT_I4
	v\lVal = lVal
	
	ProcedureReturn COM_PutProperty(disp, prop, @v, 1, *err)
EndProcedure

ProcedureDLL.l COM_PutPropertyWord(disp.IDispatch, prop.i, wVal.w, *err.COM_INVOKE_ERROR)
	Protected.VARIANT v
	
	v\vt = #VT_I2
	v\iVal = wVal
	
	ProcedureReturn COM_PutProperty(disp, prop, @v, 1, *err)
EndProcedure

ProcedureDLL.l COM_PutPropertyByte(disp.IDispatch, prop.i, bVal.b, *err.COM_INVOKE_ERROR)
	Protected.VARIANT v
	
	v\vt = #VT_I1
	v\bVal = bVal
	
	ProcedureReturn COM_PutProperty(disp, prop, @v, 1, *err)
EndProcedure

ProcedureDLL.l COM_PutPropertyFloat(disp.IDispatch, prop.i, fVal.f, *err.COM_INVOKE_ERROR)
	Protected.VARIANT v
	
	v\vt = #VT_R4
	v\fltVal = fVal
	
	ProcedureReturn COM_PutProperty(disp, prop, @v, 1, *err)
EndProcedure

ProcedureDLL.l COM_PutPropertyDouble(disp.IDispatch, prop.i, dVal.d, *err.COM_INVOKE_ERROR)
	Protected.VARIANT v
	
	v\vt = #VT_R8
	v\dblVal = dVal
	
	ProcedureReturn COM_PutProperty(disp, prop, @v, 1, *err)
EndProcedure

ProcedureDLL.l COM_PutPropertyBool(disp.IDispatch, prop.i, boolVal.w, *err.COM_INVOKE_ERROR)
	Protected.VARIANT v
	
	v\vt = #VT_BOOL
	v\boolVal = boolVal
	
	ProcedureReturn COM_PutProperty(disp, prop, @v, 1, *err)
EndProcedure

ProcedureDLL.l COM_PutPropertyString(disp.IDispatch, prop.i, sVal.s, *err.COM_INVOKE_ERROR)
	Protected.VARIANT v
	Protected.l hr
	
	v\vt = #VT_BSTR
	v\bstrVal = SysAllocString_(sVal)
	hr = COM_PutProperty(disp, prop, @v, 1, *err)
	SysFreeString_(v\pbstrVal)

	ProcedureReturn hr
EndProcedure

ProcedureDLL.l COM_PutPropertyDispatch(disp.IDispatch, prop.i, dispVal.IDispatch, *err.COM_INVOKE_ERROR)
	Protected.VARIANT v
	
	v\vt = #VT_DISPATCH
	v\pdispVal = dispVal
	
	ProcedureReturn COM_PutProperty(disp, prop, @v, 1, *err)
EndProcedure

ProcedureDLL.l COM_PutPropertyUnknown(disp.IDispatch, prop.i, unk.IUnknown, *err.COM_INVOKE_ERROR)
	Protected.VARIANT v
	
	v\vt = #VT_UNKNOWN
	v\punkVal = unk
	
	ProcedureReturn COM_PutProperty(disp, prop, @v, 1, *err)
EndProcedure

ProcedureDLL.l COM_PutPropertySafeArray(disp.IDispatch, prop.i, *sa.SAFEARRAY, *err.COM_INVOKE_ERROR)
	Protected.VARIANT v
	
	v\vt = #VT_SAFEARRAY
	v\pparray = *sa
	
	ProcedureReturn COM_PutProperty(disp, prop, @v, 1, *err)
EndProcedure

;-
ProcedureDLL.l COM_GetProperty(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *result.VARIANT, *err.COM_INVOKE_ERROR)	
	Protected.DISPPARAMS dp
	
	COM_ErrClear(*err)
	
	If *result : VariantInit_(*result) : EndIf

	If disp = 0
		*err\scode = #E_POINTER
		ProcedureReturn *err\scode
	EndIf
		
	If argsCount > 0
		dp\rgvarg = *args
		dp\cArgs = argsCount
	EndIf 

	*err\scode = disp\Invoke(COM_GetDispID(disp, prop), ?IID_NULL, #LOCALE_USER_DEFAULT, #DISPATCH_PROPERTYGET, @dp, *result, @*err\exc, @*err\argError)
	If *err\scode <> #S_OK
		*err\description = COM_PeekBstr(*err\exc\bstrDescription, #True)
		*err\exc\bstrDescription = 0
	EndIf
	
	ProcedureReturn *err\scode
EndProcedure

ProcedureDLL.l COM_GetPropertyLong(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_GetProperty(disp.IDispatch, prop, *args, argsCount, @result, *err)	= #S_OK
		ProcedureReturn result\lVal
	EndIf 
EndProcedure

ProcedureDLL.w COM_GetPropertyWord(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_GetProperty(disp.IDispatch, prop, *args, argsCount, @result, *err)	= #S_OK
		ProcedureReturn result\iVal
	EndIf 
EndProcedure

ProcedureDLL.b COM_GetPropertyByte(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_GetProperty(disp.IDispatch, prop, *args, argsCount, @result, *err)	= #S_OK
		ProcedureReturn result\bVal
	EndIf 
EndProcedure

ProcedureDLL.f COM_GetPropertyFloat(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_GetProperty(disp.IDispatch, prop, *args, argsCount, @result, *err)	= #S_OK
		ProcedureReturn result\fltVal
	EndIf 
EndProcedure

ProcedureDLL.d COM_GetPropertyDouble(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_GetProperty(disp.IDispatch, prop, *args, argsCount, @result, *err)	= #S_OK
		ProcedureReturn result\dblVal
	EndIf 
EndProcedure

ProcedureDLL.w COM_GetPropertyBool(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_GetProperty(disp.IDispatch, prop, *args, argsCount, @result, *err)	= #S_OK
		ProcedureReturn result\boolVal
	EndIf 
EndProcedure

ProcedureDLL.s COM_GetPropertyString(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	Protected.s ret
	
	If COM_GetProperty(disp.IDispatch, prop, *args, argsCount, @result, *err)	= #S_OK
		If result\bstrVal
			ret = PeekS(result\bstrVal)
			SysFreeString_(result\bstrVal)
			
			ProcedureReturn ret
		EndIf
	EndIf 
EndProcedure

ProcedureDLL.i COM_GetPropertyDispatch(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err)
	Protected.VARIANT result
	
	If COM_GetProperty(disp.IDispatch, prop, *args, argsCount, @result, *err)	= #S_OK
		ProcedureReturn result\pdispVal
	EndIf 
EndProcedure

ProcedureDLL.i COM_GetPropertyUnknown(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_GetProperty(disp.IDispatch, prop, @result, *args, argsCount, *err)	= #S_OK
		ProcedureReturn result\punkVal
	EndIf 
EndProcedure

ProcedureDLL.i COM_GetPropertySafeArray(disp.IDispatch, prop.i, *args.COM_VARIANT_ARRAY, argsCount.l, *err.COM_INVOKE_ERROR)
	Protected.VARIANT result
	
	If COM_GetProperty(disp.IDispatch, prop, *args, argsCount, @result, *err)	= #S_OK
		ProcedureReturn result\parray
	EndIf 
EndProcedure

;-
;- VARIANT
ProcedureDLL.l COM_VarClone(*vSrc.VARIANT, *vDest.VARIANT)
	*vDest\vt = *vSrc\vt
	Select *vSrc\vt
		Case #VT_I4 : *vDest\lVal = *vSrc\lVal
		Case #VT_R4 : *vDest\fltVal = *vSrc\fltVal
		Case #VT_R8 : *vDest\dblVal = *vSrc\dblVal
		Case #VT_BOOL : *vDest\boolVal = *vSrc\boolVal
		Case #VT_BSTR : *vDest\bstrVal = *vSrc\bstrVal
		Case #VT_SAFEARRAY : *vDest\parray = *vSrc\parray
		Case #VT_DISPATCH : *vDest\pdispVal = *vSrc\pdispVal
		Case #VT_UNKNOWN : *vDest\punkVal = *vSrc\punkVal
		Case #VT_I2 : *vDest\iVal = *vSrc\iVal
		Case #VT_I1 : *vDest\bVal = *vSrc\bVal
		
		Default : CopyMemory(*vSrc, *vDest, SizeOf(VARIANT))
	EndSelect
EndProcedure

ProcedureDLL.i COM_VarClear(*v)
	VariantClear_(*v)
	
	ProcedureReturn *V
EndProcedure

ProcedureDLL.i COM_VarInit(*v)
	VariantInit_(*v)
	
	ProcedureReturn *V
EndProcedure

ProcedureDLL.i COM_VarNone(*v.VARIANT)
	*v\vt = #VT_ERROR
	*v\scode = #DISP_E_PARAMNOTFOUND
	
	ProcedureReturn *v
EndProcedure

ProcedureDLL.s COM_VarToString(*v.VARIANT, free.b)
	Protected.VARIANT _v
	
	If *v\vt = #VT_BSTR
		ProcedureReturn COM_PeekBstr(*v\bstrVal, free)
	EndIf 
	
	If VariantChangeType_(@_v, *v, #VARIANT_ALPHABOOL, #VT_BSTR) = #S_OK
		ProcedureReturn COM_PeekBstr(_v\bstrVal, #True)
	EndIf 
EndProcedure

ProcedureDLL.i COM_VarLong(*v.VARIANT, lVal.l)
	*v\vt = #VT_I4
	*v\lVal = lVal
	
	ProcedureReturn *v
EndProcedure

ProcedureDLL.i COM_VarWord(*v.VARIANT, wVal.w)
	*v\vt = #VT_I2
	*v\iVal = wVal
	
	ProcedureReturn *v
EndProcedure

ProcedureDLL.i COM_VarByte(*v.VARIANT, bVal.b)
	*v\vt = #VT_I1
	*v\bVal = bVal
	
	ProcedureReturn *v
EndProcedure

ProcedureDLL.i COM_VarFloat(*v.VARIANT, fltVal.f)
	*v\vt = #VT_R4
	*v\fltVal = fltVal
	
	ProcedureReturn *v
EndProcedure

ProcedureDLL.i COM_VarDouble(*v.VARIANT, dblVal.d)
	*v\vt = #VT_R8
	*v\dblVal = dblVal
	
	ProcedureReturn *v
EndProcedure

ProcedureDLL.i COM_VarBool(*v.VARIANT, boolVal.w)
	*v\vt = #VT_BOOL
	*v\boolVal = boolVal
	
	ProcedureReturn *v
EndProcedure

ProcedureDLL.i COM_VarString(*v.VARIANT, sVal.s)
	*v\vt = #VT_BSTR
	*v\bstrVal = SysAllocString_(sVal)
	
	ProcedureReturn *v
EndProcedure

ProcedureDLL.i COM_VarDispatch(*v.VARIANT, dispVal.IDispatch)
	*v\vt = #VT_DISPATCH
	*v\pdispVal = dispVal
	
	ProcedureReturn *v
EndProcedure

ProcedureDLL.i COM_VarUnknown(*v.VARIANT, unkVal.IUnknown)
	*v\vt = #VT_UNKNOWN
	*v\punkVal = unkVal
	
	ProcedureReturn *v
EndProcedure

ProcedureDLL.i COM_VarSafeArray(*v.VARIANT, arr.i)
	*v\vt = #VT_SAFEARRAY
	*v\parray = arr
	
	ProcedureReturn *v
EndProcedure

;-
;- STRING
ProcedureDLL.s COM_PeekBstr(bstr.i, free.b)
	Protected.s ret
	
	If bstr
		ret = PeekS(bstr)
		If free : SysFreeString_(bstr) : EndIf
	EndIf
	
	ProcedureReturn ret
EndProcedure

ProcedureDLL.s COM_PeekVarString(*v.VARIANT, free.b)
	If *v And *v\vt = #VT_BSTR
		ProcedureReturn COM_PeekBstr(*v\bstrVal, free)
	EndIf
EndProcedure