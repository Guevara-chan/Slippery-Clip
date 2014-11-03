;-TOP 
; Kommentar     : Variant Helper 
; Author        : mk-soft 
; Second Author : ts-soft 
; Third Author  : srod 
; Datei         : VariantHelper_Include.pb 
; Version       : 2.08 
; Erstellt      : 30.04.2007 
; Geändert      : 23.10.2008 
; 
; Compilermode  : 
; 
; *************************************************************************************** 
; 
; Informations: 
; 
; SafesArray functions and macros supported only array with one dims 
; 
; 
; 
; 
; *************************************************************************************** 

Define.l vhLastError, saLastError 


; *************************************************************************************** 

Import "oleaut32.lib" 
  SafeArrayAllocDescriptorEx(a.l,b.l,c.l) As "_SafeArrayAllocDescriptorEx@12" 
  SafeArrayGetVartype(a.l,b.l) As "_SafeArrayGetVartype@8" 
EndImport 

; *************************************************************************************** 

;- Structure SAFEARRAY 
;Structure SAFEARRAYBOUND 
;  cElements.l 
;  lLbound.l 
;EndStructure 

Structure pData 
  StructureUnion 
    bVal.b[0]; AS BYTE            ' VT_UI1 
    iVal.w[0]; AS INTEGER         ' VT_I2 
    lVal.l[0]; AS LONG            ' VT_I4 
    llVal.q[0]; AS QUAD           ' VT_I8 
    fltVal.f[0]; AS SINGLE        ' VT_R4 
    dblVal.d[0]; AS DOUBLE        ' VT_R8 
    boolVal.w[0]; AS INTEGER      ' VT_BOOL 
    scode.l[0]; AS LONG           ' VT_ERROR 
    cyVal.l[0]; AS LONG           ' VT_CY 
    date.d[0]; AS DOUBLE          ' VT_DATE 
    bstrVal.l[0]; AS LONG         ' VT_BSTR 
    punkVal.l[0]; AS DWORD        ' VT_UNKNOWN 
    pdispVal.l[0]; AS DWORD       ' VT_DISPATCH 
    parray.l[0]; AS DWORD         ' VT_ARRAY|* 
    Value.Variant[0]; 
  EndStructureUnion 
EndStructure 
  
;Structure SAFEARRAY 
;  cDims.w 
;  fFeatures.w 
;  cbElements.l 
;  cLocks.l 
;  *pvData.pData 
;  rgsabound.SAFEARRAYBOUND[0] 
;EndStructure 

; *************************************************************************************** 

;- Type Constants helps for Variant and SafeArray 

#TLong = #VT_I4 
#TQuad = #VT_I8 
#TWord = #VT_I2 
#TFloat = #VT_R4 
#TDouble = #VT_R8 
#TString = #VT_BSTR 
#TDate = #VT_DATE 

; *************************************************************************************** 


;- Errorhandling 

; *************************************************************************************** 

Procedure.l vhGetLastError() 

  Shared vhLastError 
  
  ProcedureReturn vhLastError 
  
EndProcedure 

; *************************************************************************************** 

Procedure.s vhGetLastMessage() 

  Shared vhLastError 
  
  Protected *Buffer, len, result.s 
  
  len = FormatMessage_(#FORMAT_MESSAGE_ALLOCATE_BUFFER|#FORMAT_MESSAGE_FROM_SYSTEM,0,vhLastError,0,@*Buffer,0,0) 
  If len 
    result = PeekS(*Buffer) 
    LocalFree_(*Buffer) 
    ProcedureReturn result 
  Else 
    ProcedureReturn "Errorcode: " + Hex(vhLastError) 
  EndIf 
  
EndProcedure 

; *************************************************************************************** 

Procedure.l saGetLastError() 

  Shared saLastError 
  
  ProcedureReturn saLastError 
  
EndProcedure 

; *************************************************************************************** 

Procedure.s saGetLastMessage() 

  Shared saLastError 
  
  Protected *Buffer, len, result.s 
  
  len = FormatMessage_(#FORMAT_MESSAGE_ALLOCATE_BUFFER|#FORMAT_MESSAGE_FROM_SYSTEM,0,saLastError,0,@*Buffer,0,0) 
  If len 
    result = PeekS(*Buffer) 
    LocalFree_(*Buffer) 
    ProcedureReturn result 
  Else 
    ProcedureReturn "Errorcode: " + Hex(saLastError) 
  EndIf 
  
EndProcedure 

; *************************************************************************************** 


;- SafeArray Functions 

; *************************************************************************************** 

Procedure saCreateSafeArray(vartype, Lbound, Elements) 

  Shared saLastError 
  
  Protected rgsabound.SAFEARRAYBOUND, *psa 
  
  rgsabound\lLbound = Lbound 
  rgsabound\cElements = Elements 
  saLastError = 0 
  
  *psa = SafeArrayCreate_(vartype, 1, rgsabound) 
  If *psa 
    ProcedureReturn *psa 
  Else 
    saLastError = #E_OUTOFMEMORY 
    ProcedureReturn 0 
  EndIf 
  
EndProcedure 


; *************************************************************************************** 

Procedure saFreeSafeArray(*psa) 

  Shared saLastError 
  
  Protected hr 
  
  saLastError = 0 
  
  hr = SafeArrayDestroy_(*psa) 
  If hr = #S_OK 
    ProcedureReturn #True 
  Else 
    saLastError = hr 
    ProcedureReturn #False 
  EndIf 
  
EndProcedure 


; *************************************************************************************** 

Procedure saGetVartype(*psa) 

  Shared saLastError 
  
  Protected hr, vartype 
  
  saLastError = 0 
  
  hr = SafeArrayGetVartype(*psa, @vartype) 
  If hr = #S_OK 
    ProcedureReturn vartype 
  Else 
    saLastError = hr 
    ProcedureReturn 0 
  EndIf 
    
EndProcedure 

; *************************************************************************************** 

Procedure.l saCount(*psa.safearray) ; Result Count of Elements 
  
Protected result.l 
  
  If *psa 
    result = *psa\rgsabound\cElements 
  Else 
    result = 0 
  EndIf 
  
  ProcedureReturn result 
  
EndProcedure 

; *************************************************************************************** 

Procedure.l saLBound(*psa.safearray) ; Result first number of Array 
  
  Shared saLastError 
  
  Protected hr, result 
  
  saLastError = 0 
  
  hr = SafeArrayGetLBound_(*psa, 1, @result) 
  If hr = #S_OK 
    ProcedureReturn result 
  Else 
    saLastError = hr 
    ProcedureReturn 0 
  EndIf 
  
EndProcedure 

; *************************************************************************************** 

Procedure.l saUBound(*psa.safearray) ; Result last number of Array 
  
  Shared saLastError 
  
  Protected hr, result 
  
  saLastError = 0 
  
  hr = SafeArrayGetUBound_(*psa, 1, @result) 
  If hr = #S_OK 
    ProcedureReturn result 
  Else 
    saLastError = hr 
    ProcedureReturn 0 
  EndIf 
  
EndProcedure 

; *************************************************************************************** 


;- Type Conversion Helper 

; *************************************************************************************** 

;-T_BSTR 
Procedure helpSysAllocString(*Value) 
  ProcedureReturn SysAllocString_(*Value) 
EndProcedure 
Prototype.l ProtoSysAllocString(Value.p-unicode) 

Global T_BSTR.ProtoSysAllocString = @helpSysAllocString() 

; *************************************************************************************** 

Procedure.d T_DATE(pbDate) ; Result Date from PB-Date 
  
  Protected date.d 
  
  date = pbDate / 86400.0 + 25569.0 
  ProcedureReturn date 
  
EndProcedure 

; *************************************************************************************** 

Procedure T_BOOL(Assert) ; Result Variant Type Boolean 

  If Assert 
    ProcedureReturn #VARIANT_TRUE 
  Else 
    ProcedureReturn #VARIANT_FALSE 
  EndIf 
  
EndProcedure 

; *************************************************************************************** 


;- Conversion Variant to PB Values 

; *************************************************************************************** 

Procedure.s VT_STR(*Var.Variant) 

  Shared vhLastError.l 

  Protected hr.l, result.s, VarDest.Variant 
  
  vhLastError = 0 
  
  If *Var 
    hr = VariantChangeType_(VarDest, *Var, 0, #VT_BSTR) 
    If hr = #S_OK 
      result = PeekS(VarDest\bstrVal, #PB_Any, #PB_Unicode) 
      VariantClear_(VarDest) 
      ProcedureReturn result 
    
    Else 
      vhLastError = hr 
      ProcedureReturn "" 
    EndIf 
    
  EndIf 
EndProcedure 
  
; *************************************************************************************** 
  
Procedure.l VT_BOOL(*Var.Variant) 

  Shared vhLastError.l 

  Protected hr.l, result.l, VarDest.Variant 
  
  vhLastError = 0 
  
  If *Var 
    hr = VariantChangeType_(VarDest, *Var, 0, #VT_BOOL) 
    If hr = #S_OK 
      result = VarDest\boolVal 
      VariantClear_(VarDest) 
      If result 
        ProcedureReturn #True 
      Else 
        ProcedureReturn #False 
      EndIf 
      
    Else 
      vhLastError = hr 
      ProcedureReturn 0 
    EndIf 
    
  EndIf 

EndProcedure 
  
; *************************************************************************************** 

Procedure.b VT_BYTE(*Var.Variant) 

  Shared vhLastError.l 

  Protected hr.l, result.b, VarDest.Variant 
  
  vhLastError = 0 
  
  If *Var 
    hr = VariantChangeType_(VarDest, *Var, 0, #VT_I1) 
    If hr = #S_OK 
      result = VarDest\bVal 
      VariantClear_(VarDest) 
      ProcedureReturn result 
    
    Else 
      vhLastError = hr 
      ProcedureReturn 0 
    EndIf 
    
  EndIf 

EndProcedure 
  
; *************************************************************************************** 

Procedure.w VT_WORD(*Var.Variant) 

  Shared vhLastError.l 

  Protected hr.l, result.w, VarDest.Variant 
  
  vhLastError = 0 
  
  If *Var 
    hr = VariantChangeType_(VarDest, *Var, 0, #VT_I2) 
    If hr = #S_OK 
      result = VarDest\iVal 
      VariantClear_(VarDest) 
      ProcedureReturn result 
    
    Else 
      vhLastError = hr 
      ProcedureReturn 0 
    EndIf 
    
  EndIf 

EndProcedure 
  
; *************************************************************************************** 

Procedure.l VT_LONG(*Var.Variant) 

  Shared vhLastError.l 

  Protected hr.l, result.l, VarDest.Variant 
  
  vhLastError = 0 
  
  If *Var 
    hr = VariantChangeType_(VarDest, *Var, 0, #VT_I4) 
    If hr = #S_OK 
      result = VarDest\lVal 
      VariantClear_(VarDest) 
      ProcedureReturn result 
    
    Else 
      vhLastError = hr 
      ProcedureReturn 0 
    EndIf 
    
  EndIf 

EndProcedure 
  
; *************************************************************************************** 

Procedure.q VT_QUAD(*Var.Variant) 
  
  Shared vhLastError.l 

  Protected hr.l, result.q, VarDest.Variant 
  
  vhLastError = 0 
  
  If *Var 
    hr = VariantChangeType_(VarDest, *Var, 0, #VT_I8) 
    If hr = #S_OK 
      result = VarDest\llVal 
      VariantClear_(VarDest) 
      ProcedureReturn result 
    
    Else 
      vhLastError = hr 
      ProcedureReturn 0 
    EndIf 
    
  EndIf 

EndProcedure 
  
; *************************************************************************************** 

Procedure.f VT_FLOAT(*Var.Variant) 

  Shared vhLastError.l 

  Protected hr.l, result.f, VarDest.Variant 
  
  vhLastError = 0 
  
  If *Var 
    hr = VariantChangeType_(VarDest, *Var, 0, #VT_R4) 
    If hr = #S_OK 
      result = VarDest\fltVal 
      VariantClear_(VarDest) 
      ProcedureReturn result 
    
    Else 
      vhLastError = hr 
      ProcedureReturn 0 
    EndIf 
    
  EndIf 

EndProcedure 
  
; *************************************************************************************** 

Procedure.d VT_DOUBLE(*Var.Variant) 

  Shared vhLastError.l 

  Protected hr.l, result.d, VarDest.Variant 
  
  vhLastError = 0 
  
  If *Var 
    hr = VariantChangeType_(VarDest, *Var, 0, #VT_R8) 
    If hr = #S_OK 
      result = VarDest\dblVal 
      VariantClear_(VarDest) 
      ProcedureReturn result 
    
    Else 
      vhLastError = hr 
      ProcedureReturn 0 
    EndIf 
    
  EndIf 

EndProcedure 
  
; *************************************************************************************** 

Procedure.l VT_DATE(*Var.Variant) ; Result PB-Date from Variant Date 

  Shared vhLastError.l 
  
  Protected pbDate 
  
  Protected hr.l, result.d, VarDest.Variant 
  
  vhLastError = 0 
  
  If *Var 
    hr = VariantChangeType_(VarDest, *Var, 0, #VT_DATE) 
    If hr = #S_OK 
      pbDate = (VarDest\dblVal  - 25569.0) * 86400.0 
      VariantClear_(VarDest) 
      ProcedureReturn pbDate 
    Else 
      vhLastError = hr 
      ProcedureReturn 0 
    EndIf 
    
  EndIf 

EndProcedure 
  
; *************************************************************************************** 

Procedure.l VT_ARRAY(*Var.Variant) ; Result a Pointer to SafeArray 
  
  Protected result.l 
  
  vhLastError = 0 
  
  If *Var 
    If (*Var\vt & #VT_ARRAY) = #VT_ARRAY 
      result = *Var\parray 
    Else 
      result = 0 
    EndIf 
  Else 
    result = 0 
  EndIf 
  ProcedureReturn result 
  
EndProcedure 

; *************************************************************************************** 


;- Converions PB Values to Variant 

; *************************************************************************************** 

Macro V_EMPTY(Arg) 
  VariantClear_(Arg) 
EndMacro 

; *************************************************************************************** 

Macro V_NULL(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_NULL 
  Arg\llVal 
EndMacro 

; *************************************************************************************** 

Macro V_DISP(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_DISPATCH 
  Arg\ppdispVal 
EndMacro 

; *************************************************************************************** 

Macro V_UNKNOWN(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_UNKNOWN 
  Arg\punkVal 
EndMacro 

; *************************************************************************************** 

Macro V_STR(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BSTR 
  Arg\bstrVal 
EndMacro 

; *************************************************************************************** 

Macro V_BOOL(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BOOL 
  Arg\boolVal 
EndMacro 

; *************************************************************************************** 

Macro V_BYTE(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_I1 
  Arg\bVal 
EndMacro 

; *************************************************************************************** 

Macro V_UBYTE(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_UI1 
  Arg\bVal 
EndMacro 

; *************************************************************************************** 

Macro V_WORD(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_I2 
  Arg\iVal 
EndMacro 

; *************************************************************************************** 

Macro V_UWORD(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_UI2 
  Arg\iVal 
EndMacro 

; *************************************************************************************** 

Macro V_LONG(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_I4 
  Arg\lVal 
EndMacro 

; *************************************************************************************** 

Macro V_ULONG(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_UI4 
  Arg\lVal 
EndMacro 

; *************************************************************************************** 

Macro V_QUAD(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_I8 
  Arg\llVal 
EndMacro 

; *************************************************************************************** 

Macro V_FLOAT(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_R4 
  Arg\fltVal 
EndMacro 

; *************************************************************************************** 

Macro V_DOUBLE(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_R8 
  Arg\dblVal 
EndMacro 

; *************************************************************************************** 

Macro V_DATE(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_DATE 
  Arg\dblVal 
EndMacro 

; *************************************************************************************** 

Macro V_VARIANT(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_VARIANT 
  Arg\pvarVal 
EndMacro 

; *************************************************************************************** 
Macro V_NULL_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_NULL 
  Arg\pllVal 
EndMacro 

; *************************************************************************************** 

Macro V_DISP_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_DISPATCH 
  Arg\ppdispVal 
EndMacro 

; *************************************************************************************** 

Macro V_UNKNOWN_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_UNKNOWN 
  Arg\ppunkVal 
EndMacro 

; *************************************************************************************** 

Macro V_STR_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_BSTR 
  Arg\pbstrVal 
EndMacro 

; *************************************************************************************** 

Macro V_BOOL_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_BOOL 
  Arg\pboolVal 
EndMacro 

; *************************************************************************************** 

Macro V_BYTE_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_I1 
  Arg\pbVal 
EndMacro 

; *************************************************************************************** 

Macro V_UBYTE_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_UI1 
  Arg\pbVal 
EndMacro 

; *************************************************************************************** 

Macro V_WORD_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_I2 
  Arg\piVal 
EndMacro 

; *************************************************************************************** 

Macro V_UWORD_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_UI2 
  Arg\piVal 
EndMacro 

; *************************************************************************************** 

Macro V_LONG_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_I4 
  Arg\plVal 
EndMacro 

; *************************************************************************************** 

Macro V_ULONG_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_UI4 
  Arg\plVal 
EndMacro 

; *************************************************************************************** 

Macro V_QUAD_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_I8 
  Arg\pllVal 
EndMacro 

; *************************************************************************************** 

Macro V_FLOAT_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_R4 
  Arg\pfltVal 
EndMacro 

; *************************************************************************************** 

Macro V_DOUBLE_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_R8 
  Arg\pdblVal 
EndMacro 

; *************************************************************************************** 

Macro V_DATE_BYREF(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_BYREF | #VT_DATE 
  Arg\pdblVal 
EndMacro 

; *************************************************************************************** 


;- Conversion SafeArray 

; *************************************************************************************** 

Macro V_ARRAY_DISP(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY |#VT_DISPATCH 
  Arg\ppdispVal 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_STR(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_BSTR 
  Arg\parray 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_BOOL(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_BOOL 
  Arg\parray 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_BYTE(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_I1 
  Arg\parray 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_UBYTE(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_UI1 
  Arg\parray 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_WORD(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_I2 
  Arg\parray 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_UWORD(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_UI2 
  Arg\parray 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_LONG(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_I4 
  Arg\parray 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_ULONG(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_UI4 
  Arg\parray 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_QUAD(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_I8 
  Arg\parray 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_FLOAT(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_R4 
  Arg\parray 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_DOUBLE(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_R8 
  Arg\parray 
EndMacro 

; *************************************************************************************** 

Macro V_ARRAY_DATE(Arg) 
  VariantClear_(Arg) 
  Arg\vt = #VT_ARRAY | #VT_DATE 
  Arg\parray 
EndMacro 

; *************************************************************************************** 


;- Macros for Safearray to get and put values 

; *************************************************************************************** 

Macro SA_BYTE(psa, index) 
  psa#\pvData\bVal[index-psa#\rgsabound\lLbound] 
EndMacro 

; *************************************************************************************** 

Macro SA_WORD(psa, index) 
  psa#\pvDataiVal[index-psa#\rgsabound\lLbound] 
EndMacro 

; *************************************************************************************** 

Macro SA_LONG(psa, index) 
  psa#\pvData\lVal[index-psa#\rgsabound\lLbound] 
EndMacro 

; *************************************************************************************** 

Macro SA_FLOAT(psa, index) 
  psa#\pvData\fltVal[index-psa#\rgsabound\lLbound] 
EndMacro 

; *************************************************************************************** 

Macro SA_DOUBLE(psa, index) 
  psa#\pvData\dblVal[index-psa#\rgsabound\lLbound] 
EndMacro 

; *************************************************************************************** 

Macro SA_DATE(psa, index) 
  psa#\pvData\dblVal[index-psa#\rgsabound\lLbound] 
EndMacro 

; *************************************************************************************** 

Macro SA_BSTR(psa, index) 
  psa#\pvData\bStrVal[index-psa#\rgsabound\lLbound] 
EndMacro 

; *************************************************************************************** 

Procedure.s SA_STR(*psa.safearray, index) ; Result PB-String from SafeArray BSTR 
  Protected *BSTR 
  *BSTR = *psa\pvData\bStrVal[index-*psa\rgsabound\lLbound] 
  ProcedureReturn PeekS(*BSTR, #PB_Any, #PB_Unicode) 
EndProcedure 

; *************************************************************************************** 

Macro SA_VARIANT(psa, index) 
  psa#\pvData\Value[index-psa#\rgsabound\lLbound] 
EndMacro 

; *************************************************************************************** 

Macro SA_DISPATCH(psa, index) 
  psa#\pvData\pdispVal[index-psa#\rgsabound\lLbound] 
EndMacro 

; *************************************************************************************** 


; IDE Options = PureBasic 4.61 Beta 1 (Windows - x86)
; CursorPosition = 68
; FirstLine = 35
; Folding = --------------
; EnableUnicode