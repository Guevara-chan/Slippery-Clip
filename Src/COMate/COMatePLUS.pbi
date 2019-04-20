CompilerIf Defined(INCLUDE_COMATE, #PB_Constant)=0
#INCLUDE_COMATE=1
;/////////////////////////////////////////////////////////////////////////////////
;***COMate***  COM automation through iDispatch.
;*===========
;*
;*COMatePLUS.  Version 1.2 released 09th July 2010.
;*
;*©nxSoftWare (www.nxSoftware.com) 2009.
;*======================================
;*   With thanks to ts-soft, kiffi, mk-soft.
;*   The EventSink code is based on that produced by Freak : http://www.purebasic.fr/english/viewtopic.php?t=26744&postdays=0&postorder=asc&start=75
;*   Created with Purebasic 4.3 for Windows.
;*
;*   Platforms:  Windows.
;/////////////////////////////////////////////////////////////////////////////////

;/////////////////////////////////////////////////////////////////////////////////
;*NOTES.
; i)    This code has arisen from my study of COM automation; the mechanism through which applications can
;       connect to COM servers through what is termed 'late binding'.
;
; ii)   At present this can only be used for servers on the local machine.

; iii)  This code is based upon the DispHelper sourcecode : http://disphelper.sourceforge.net/.
;
; iv)   Define the constant #COMATE_NOINCLUDEATL = 1 before including this source to remove all ActiveX code from this library.
;       Useful for NT in which the ATL library may not be present.
;
; v)    Define the constant #COMATE_NOERRORREPORTING = 1 before including this source to remove all error reporting. Might be useful if
;       looking to squeeze a little more speed out of your code!
;/////////////////////////////////////////////////////////////////////////////////

CompilerIf #PB_Compiler_Version < 530 ; In case of outdated (as in, LTS) compiler...
XIncludeFile "COMate\COMatePLUS_Residents.pbi" : CompilerElse : XIncludeFile "COMatePLUS_Residents.pbi"
CompilerEndIf

;-IMPORTS.
  CompilerIf Defined(COMATE_NOINCLUDEATL, #PB_Constant)=0
    Import "atl.lib"
      AtlAxCreateControl(lpszName,hWnd.i,*pStream.IStream,*ppUnkContainer.IUnknown)
      AtlAxGetControl(hWnd.i,*pp.IUnknown)
      AtlAxGetHost(hWnd, *pp.IUnknown)
      AtlAxWinInit()
    EndImport
  CompilerEndIf

;-PROTOTYPES.
  ;The following prototype caters for Ansi / Unicode when creating BSTR's.
    Prototype.i COMate_ProtoMakeBSTR(value.p-unicode)
  ;The following is for any automation servers opting to defer filling in EXCEPINFO2 structures in the case of a dispinterface
  ;function yielding an error etc. In these cases a callback is provided by the server which we call manually.
    Prototype.i COMate_ProtoDeferredFillIn(*EXCEPINFO2)
  ;The following prototypes allow for various return types from event handlers.
    Prototype COMate_EventCallback_NORETURN(COMateObject.COMateObject, EventName$, ParameterCount)
    Prototype.q COMate_EventCallback_INTEGERRETURN(COMateObject.COMateObject, EventName$, ParameterCount)
    Prototype.d COMate_EventCallback_REALRETURN(COMateObject.COMateObject, EventName$, ParameterCount)
    Prototype.s COMate_EventCallback_STRINGRETURN(COMateObject.COMateObject, EventName$, ParameterCount)
    Prototype COMate_EventCallback_UNKNOWNRETURN(COMateObject.COMateObject, EventName$, ParameterCount, *returnValue.VARIANT)

;-CONSTANTS (private)
  #COMate_MAXNUMSUBOBJECTS      = 20      ;Used for nested object calls; e.g. "Cells(1, 2)\Value = 'COMate'"
  #COMate_MAXNUMSYMBOLSINALINE  = 200
  #COMate_MAXNUMVARIANTARGS     = 20       ;The max number of arguments which can be passed to a single COM method.
  Enumeration 
    #CLSCTX_INPROC_SERVER  = 1 
    #CLSCTX_INPROC_HANDLER = 2 
    #CLSCTX_LOCAL_SERVER   = 4 
    #CLSCTX_REMOTE_SERVER  = 16 
    #CLSCTX_FROM_DEFAULT_CONTEXT = $20000
    #CLSCTX_SERVER = (#CLSCTX_INPROC_SERVER | #CLSCTX_LOCAL_SERVER | #CLSCTX_REMOTE_SERVER) 
  EndEnumeration 

  Enumeration ;Used when parsing command strings and setting up the variant array.
    #COMate_Operator
    #COMate_Operand
    #COMate_OpenParanthesis
    #COMate_CloseParanthesis
    #COMate_Method
  EndEnumeration

  #DISPID_PROPERTYPUT = -3 ;The iDisp value for property put calls to iDispatch\Invoke() which require a single named parameter.
  #DISPID_NEWENUM = -4     ;The iDisp value for propertyget calls in which a new enumeration is being requested.

  #CONNECT_E_ADVISELIMIT = -2147220991

;-STRUCTURES.

  ;The following structure contains the class template and private properties for the main COMateObject.
    Structure _membersCOMateClass
      *vTable
      iDisp.iDispatch
      containerID.i
      hWnd.i
      *eventSink._COMateEventSink
    EndStructure 

  ;The following structure contains the class template and private properties for the COMateEnumObject.
    Structure _membersCOMateEnumClass
      *vTable
      *parent._membersCOMateClass   ;Points to the COMate Object which is hosting this enumeration. Used for error reporting.
      iEV.IEnumVARIANT
    EndStructure 

  ;The following structure is used in thread local storage to store info on the latest error recorded by an object within the current thread.
    Structure _COMateThreadErrors
      lastErrorCode.i
      lastError$
    EndStructure

  ;The following structure holds a COMatePLUS 'statement' object representing a compiled command string.
  ;A statement handle is simply a pointer to one of these structures.
    Structure _COMatePLUSStatement
      numSubObjects.i
      methodName.i[#COMate_MAXNUMSUBOBJECTS+1]  ;1-based indexing. BSTRs.
      numArgs.i[#COMate_MAXNUMSUBOBJECTS+1]     ;1-based indexing.
      ptrVarArgs.i[#COMate_MAXNUMSUBOBJECTS+1]  ;1-based indexing.
    EndStructure

  ;The following structure is used in an array when parsing method parameters etc.
    Structure _COMateParse
      numberOfTokens.i
      numOpenBrackets.i
      numCloseBrackets.i
      tokens$[#COMate_MAXNUMSYMBOLSINALINE]
    EndStructure

  ;The following structure is used in the iDispatch\Invoke() method call to receive detailed errors.
  CompilerIf Defined(EXCEPINFO2, #PB_Structure) = 0
     CompilerIf #PB_Compiler_Processor = #PB_Processor_x64
        Structure EXCEPINFO2
         wCode.w
         wReserved.w
         pad.b[4] ; Only on x64
         bstrSource.i                ;BSTR
         bstrDescription.i
         bstrHelpFile.i
         dwHelpContext.l
         pvReserved.i
         pfnDeferredFillIn.COMate_ProtoDeferredFillIn
         scode.l
         pad2.b[4] ; Only on x64
      EndStructure
    CompilerElse
      Structure EXCEPINFO2
         wCode.w
         wReserved.w
         bstrSource.i                ;BSTR
         bstrDescription.i
         bstrHelpFile.i
         dwHelpContext.l
         pvReserved.i
         pfnDeferredFillIn.COMate_ProtoDeferredFillIn
         scode.l
      EndStructure      
    CompilerEndIf
  CompilerEndIf

  ;The following structure is used when connecting an outgoing interface (sink) to a COM object's connection point.
    Structure _COMateEventSink
      *Vtbl
      refCount.i 
      cookie.i
      connIID.IID
      typeInfo.ITypeInfo
      Callback.COMate_EventCallback_NORETURN
      returnType.i
      *dispParams.DISPPARAMS
      *parent._membersCOMateClass     ;A pointer back to the parent COMate object so that we can pass this to the event procedure.
    EndStructure


;-MACROS
  
  ;The following two macros are used to test for success or failure when calling com methods.
  ;They are pretty superfluous really but do aid readability.
    Macro SUCCEEDED(HRESULT)
      HRESULT & $80000000 = 0
    EndMacro
    Macro FAILED(HRESULT)
      HRESULT & $80000000
    EndMacro
  

;-DECLARES.
  Declare.i COMate_INTERNAL_CheckNumeric(arg$, *var.VARIANT)
  Declare COMate_INTERNAL_EscapeString(ptrText)
  Declare.i COMateClass_INTERNAL_InvokePlus(*this._membersCOMateClass, invokeType, returnType, *ret.VARIANT, command$, *hStatement._COMatePLUSStatement)
  Declare COMate_INTERNAL_FreeStatementHandle(*hStatement._COMatePLUSStatement)
  Declare.i COMate_INTERNAL_PrepareStatement(command$, *ptrStatement.INTEGER)
  Declare.i COMatePLUS_TokeniseCommand(command$, separator$, Array parse._COMateParse(1))
  Declare.i COMatePLUS_CompileSubobjectInvokation(*hStatement._COMatePLUSStatement, subObjectIndex, Array parse._COMateParse(1))

  Declare COMateClass_INTERNAL_SetError(*this._membersCOMateClass, result, blnAllowDispError = 0, dispError$="")
  Declare.i COMateClass_UTILITY_MakeBSTR(value)

  Declare.i COMateClass_GetObjectProperty(*this._membersCOMateClass, command$, *hStatement=0, objectType = #VT_DISPATCH)

  CompilerIf Defined(COMATE_NOINCLUDEATL, #PB_Constant)=0
    Declare.i COMate_DelSinkPropsCallback(hWnd, lpszString, hData,dwData)
  CompilerEndIf

;-GLOBALS.
    Global COMate_MakeBSTR.COMate_ProtoMakeBSTR = @COMateClass_UTILITY_MakeBSTR()  ;Prototype.
    Global COMate_gErrorTLS.i  ;A TLS index used to store per-thread error info.
    Global COMate_gNumObjects.i  ;Used to manage the error-TLS index.
    Global COMate_gPtrThreadArray.i   ;A pointer to an array of pointers to _COMateThreadErrors structures. 
    Global COMate_gNumThreadElements.i
    Global COMate_gAtlAXIsInit.i

;-=======================
;-COMate OBJECT CODE.
;-=======================

;/////////////////////////////////////////////////////////////////////////////////
;The following function creates a new instance of a COMate object which itself contains a COM object (iDispatch).
;Change the optional parameter blnInitCOM to #False if COM has already been initialised.
;Returns the new COMate object or zero if an error.
Procedure.i COMate_CreateObject(progID$, hWnd = 0, blnInitCOM = #True)
  Protected *this._membersCOMateClass, clsid.CLSID, hResult, cf.IClassFactory, progID, container.iUnknown, iDisp
  If blnInitCOM 
    CoInitialize_(0)
  EndIf
  If progID$
    progID = COMate_MakeBSTR(progID$)
    If progID
      *this = AllocateMemory(SizeOf(_membersCOMateClass))
      If *this
        *this\vTable = ?VTable_COMateClass
        If hWnd = 0 ;No ActiveX control to house.
          ;Get classID from the registry.
            If Left(progID$, 1) = "{"
              hResult = CLSIDFromString_(progID, @clsid)
              If SUCCEEDED(hResult)
                hResult = ProgIDFromCLSID_(clsid, @iDisp)
                If SUCCEEDED(hResult) And iDIsp
                  SysFreeString_(iDisp)
                EndIf
              EndIf
            Else
	  	        hResult = CLSIDFromProgID_(progID, @clsid);
	          EndIf
	        If SUCCEEDED(hResult)
            hResult = CoGetClassObject_(clsid, #CLSCTX_LOCAL_SERVER|#CLSCTX_INPROC_SERVER, 0, ?IID_IClassFactory, @cf)
	          If SUCCEEDED(hResult)
              hResult = cf\CreateInstance(0, ?IID_IDispatch, @*this\iDisp)
              If FAILED(hResult)
                hResult = cf\CreateInstance(0, ?IID_IUnknown, @container)
	              If SUCCEEDED(hResult)
                  hResult = container\QueryInterface(?IID_IDispatch, @*this\iDisp)
                  container\Release()
                EndIf
              EndIf
              If FAILED(hResult)
                FreeMemory(*this)
                *this = 0
              Else; Success.
                COMate_gNumObjects+1
              EndIf
            Else
              FreeMemory(*this)
              *this = 0
            EndIf
            If cf
              cf\Release()
            EndIf
          Else
            FreeMemory(*this)
            *this = 0
          EndIf
CompilerIf Defined(COMATE_NOINCLUDEATL, #PB_Constant)=0
        Else  ;An ActiveX control requires housing.
          ;Get classID from the registry. This is a simple check to ensure the control is registered. Otherwise ATL will embed a browser in our container.
            If Left(progID$, 1) = "{"
              hResult = CLSIDFromString_(progID, @clsid)
              If SUCCEEDED(hResult)
                hResult = ProgIDFromCLSID_(clsid, @iDisp)
                If SUCCEEDED(hResult) And iDIsp
                  SysFreeString_(iDisp)
                EndIf
              EndIf
            Else
	  	        hResult = CLSIDFromProgID_(progID, @clsid);
	          EndIf
          If SUCCEEDED(hResult)
            If COMate_gAtlAXIsInit = #False
              If AtlAxWinInit()
                COMate_gAtlAXIsInit = #True
              Else
                hresult = #E_FAIL
                FreeMemory(*this)
                *this = 0
              EndIf
            EndIf
            If COMate_gAtlAXIsInit
              hResult = AtlAxCreateControl(ProgId, hWnd, 0, 0)
              If SUCCEEDED(hResult)
                hResult = AtlAxGetControl(hWnd, @*this\iDisp)
                If SUCCEEDED(hResult)
                  hresult = *this\iDisp\QueryInterface(?IID_IDispatch, @iDisp)
                  *this\iDisp\Release()
                  If SUCCEEDED(hresult)
                    *this\hWnd = hWnd
                    *this\iDisp = iDisp
                    COMate_gNumObjects+1
                  Else
                    FreeMemory(*this)
                    *this = 0
                  EndIf  
                Else
                  FreeMemory(*this)
                  *this = 0
                EndIf
              Else
                FreeMemory(*this)
                *this = 0
              EndIf
            EndIf
          Else
            FreeMemory(*this)
            *this = 0
          EndIf
CompilerEndIf
        EndIf
      Else
        hresult = #E_OUTOFMEMORY
      EndIf
      SysFreeString_(progID)
    Else
      hresult = #E_OUTOFMEMORY
    EndIf
  Else
    hresult = #E_INVALIDARG
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, hResult)
  CompilerEndIf
  ProcedureReturn *this
EndProcedure
;=================================================================================


;/////////////////////////////////////////////////////////////////////////////////
;The following function is used either to load an instance of a com object from a file (or based upon the filename given), or to
;create a new instance of a currently active object.
;If file$ is empty, the function attempts to create a new COMate object containing a new instance of a currently active object
;(the existing COM object's reference count is increased).
;If file$ is not empty, progID$ is used to specify the class of the object in cases where the file contains multiple objects.
;(This mimicks VB's GetObject() function.)
;Returns the new COMate object or zero if an error.
Procedure.i COMate_GetObject(file$, progID$="", blnInitCOM = #True)
  Protected *this._membersCOMateClass, hResult = #E_OUTOFMEMORY, iPersist.IPERSISTFILE, clsid.CLSID, cf.IClassFactory, iUnknown.IUNKNOWN
  Protected bstr1, t1
  If blnInitCOM 
    CoInitialize_(0)
  EndIf
  If file$ Or progID$
    *this = AllocateMemory(SizeOf(_membersCOMateClass))
    If *this
      *this\vTable = ?VTable_COMateClass
      If file$
        If progID$ = ""
          ;Here we attempt to create an object based upon the filename only.
            bstr1 = COMate_MakeBSTR(file$)
            If bstr1 ;If an error then hResult already equals #E_OUTOFMEMORY!
              hResult = CoGetObject_(bstr1, 0, ?IID_IDispatch, @*this\iDisp)
              SysFreeString_(bstr1)
            EndIf
        Else
          ;Here we attempt to create an object based upon the filename and the progID.
            bstr1 = COMate_MakeBSTR(progID$)
            If bstr1
              hResult = CLSIDFromProgID_(bstr1, @clsid)
              If SUCCEEDED(hResult)
                hResult = CoGetClassObject_(clsid, #CLSCTX_LOCAL_SERVER|#CLSCTX_INPROC_SERVER, 0, ?IID_IClassFactory, @cf)
                If SUCCEEDED(hResult)
                  hResult = cf\CreateInstance(0, ?IID_IPersistFile, @iPersist)
                  If SUCCEEDED(hResult)
                    hResult = iPersist\Load(file$, 0)
                    If SUCCEEDED(hResult)
                      hResult = iPersist\QueryInterface(?IID_IDispatch, @*this\iDisp)
                    EndIf
                  EndIf
                  If iPersist
                    iPersist\Release()
                  EndIf
                EndIf
                If cf
                  cf\Release()
                EndIf
              EndIf
            EndIf
            If bstr1
              SysFreeString_(bstr1)
            EndIf
        EndIf
      Else
        ;Here we attempt to create a new COMate object containing a new instance of a currently active object.
          bstr1 = COMate_MakeBSTR(progID$)
          If bstr1
            If Left(progID$, 1) = "{"
              hResult = CLSIDFromString_(bstr1, @clsid)
              If SUCCEEDED(hResult)
                hResult = ProgIDFromCLSID_(clsid, @t1)
                If SUCCEEDED(hResult) And t1
                  SysFreeString_(t1)
                EndIf
              EndIf
            Else
	  	        hResult = CLSIDFromProgID_(bstr1, @clsid);
	          EndIf
            If SUCCEEDED(hResult)
              hResult = GetActiveObject_(clsid, 0, @iUnknown)
              If SUCCEEDED(hResult)
                hResult = iUnknown\QueryInterface(?IID_IDispatch, @*this\iDisp)
              EndIf
              If iUnknown
                iUnknown\Release()
              EndIf
            EndIf
            SysFreeString_(bstr1)
          EndIf
      EndIf
    EndIf
  Else
    hResult = #E_INVALIDARG
  EndIf
  If SUCCEEDED(hResult)
    COMate_gNumObjects+1
  ElseIf *this
    FreeMemory(*this)
    *this = 0
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, hResult)
  CompilerEndIf
  ProcedureReturn *this
EndProcedure
;=================================================================================


;/////////////////////////////////////////////////////////////////////////////////
;The following function creates a new instance of a COMate object from an object supplied directly from the user.
;This object is in the form of a iUnknown pointer to which we use QueryInterface() in an attempt to locate an iDispatch pointer.
;Useful for event procedures attached to ActiveX controls in which some parameters may be a 'raw' COM object. This function can be used to package that
;object up into the form of a COMate object.
;Returns the new COMate object or zero if an error.
Procedure.i COMate_WrapCOMObject(object.iUnknown)
  Protected *this._membersCOMateClass, hResult, iDisp.iUnknown
  If object
    hresult = object\QueryInterface(?IID_IDispatch, @iDisp)
    If SUCCEEDED(hresult)
      *this = AllocateMemory(SizeOf(_membersCOMateClass))
      If *this
        *this\vTable = ?VTable_COMateClass
        *this\iDisp = iDisp
        COMate_gNumObjects+1
      Else
        hresult = #E_OUTOFMEMORY
        iDisp\Release()
      EndIf
    EndIf
  Else
    hresult = #E_INVALIDARG
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, hResult)
  CompilerEndIf
  ProcedureReturn *this
EndProcedure
;=================================================================================


;/////////////////////////////////////////////////////////////////////////////////
;The following function creates a new instance of a COMate object which itself contains a COM object (iDispatch) representing an
;ActiveX server. The underlying ActiveX control is placed within a container gadget.
;Change the optional parameter blnInitCOM to #False if COM has already been initialised.
;Returns the new COMate object or zero if an error.
Procedure.i COMate_CreateActiveXControl(x, y, width, height, progID$, blnInitCOM = #True)
  Protected *this._membersCOMateClass, hResult, id, hWnd, iDisp
CompilerIf Defined(COMATE_NOINCLUDEATL, #PB_Constant)=0
  If progID$
    id = ContainerGadget(#PB_Any, x, y, width, 	height)
    CloseGadgetList()
    If id
      hWnd = GadgetID(id)
      *this = COMate_CreateObject(progID$, hWnd, blnInitCOM) ;This procedure will set any HRESULT codes.
      If *this
        SetWindowLong_(hWnd, #GWL_STYLE, GetWindowLong_(hWnd, #GWL_STYLE)|#WS_CLIPCHILDREN)
        *this\containerID = ID
        *this\hWnd = hWnd
      Else ;Cannot locate an iDispatch interface.
        FreeGadget(id)
      EndIf
      ProcedureReturn *this
    Else
      hResult = #E_OUTOFMEMORY
    EndIf
  Else
    hresult = #E_INVALIDARG
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, hResult)
  CompilerEndIf
CompilerEndIf
  ProcedureReturn 0
EndProcedure
;=================================================================================



;-COMate CLASS METHODS.
;----------------------------------------------

;=================================================================================
;The following method calls a dispinterface method where no return value is required.
;Returns a HRESULT value. #S_OK for no errors.
Procedure.i COMateClass_Invoke(*this._membersCOMateClass, command$, *hStatement=0)
  Protected result.i = #S_OK
  If command$ Or *hStatement
    result = COMateClass_INTERNAL_InvokePlus(*this, #DISPATCH_METHOD, #VT_EMPTY, 0, command$, *hStatement)
  Else
    result = #E_INVALIDARG
  EndIf
  ;Set any error code. iDispatch errors will already have been set.
    If result = -1
      result = #S_FALSE
    Else
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, result)
  CompilerEndIf
    EndIf
  ProcedureReturn result
EndProcedure
;=================================================================================


;=================================================================================
;The following method releases a com object created by any of the functions which return object pointers.
;Any sink interface connected to the underlying COM object will automatically be disconnected, resulting in the
;Release() method being called and able to tidy up.
Procedure COMateClass_Release(*this._membersCOMateClass)
  Protected *error._COMateThreadErrors, i.i, sink.IDispatch 
  If *this\iDisp ;Just in case.
    ;Release underlying iDispatch object.
      *this\iDisp\Release()
  EndIf
  If *this\containerID  ;OCX controls.
    FreeGadget(*this\containerID)
;We have to assume that the container will call the release() method on the connection point.
;  ElseIf *this\eventSink
;    sink = *this\eventSink
;    sink\Release()
  EndIf
  COMate_gNumObjects-1
  If COMate_gNumObjects = 0
    ;Here, in the anticipation that no more objects will be created, we release all memory associated with the TLS index.
    ;We recreate all this later on if required.
      If COMate_gErrorTLS <> -1
        For i = 0 To COMate_gNumThreadElements-1
          *error = PeekI(COMate_gPtrThreadArray + i*SizeOf(i))
          If *error ;Just in case!
            ClearStructure(*error, _COMateThreadErrors)
            FreeMemory(*error)
          EndIf
        Next
        FreeMemory(COMate_gPtrThreadArray)
        COMate_gPtrThreadArray = 0
        COMate_gNumThreadElements = 0
        TlsFree_(COMate_gErrorTLS)
        COMate_gErrorTLS = -1
      EndIf 
  EndIf
  ;Free object.
  FreeMemory(*this)
EndProcedure
;=================================================================================


;=================================================================================
;The following method creates a new instance of a COMateEnum object based upon an enumeration applied to the underlying COMate object.
;Returns the new COMateEnum object or zero if an error.
Procedure.i COMateClass_CreateEnumeration(*this._membersCOMateClass, command$, *hStatement=0)
  Protected result.i = #S_OK, *object._membersCOMateEnumClass, *tempCOMateObject._membersCOMateClass, iDisp.IDISPATCH
  Protected dp.DISPPARAMS, excep.EXCEPINFO2, var.VARIANT
  *object = AllocateMemory(SizeOf(_membersCOMateEnumClass))
  If *object
    *object\vTable = ?VTable_COMateEnumClass
    *object\parent = *this
    If command$
      *tempCOMateObject = COMateClass_GetObjectProperty(*this, command$, *hStatement, #VT_DISPATCH) ;This will set any error codes etc.
      If *tempCOMateObject
        iDisp = *tempCOMateObject\iDisp
      Else
        FreeMemory(*object)
        ProcedureReturn 0  ;Error codes already set.
      EndIf
    Else
      iDisp = *this\iDisp
    EndIf
    result = iDisp\Invoke(#DISPID_NEWENUM, ?IID_NULL, #LOCALE_USER_DEFAULT, #DISPATCH_METHOD | #DISPATCH_PROPERTYGET, dp, var, excep, 0)
    If command$
      COMateClass_Release(*tempCOMateObject)
    EndIf
    If SUCCEEDED(result)
      Select var\vt
        Case #VT_DISPATCH
		      result = var\pdispVal\QueryInterface(?IID_IEnumVARIANT, @*object\iEV)
        Case#VT_UNKNOWN
		      result = var\punkVal\QueryInterface(?IID_IEnumVARIANT, @*object\iEV)
        Default
          result = #E_NOINTERFACE;
      EndSelect
      If FAILED(result)
        FreeMemory(*object)
        *object = 0
      EndIf
    Else
      If result = #DISP_E_EXCEPTION
      ;Has the automation server deferred from filling in the EXCEPINFO2 structure?
        If excep\pfnDeferredFillIn
          excep\pfnDeferredFillIn(excep)
        EndIf
        If excep\bstrSource
          SysFreeString_(excep\bstrSource)
        EndIf          
        If excep\bstrDescription
          CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
            COMateClass_INTERNAL_SetError(*this, result,  #True, PeekS(excep\bstrDescription, -1, #PB_Unicode))
          CompilerEndIf
          SysFreeString_(excep\bstrDescription)
        Else
          CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
            COMateClass_INTERNAL_SetError(*this, result, #True)
          CompilerEndIf
        EndIf          
        If excep\bstrHelpFile
          SysFreeString_(excep\bstrHelpFile)
        EndIf
      EndIf
      FreeMemory(*object)
      *object = 0
    EndIf
    VariantClear_(var)
  Else
    result = #E_OUTOFMEMORY
  EndIf
  ;Set any error code.
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, result)
  CompilerEndIf
  ProcedureReturn *object
EndProcedure
;=================================================================================


;=================================================================================
;Returns the COMate object's underlying iDispatch object pointer.
;AddRef() is called on this object and so the developer must call Release() at some point.
Procedure.i COMateClass_GetCOMObject(*this._membersCOMateClass)
  Protected result.i = #S_OK, id.i
  *this\iDisp\AddRef()
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, #S_OK)
  CompilerEndIf
  ProcedureReturn *this\iDisp
EndProcedure
;=================================================================================


;=================================================================================
;The following method returns, in the case of an ActiveX control, either the handle of the container used to house the control
;or the Purebasic gadget#. Returning the gadget# is only viable if using COMate as a source code include (or a Tailbitten library!)
Procedure.i COMateClass_GetContainerhWnd(*this._membersCOMateClass, returnCtrlID=0)
  Protected result.i = #S_OK, id.i
CompilerIf Defined(COMATE_NOINCLUDEATL, #PB_Constant)=0
  If *this\hWnd
    id = *this\containerID
    If returnCtrlID = 0
      id = *this\hWnd
    EndIf
  Else
    result = #E_FAIL
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, result)
  CompilerEndIf
  ProcedureReturn id
CompilerEndIf
EndProcedure
;=================================================================================


;=================================================================================
;The following method attempts to set (or clear) the design time mode of the container.
Procedure.i COMateClass_SetDesignTimeMode(*this._membersCOMateClass, state=#True)
  Protected result.i = #S_OK, id, iUnk.IUnknown, iDisp.IDispatch, comate.COMateObject
CompilerIf Defined(COMATE_NOINCLUDEATL, #PB_Constant)=0
  If *this\containerID
    id = *this\containerID
    result = AtlAxGetHost(GadgetID(*this\containerID), @iUnk)
    If iUnk
      result = iUnk\QueryInterface(?IID_IAxWinAmbientDispatch, @iDisp)
      If iDisp
        comate = COMate_WrapCOMObject(iDisp)
        If comate
          If state
            result = comate\SetProperty("UserMode = #False")
          Else
            result = comate\SetProperty("UserMode = #True")
          EndIf
          comate\Release()
        Else
          result = COMate_GetLastErrorCode()
        EndIf
        iDisp\Release()
        iUnk\Release()
        ProcedureReturn result
      EndIf
      iUnk\Release()
    EndIf
  Else
    result = #E_FAIL
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, result)
  CompilerEndIf
CompilerEndIf
  ProcedureReturn result
EndProcedure
;=================================================================================


;=================================================================================
;The following method calls a dispinterface function and returns a PB (system) date value.
;Any HRESULT return value is accessible through the GetLastErrorCode() method.
Procedure.i COMateClass_GetDateProperty(*this._membersCOMateClass, command$, *hStatement=0)
  Protected result.i = #S_OK, retVar.VARIANT, retValue
  If command$ Or *hStatement
    result = COMateClass_INTERNAL_InvokePlus(*this, #DISPATCH_PROPERTYGET|#DISPATCH_METHOD, #VT_DATE, retVar, command$, *hStatement)
    If SUCCEEDED(result)
      retValue = (retVar\date - 25569) * 86400
    EndIf
    VariantClear_(retVar)
  Else
    result = #E_INVALIDARG
  EndIf
  ;Set any error code. iDispatch errors will alreay have been set.
    CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
      COMateClass_INTERNAL_SetError(*this, result)
    CompilerEndIf
  ProcedureReturn retValue
EndProcedure
;=================================================================================


;=================================================================================
;The following method calls a dispinterface function and returns an integer value.
;Any HRESULT return value is accessible through the GetLastErrorCode() method.
Procedure.q COMateClass_GetIntegerProperty(*this._membersCOMateClass, command$, *hStatement=0)
  Protected result.i = #S_OK, retVar.VARIANT, retValue.q
  If command$ Or *hStatement
    If OSVersion() <= #PB_OS_Windows_2000
      result = COMateClass_INTERNAL_InvokePlus(*this, #DISPATCH_PROPERTYGET|#DISPATCH_METHOD, #VT_I4, retVar, command$, *hStatement)
    Else
      result = COMateClass_INTERNAL_InvokePlus(*this, #DISPATCH_PROPERTYGET|#DISPATCH_METHOD, #VT_I8, retVar, command$, *hStatement)
    EndIf
    If SUCCEEDED(result)
      If OSVersion() <= #PB_OS_Windows_2000
        retValue = retVar\lval
      Else
        retValue = retVar\llval
      EndIf
    EndIf
    VariantClear_(retVar)
  Else
    result = #E_INVALIDARG
  EndIf
  ;Set any error code. iDispatch errors will alreay have been set.
    CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
      COMateClass_INTERNAL_SetError(*this, result)
    CompilerEndIf
  ProcedureReturn retValue
EndProcedure
;=================================================================================


;=================================================================================
;Returns a COMate object or an iUnknown interface pointer depending on the value of the 'objectType' parameter.
;For 'regular' objects based upon iDispatch, leave the optional parameter 'objectType' as it is. 
;Otherwise, for unknown object types set objectType to equal #COMate_UnknownObjectType. In these cases, this method will return the
;interface pointer directly (as opposed to a COMate object).
;In either case the object should be released as soon as it is no longer required.
;Any HRESULT return value is accessible through the GetLastErrorCode() method.
Procedure.i COMateClass_GetObjectProperty(*this._membersCOMateClass, command$, *hStatement=0, objectType = #VT_DISPATCH)
  Protected result.i = #S_OK, retVar.VARIANT, *newObject._membersCOMateClass
  If command$ Or *hStatement
    If objectType <> #VT_DISPATCH
      objectType = #VT_UNKNOWN
    EndIf
    result = COMateClass_INTERNAL_InvokePlus(*this, #DISPATCH_PROPERTYGET|#DISPATCH_METHOD, objectType, retVar, command$, *hStatement)
    If SUCCEEDED(result)
      If objectType = #VT_DISPATCH
        If retVar\pdispVal
          ;We now create a COMate object to house this instance variable.
            *newObject = AllocateMemory(SizeOf(_membersCOMateClass))
            If *newObject
              *newObject\vTable = ?VTable_COMateClass
              *newObject\iDisp = retVar\pdispVal
              COMate_gNumObjects+1
            Else
              VariantClear_(retVar)  ;This will call the Release() method of the COM object.
              result = #E_OUTOFMEMORY
            EndIf
        Else
          VariantClear_(retVar)
          ;In this case we set an error with extra info.
            CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
              COMateClass_INTERNAL_SetError(*this, #S_FALSE, 0, "The property returned a NULL object!")
            CompilerEndIf
            result = -1
        EndIf
      Else
        *newObject = retVar\punkVal
      EndIf
    Else
      VariantClear_(retVar)
    EndIf
  Else
    result = #E_INVALIDARG
  EndIf
  ;Set any error code. iDispatch errors will alreay have been set.
    If result <> -1
      CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
        COMateClass_INTERNAL_SetError(*this, result)
      CompilerEndIf
    EndIf
  ProcedureReturn *newObject
EndProcedure
;=================================================================================


;=================================================================================
;The following method calls a dispinterface function and returns a double value.
;Any HRESULT return value is accessible through the GetLastErrorCode() method.
Procedure.d COMateClass_GetRealProperty(*this._membersCOMateClass, command$, *hStatement=0)
  Protected result.i = #S_OK, retVar.VARIANT, retValue.d
  If command$ Or *hStatement
    result = COMateClass_INTERNAL_InvokePlus(*this, #DISPATCH_PROPERTYGET|#DISPATCH_METHOD, #VT_R8, retVar, command$, *hStatement)
    If SUCCEEDED(result)
      retValue = retVar\dblval
    EndIf
    VariantClear_(retVar)
  Else
    result = #E_INVALIDARG
  EndIf
  ;Set any error code. iDispatch errors will alreay have been set.
    CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
      COMateClass_INTERNAL_SetError(*this, result)
    CompilerEndIf
  ProcedureReturn retValue
EndProcedure
;=================================================================================


;=================================================================================
;The following method calls a dispinterface function and returns a string value.
;Any HRESULT return value is accessible through the GetLastErrorCode() method.
Procedure.s COMateClass_GetStringProperty(*this._membersCOMateClass, command$, *hStatement=0)
  Protected result.i = #S_OK, retVar.VARIANT, result$
  If command$ Or *hStatement
    result = COMateClass_INTERNAL_InvokePlus(*this, #DISPATCH_PROPERTYGET|#DISPATCH_METHOD, #VT_BSTR, retVar, command$, *hStatement)
    If SUCCEEDED(result) And retVar\bstrVal
      result$ = PeekS(retVar\bstrVal, -1, #PB_Unicode)
    EndIf
    VariantClear_(retVar)
  Else
    result = #E_INVALIDARG
  EndIf
  ;Set any error code. iDispatch errors will alreay have been set.
    CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
      COMateClass_INTERNAL_SetError(*this, result)
    CompilerEndIf
  ProcedureReturn result$
EndProcedure
;=================================================================================


;=================================================================================
;The following method calls a dispinterface function and, if there are no errors, returns a pointer to a new variant which must be
;'freed' by the user with VariantClear_() etc.
;Any HRESULT return value is accessible through the GetLastErrorCode() method.
Procedure.i COMateClass_GetVariantProperty(*this._membersCOMateClass, command$, *hStatement=0)
  Protected result.i = #S_OK, *retVar.VARIANT
  If command$ Or *hStatement
    ;Allocate memory for a new variant.
      *retVar = AllocateMemory(SizeOf(VARIANT))
    If *retVar
      result = COMateClass_INTERNAL_InvokePlus(*this, #DISPATCH_PROPERTYGET|#DISPATCH_METHOD, #VT_EMPTY, *retVar, command$, *hStatement)
      If FAILED(result)
        FreeMemory(*retVar)
        *retVar = 0
      EndIf
    Else
      result = #E_OUTOFMEMORY
    EndIf
  Else
    result = #E_INVALIDARG
  EndIf
  ;Set any error code. iDispatch errors will alreay have been set.
    CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
      COMateClass_INTERNAL_SetError(*this, result)
    CompilerEndIf
  ProcedureReturn *retVar
EndProcedure
;=================================================================================


;=================================================================================
;The following method calls a dispinterface method where no return value is required.
;Returns a HRESULT value. #S_OK for no errors.
;Errors reported by the methods called by the user will be reported elsewhere (eventually!)
Procedure.i COMateClass_SetProperty(*this._membersCOMateClass, command$, *hStatement=0)
  Protected result.i = #S_OK
  If command$ Or *hStatement
    result = COMateClass_INTERNAL_InvokePlus(*this, #DISPATCH_PROPERTYPUT, #VT_EMPTY, 0, command$, *hStatement)
  Else
    result = #E_INVALIDARG
  EndIf
  ;Set any error code. iDispatch errors will alreay have been set.
    If result = -1
      result = #S_FALSE
    Else
      CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
        COMateClass_INTERNAL_SetError(*this, result)
      CompilerEndIf
    EndIf
  ProcedureReturn result
EndProcedure
;=================================================================================


;=================================================================================
;The following function calls a dispinterface method where no return value is required.
;Returns a HRESULT value. #S_OK for no errors.
;Errors reported by the methods called by the user will be reported elsewhere (eventually!)
Procedure.i COMateClass_SetPropertyRef(*this._membersCOMateClass, command$, *hStatement=0)
  Protected result.i = #S_OK
  If command$ Or *hStatement
    result = COMateClass_INTERNAL_InvokePlus(*this, #DISPATCH_PROPERTYPUTREF, #VT_EMPTY, 0, command$, *hStatement)
  Else
    result = #E_INVALIDARG
  EndIf
  ;Set any error code. iDispatch errors will alreay have been set.
    If result = -1
      result = #S_FALSE
    Else
      CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
        COMateClass_INTERNAL_SetError(*this, result)
      CompilerEndIf
    EndIf
  ProcedureReturn result
EndProcedure
;=================================================================================


CompilerIf Defined(COMATE_NOINCLUDEATL, #PB_Constant)=0

;-COMate CLASS - EVENT RELATED METHODS.
;---------------------------------------------------------------------------------

;=================================================================================
;The following method attaches an event handler from the user's program to the underlying COM object. (Code based on that written by Freak.)
;Set callback to zero to remove any existing callback.
;Returns a HRESULT value. #S_OK for no errors.
Procedure.i COMateClass_SetEventHandler(*this._membersCOMateClass, eventName$, callback, returnType = #COMate_NORETURN, *riid.IID=0)
  Protected result.i = #S_OK
  Protected container.IConnectionPointContainer, enum.IEnumConnectionPoints, connection.IConnectionPoint, connIID.IID
  Protected dispTypeInfo.ITypeInfo, typeLib.ITypeLib, typeInfo.ITypeInfo
  Protected infoCount, index
  Protected *sink._COMateEventSink, newSink.IDispatch 
  If eventName$ = #COMate_CatchAllEvents Or *this\hWnd
    If returnType < #COMate_NoReturn Or returnType > #COMate_OtherReturn
      returnType = #COMate_NoReturn
    EndIf
    If eventName$ = #COMate_CatchAllEvents
      If returnType <> #COMate_NORETURN
        returnType = #COMate_OtherReturn  ;No sense in an explicit return value when dealing with any event!
      EndIf
      ;If their already exists a sink for this object then we just switch the main callback.
        If *this\eventSink And callback
          *this\eventSink\callback = callback
          *this\eventSink\returnType = returnType
          CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
            COMateClass_INTERNAL_SetError(*this, result)
          CompilerEndIf
          ProcedureReturn result ;No error.
        ElseIf *this\eventSink = 0 And callback = 0 ;No point proceeding with this.
          CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
            COMateClass_INTERNAL_SetError(*this, result)
          CompilerEndIf
          ProcedureReturn result ;No error reported.
        EndIf
    ElseIf *this\eventSink
      If callback And *this\hWnd
        SetProp_(*this\hWnd, eventName$+"_COMate", callback)
        SetProp_(*this\hWnd, eventName$+"_RETURN_COMate", returnType)
      ElseIf *this\hWnd
        RemoveProp_(*this\hWnd, eventName$+"_COMate")
        RemoveProp_(*this\hWnd, eventName$+"_RETURN_COMate")
      EndIf
        CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
          COMateClass_INTERNAL_SetError(*this, result)
        CompilerEndIf
      ProcedureReturn result
    ElseIf callback = 0 ;*this\eventSink will equal 0 as well.
      CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
        COMateClass_INTERNAL_SetError(*this, result)
      CompilerEndIf
      ProcedureReturn result
    EndIf
    ;Only remaining options are for wishing to remove a previously installed sink (requires #COMate_CatchAllEvents) or a completely new sink is to be installed.
    result = *this\iDisp\GetTypeInfoCount(@infoCount)
    If SUCCEEDED(result)  
      If InfoCount = 1
        result = *this\iDisp\GetTypeInfo(0, 0, @dispTypeInfo)
        If SUCCEEDED(result)
          result = dispTypeInfo\GetContainingTypeLib(@typeLib, @index)
          If SUCCEEDED(result)
            result = *this\iDisp\QueryInterface(?IID_IConnectionPointContainer, @container)
            If SUCCEEDED(result)
              If *riid.IID = 0
                result = container\EnumConnectionPoints(@enum.IEnumConnectionPoints)
                If SUCCEEDED(result)
                  enum\Reset()
                  result = enum\Next(1, @connection, #Null)
                  While result = #S_OK
                    result = Connection\GetConnectionInterface(@connIID)
                    If SUCCEEDED(result)  ;We have a valid IID for the outgoing interface managed by this connection point.
                      result = typeLib\GetTypeInfoOfGuid(connIID, @typeInfo)
                      If SUCCEEDED(result)
                        enum\Release()
                        Goto COMateClass_SetEventHandler_L1
                      EndIf
                    EndIf
                    connection\Release()
                    result = enum\Next(1, @connection, #Null)
                  Wend
                  enum\Release()
                EndIf
              Else ;The user has specified a connection point interface.
                result = container\FindConnectionPoint(*riid, @connection)
                If SUCCEEDED(result)
                  result = Connection\GetConnectionInterface(@connIID) ;May or may not equal the IID pointed to by *riid.
                  If SUCCEEDED(result)  ;We have a valid IID for the outgoing interface managed by this connection point.
                    result = typeLib\GetTypeInfoOfGuid(*riid, @typeInfo)
                    If SUCCEEDED(result)
COMateClass_SetEventHandler_L1:
                      If eventName$ = #COMate_CatchAllEvents And callback = 0 ;Remove existing sink.
                        ;The Unadvise() method will call Release() on our sink and so we leave all tidying up to this method.
                          connection\Unadvise(*this\eventSink\cookie)
                          TypeInfo\Release()
                      Else ;New sink needs creating.
                        *sink = AllocateMemory(SizeOf(_COMateEventSink))
                        If *sink
                          *this\eventSink = *sink
                          With *this\eventSink
                            \Vtbl = ?VTable_COMateEventSink
                            \refCount = 1
                            \typeInfo = typeInfo
                            If eventName$ = #COMate_CatchAllEvents
                              \callback = Callback
                              \returnType = returnType
                            ElseIf *this\hWnd
                              SetProp_(*this\hWnd, eventName$+"_COMate", callback)
                              SetProp_(*this\hWnd, eventName$+"_RETURN_COMate", returnType)
                            EndIf
                            CopyMemory(connIID, @\connIID, SizeOf(IID))
                            \parent = *this
                          EndWith
                          newSink = *sink
                          result = connection\Advise(newSink, @*this\eventSink\cookie)  ;Calls QueryInterface() on NewSink hence the subsequent Release().
                          ;In the case of an error this release will decrement the ref counter to zero and then tidy up!
                            NewSink\Release()
                        Else
                          TypeInfo\Release()
                          result = #E_OUTOFMEMORY
                        EndIf
                      EndIf
                    EndIf                
                  EndIf
                  connection\Release()
                EndIf         
              EndIf
              container\Release()
            EndIf
            typeLib\Release()
          EndIf
          dispTypeInfo\Release()
        EndIf
      EndIf
    EndIf
  Else
    result = #E_FAIL
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, result)
  CompilerEndIf
  ProcedureReturn result
EndProcedure
;=================================================================================


;=================================================================================
;The following method, valid only when called from a user's event procedure, retrieves the specified parameter in the form
;of a quad.
Procedure.q COMateClass_GetIntegerEventParam(*this._membersCOMateClass, index)
  Protected result.i = #S_OK, var.VARIANT, puArgErr
  If *this\eventSink And *this\eventSink\dispParams
    If index > 0 And index <= *this\eventSink\dispParams\cArgs+*this\eventSink\dispParams\cNamedArgs
      result = DispGetParam_(*this\eventSink\dispParams, index-1, #VT_I8, var, @puArgErr)
    Else
      result = #E_INVALIDARG
    EndIf
  Else
    result = #E_FAIL
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, result)
  CompilerEndIf
  ProcedureReturn var\llval
EndProcedure
;=================================================================================


;=================================================================================
;The following method, valid only when called from a user's event procedure, retrieves the specified parameter in the form
;of a COM interface. It does not wrap any returned object into a COMate object. Returns zero if an error.
;The user MUST call Release() on any object returned.
;Leave objectType = #VT_DISPATCH to have an iDispatch interface returned. Any other value will result in an iUnknown interface.
Procedure.i COMateClass_GetObjectEventParam(*this._membersCOMateClass, index, objectType = #VT_DISPATCH)
  Protected result.i = #S_OK, var.VARIANT, puArgErr
  If *this\eventSink And *this\eventSink\dispParams
    If index > 0 And index <= *this\eventSink\dispParams\cArgs+*this\eventSink\dispParams\cNamedArgs
      If objectType <> #VT_DISPATCH
        objectType = #VT_UNKNOWN
      EndIf
      result = DispGetParam_(*this\eventSink\dispParams, index-1, objectType, @var, @puArgErr)
    Else
      result = #E_INVALIDARG
    EndIf
  Else
    result = #E_FAIL
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, result)
  CompilerEndIf
  ProcedureReturn var\pDispVal
EndProcedure
;=================================================================================


;=================================================================================
;The following method, valid only when called from a user's event procedure, retrieves the specified parameter in the form
;of a double.
Procedure.d COMateClass_GetRealEventParam(*this._membersCOMateClass, index)
  Protected result.i = #S_OK, var.VARIANT, puArgErr
  If *this\eventSink And *this\eventSink\dispParams
    If index > 0 And index <= *this\eventSink\dispParams\cArgs+*this\eventSink\dispParams\cNamedArgs
      result = DispGetParam_(*this\eventSink\dispParams, index-1, #VT_R8, var, @puArgErr)
    Else
      result = #E_INVALIDARG
    EndIf
  Else
    result = #E_FAIL
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, result)
  CompilerEndIf
  ProcedureReturn var\dblVal
EndProcedure
;=================================================================================


;=================================================================================
;The following method, valid only when called from a user's event procedure, retrieves the specified parameter in the form
;of a string.
Procedure.s COMateClass_GetStringEventParam(*this._membersCOMateClass, index)
  Protected result.i = #S_OK, var.VARIANT, puArgErr, text$
  If *this\eventSink And *this\eventSink\dispParams
    If index > 0 And index <= *this\eventSink\dispParams\cArgs+*this\eventSink\dispParams\cNamedArgs
      result = DispGetParam_(*this\eventSink\dispParams, index-1, #VT_BSTR, var, @puArgErr)
      If var\bstrVal
        text$ = PeekS(var\bstrVal, -1, #PB_Unicode)
        SysFreeString_(var\bstrVal)
      EndIf
    Else
      result = #E_INVALIDARG
    EndIf
  Else
    result = #E_FAIL
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, result)
  CompilerEndIf
  ProcedureReturn text$
EndProcedure
;=================================================================================


;=================================================================================
;The following method, valid only when called from a user's event procedure, returns 0 if the specified parameter was
;not passed by reference.
;Otherwise, it returns the variant #VT_... type of the underlying parameter and it places the address of the underlying
;parameter into the *ptrParameter parameter (if non-zero). This allows the client application to alter the value of the parameter
;as appropriate.
;For even more flexibility, you can obtain a pointer to the actual variant containing the parameter as supplied by the ActiveX control.
Procedure.i COMateClass_IsEventParamPassedByRef(*this._membersCOMateClass, index, *ptrParameter.INTEGER=0, *ptrVariant.INTEGER=0)
  Protected result.i = #S_OK, *var.VARIANT, *ptr.INTEGER, numArgs
  If *this\eventSink And *this\eventSink\dispParams
    numArgs = *this\eventSink\dispParams\cArgs+*this\eventSink\dispParams\cNamedArgs
    If index > 0 And index <= numArgs
      *var = *this\eventSink\dispParams\rgvarg + (numArgs-index)*SizeOf(VARIANT)
      If *var\vt&#VT_BYREF
        result = *var\vt&~#VT_BYREF
        If *ptrParameter
          *ptrParameter\i = *var\pllval
        EndIf
        If *ptrVariant
          *ptrVariant\i = *var
        EndIf
      EndIf
    Else
      CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
        COMateClass_INTERNAL_SetError(*this, #E_INVALIDARG)
      CompilerEndIf
    EndIf
  Else
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(*this, #E_FAIL)
  CompilerEndIf
  EndIf
  ProcedureReturn result
EndProcedure
;=================================================================================

CompilerEndIf


;/////////////////////////////////////////////////////////////////////////////////
;The following function is called by the COMatePLUS_CompileSubobjectInvokation() function when extracting method arguments and only when the final
;option is a numeric argument.
;Returns #True if a valid variant numeric type is found and also places the relevant value within the given variant.
Procedure.i COMate_INTERNAL_CheckNumeric(arg$, *var.VARIANT)
  Protected result.i = #True, i, blnPoint, length, *ptr.CHARACTER
  Protected val.q, byte.b, word.w, long.l
  length = Len(arg$)
  *ptr = @arg$
  For i = 1 To length
    If *ptr\c = '-' Or *ptr\c = '+'
      If i > 1
        result = 0
        Break
      EndIf
    ElseIf *ptr\c = '.' And blnPoint = #False
      blnPoint = #True
    ElseIf *ptr\c < '0' Or *ptr\c > '9'
      result = 0
      Break
    EndIf    
    *ptr+SizeOf(CHARACTER)
  Next
  If result
    If blnPoint ;Decimal.
      *var\vt = #VT_R8
      *var\dblVal = ValD(arg$)
    Else ;Some kind of integer.
      val = Val(arg$)
      If val = 0 ;Shove this into a signed 'long'.
        *var\vt = #VT_I4
        *var\lVal = 0
      Else
        ;Check if the value will fit into a signed-byte or a signed-word or a signed-long or a signed-quad.
          byte = val
          If byte = val ;Signed byte.
            *var\vt = #VT_I1
            *var\cVal = val
          Else
            word = val
            If word = val ;Signed word.
              *var\vt = #VT_I2
              *var\iVal = val
            Else
              long = val
              If long = val ;Signed long.
                *var\vt = #VT_I4
                *var\lVal = val
              Else ;Quad.
                *var\vt = #VT_I8
                *var\llVal = val
              EndIf
            EndIf
          EndIf
      EndIf
    EndIf
  EndIf
  ProcedureReturn result
EndProcedure
;/////////////////////////////////////////////////////////////////////////////////


;/////////////////////////////////////////////////////////////////////////////////
;The following function is called by the COMatePLUS_CompileSubobjectInvokation() function when extracting method arguments and only for non-empty strings.
;Quoted strings (beginning with ') can contain escaped sequences of the form $xxxx where xxxx represent a hex number.
;Together $xxxx represents a single character code; e.g. an Ascii code. E.g. $0024 would be replaced by a $ character and
;$0027 would be replaced by a ' character.
;Adjusts the string in-place.
Procedure COMate_INTERNAL_EscapeString(ptrText)
  Protected *source.CHARACTER, *destination.CHARACTER, blnEscape, value, t1, pow, i
  *source.CHARACTER = ptrText
  *destination = *source
  While *source\c
    If *source\c = 36
      ;Is this the beginning of an escape sequence?
        blnEscape = #True
        t1 = *source
        value = 0
        pow = 4096 ;16^3.
        For i = 1 To 4
          *source + SizeOf(CHARACTER)
          If *source\c = 0 ;Null terminator.
            Break 2
          ElseIf *source\c >= '0' And *source\c <= '9'
            value + (*source\c-'0')*pow
          ElseIf *source\c >= 'A' And *source\c <= 'F'
            value + (*source\c-'A'+10)*pow
          ElseIf *source\c >= 'a' And *source\c <= 'f'
            value + (*source\c-'a'+10)*pow
          Else
            blnEscape = #False
            Break
          EndIf
          pow>>4
        Next
        If blnEscape ;We have an escape sequence.
          *destination\c = value&$ff : *destination + SizeOf(CHARACTER)
          *source + SizeOf(CHARACTER)
        Else
          *source = t1
          Goto COMate_labelEscape1          
        EndIf        
    Else
COMate_labelEscape1:
      *destination\c = *source\c
      *destination + SizeOf(CHARACTER)
      *source + SizeOf(CHARACTER)
    EndIf
  Wend
  *destination\c = 0 ;Null termminator.
EndProcedure
;/////////////////////////////////////////////////////////////////////////////////


;/////////////////////////////////////////////////////////////////////////////////
;The following function is called (possibly more than once) by the COMateClass_INTERNAL_InvokePlus as we drill down through
;subobject method calls etc. This performs the task of calling the dispinterface methods.
;Returns a HRESULT value; #S_OK for no errors.
Procedure.i COMateClass_INTERNAL_InvokeiDispatch(*this._membersCOMateClass, invokeType, returnType, *ret.VARIANT, iDisp.iDispatch, subObjectIndex, *statement._COMatePLUSStatement)
  Protected result.i = #S_OK
  Protected dispID, dp.DISPPARAMS, dispIDNamed, excep.EXCEPINFO2, uiArgErr
  ;First task is to retrieve the dispID corresponding to the method/property.
    result = iDisp\GetIDsOfNames(?IID_NULL, @*statement\methodName[subObjectIndex], 1, #LOCALE_USER_DEFAULT, @dispID)
  If SUCCEEDED(result)
    ;Now prepare to call the method/property.
      dispidNamed  = #DISPID_PROPERTYPUT
      If *statement\numArgs[subObjectIndex]
        dp\rgvarg = *statement\ptrVarArgs[subObjectIndex] + (#COMate_MAXNUMVARIANTARGS - *statement\numArgs[subObjectIndex])*SizeOf(VARIANT)
      EndIf
      dp\cargs = *statement\numArgs[subObjectIndex]
      If invokeType & (#DISPATCH_PROPERTYPUT | #DISPATCH_PROPERTYPUTREF)
        dp\cNamedArgs = 1
        dp\rgdispidNamedArgs = @dispidNamed
      EndIf
    ;Call the method/property.
      result = iDisp\Invoke(dispID, ?IID_NULL, #LOCALE_USER_DEFAULT, invokeType, dp, *ret, excep, @uiArgErr)
    If result = #DISP_E_EXCEPTION
      ;Has the automation server deferred from filling in the EXCEPINFO2 structure?
        If excep\pfnDeferredFillIn
          excep\pfnDeferredFillIn(excep)
        EndIf
        If excep\bstrSource
          SysFreeString_(excep\bstrSource)
        EndIf          
        If excep\bstrDescription
          CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
            COMateClass_INTERNAL_SetError(*this, result,  #True, PeekS(excep\bstrDescription, -1, #PB_Unicode))
          CompilerEndIf
          SysFreeString_(excep\bstrDescription)
        Else
          CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
            COMateClass_INTERNAL_SetError(*this, result,  #True)
          CompilerEndIf
        EndIf          
        If excep\bstrHelpFile
          SysFreeString_(excep\bstrHelpFile)
        EndIf          
    EndIf
  EndIf
  ProcedureReturn result
EndProcedure
;/////////////////////////////////////////////////////////////////////////////////


;/////////////////////////////////////////////////////////////////////////////////
;The following function is called by all methods which need to invoke a COM method through iDispatch etc.
;It drills down through all the sub-objects of a method call as appropriate.
;Returns a HRESULT value; #S_OK for no errors.
Procedure.i COMateClass_INTERNAL_InvokePlus(*this._membersCOMateClass, invokeType, returnType, *ret.VARIANT, command$, *hStatement._COMatePLUSStatement)
  Protected result.i = #S_OK, *statement._COMatePLUSStatement, subObjectIndex
  Protected iDisp.iDispatch, var.VARIANT
  ;First job is to prepare a statement if one has not been provided by the developer.
    If *hStatement
      *statement = *hStatement
    Else
      result = COMate_INTERNAL_PrepareStatement(command$, @*statement)
    EndIf
  If *statement
    VariantInit_(var)
    iDisp = *this\iDisp
    iDisp\AddRef()  ;This seemingly extraneous AddRef() will be balanced (released) in the following loop or the code following the loop.
    For subObjectIndex = 1 To *statement\numSubObjects-1
      result = COMateClass_INTERNAL_InvokeiDispatch(*this, #DISPATCH_METHOD|#DISPATCH_PROPERTYGET, #VT_DISPATCH, var, iDisp, subObjectIndex, *statement)
      iDisp\Release()
      iDisp = var\pdispVal
      If FAILED(result) Or iDisp = 0
        Break
      EndIf
      VariantInit_(var)
    Next
    If SUCCEEDED(result)
      If iDisp
        result = COMateClass_INTERNAL_InvokeiDispatch(*this, invokeType, returnType, *ret, iDisp, *statement\numSubObjects, *statement)
        iDisp\Release()
      Else
        ;In this case we set an error with extra info.
          CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
            COMateClass_INTERNAL_SetError(*this, #S_FALSE, 0, "The '" + PeekS(*statement\methodName[subObjectIndex], -1, #PB_Unicode)  + "' property returned a NULL object!")
          CompilerEndIf
        result = -1 ;This will ensure that the COMateClass_INTERNAL_SetError() does not reset the error.
      EndIf
      If SUCCEEDED(result)
        ;Sort out any return.
          If *ret And *ret\vt <> returnType And returnType <> #VT_EMPTY
            result = VariantChangeType_(*ret, *ret, 16, returnType)
          EndIf
      EndIf
    EndIf
    ;Tidy up.
      If *hStatement = 0
        COMate_INTERNAL_FreeStatementHandle(*statement)
      EndIf
  EndIf
  ProcedureReturn result
EndProcedure
;/////////////////////////////////////////////////////////////////////////////////


;///////////////////////////////////////////////////////////////////////////////////////////
;iDispatch errors will already have been processed.
Procedure COMateClass_INTERNAL_SetError(*this._membersCOMateClass, result, blnAllowDispError = 0, dispError$="")
  Protected *error._COMateThreadErrors, Array.i, winError, len, *buffer
  If COMate_gErrorTLS = 0 Or COMate_gErrorTLS = -1
    ;Create a new TLS index to hold error information.
      COMate_gErrorTLS = TlsAlloc_()
  EndIf
  If COMate_gErrorTLS = -1 Or result = -1 Or (result = #DISP_E_EXCEPTION And blnAllowDispError = 0)
    ProcedureReturn
  EndIf
  ;Is there a TLS entry for this thread.
    *error = TlsGetValue_(COMate_gErrorTLS)
    If *error = 0 ;No existing entry.
      ;Attempt to allocate memory for a TLS entry for this thread.
        *error = AllocateMemory(SizeOf(_COMateThreadErrors))
      If *error
        If TlsSetValue_(COMate_gErrorTLS, *error)
          ;Need to extend the memory if already allocated for the *COMate_gPtrThreadArray array so that the error memory can be freed later on.
            Array = ReAllocateMemory(COMate_gPtrThreadArray, (COMate_gNumThreadElements+1)*SizeOf(Array))
            If Array
              COMate_gPtrThreadArray = Array
              PokeI(COMate_gPtrThreadArray + COMate_gNumThreadElements*SizeOf(Array), *error)
              COMate_gNumThreadElements+1
            Else
              TlsSetValue_(COMate_gErrorTLS, 0)
              FreeMemory(*error)          
              *error = 0
            EndIf
        Else
          FreeMemory(*error)          
          *error = 0
        EndIf
      EndIf
    EndIf
  If *error
    *error\lastErrorCode = result
    Select result
      Case #S_OK
        *error\lastError$ = "Okay."
      Case #S_FALSE
        If dispError$
          *error\lastError$ = "The operation completed, but was only partially successful. (" + dispError$ + ")"
        Else
          *error\lastError$ = "The operation completed, but was only partially successful."
        EndIf
      Case #E_FAIL
        *error\lastError$ = "Unspecified error."
      Case #E_INVALIDARG
        *error\lastError$ = "One or more arguments are invalid. Possibly a numerical overflow or too many nested objects, -if so, try splitting your method call into two or more subcalls."
      Case #E_NOINTERFACE
        *error\lastError$ = "Method is not implemented."
      Case #E_OUTOFMEMORY
        *error\lastError$ = "Problem allocating memory." + #CRLF$ + #CRLF$ + "(Possibly too many method arguments. Each method/property is limited by COMatePLUS to a maximum of " + Str(#COMate_MAXNUMVARIANTARGS) + " arguments.)"
      Case #E_UNEXPECTED
        *error\lastError$ = "An unexpected error."
      Case #E_POINTER
        *error\lastError$ = "An invalid pointer was supplied."
      Case #E_NOTIMPL
        *error\lastError$ = "Not implemented. In the case of attaching an event handler to a COM object, this could signify that the object does not provide any type information."
      Case #CO_E_CLASSSTRING, #REGDB_E_CLASSNOTREG
        *error\lastError$ = "Invalid progID/CLSID. Check your spelling of the programmatic identifier. Also check that the component / ActiveX control has been registered."
      Case #CO_E_SERVER_EXEC_FAILURE
        *error\lastError$ = "Server execution failed. Usually caused by an 'out of process server' timing out when asked to create an instance of a 'class factory'."
      Case #DISP_E_TYPEMISMATCH
        *error\lastError$ = "Type mismatch in the method parameters."
      Case #TYPE_E_ELEMENTNOTFOUND
        *error\lastError$ = "No type description was found in the library with the specified GUID whilst trying to create an event handler."
      Case #CONNECT_E_ADVISELIMIT
        *error\lastError$ = "Unable to set event handler because the connection point has already reached its limit of connections."
      Case #CLASS_E_NOAGGREGATION
        *error\lastError$ = "Class does not support aggregation (or class object is remote)."
      Case #DISP_E_OVERFLOW
        *error\lastError$ = "Overflow error whilst converting between types."
      Case #DISP_E_UNKNOWNNAME
        *error\lastError$ = "Method/property not supported by this object."
      Case #DISP_E_BADPARAMCOUNT
        *error\lastError$ = "Invalid number of method/property parameters."
      Case #DISP_E_BADVARTYPE
        *error\lastError$ = "A method/property parameter is not a valid (variant) type."
      Case #DISP_E_MEMBERNOTFOUND
        *error\lastError$ = "Member not found. (Check that you have not omitted any optional parameters and are not trying to set a read-only property etc.)"
      Case #DISP_E_NOTACOLLECTION
        *error\lastError$ = "Does not support a collection."
      Case #E_ACCESSDENIED
        *error\lastError$ = "A 'general' access denied error."
      Case #RPC_E_WRONG_THREAD
        *error\lastError$ = "The application called upon an interface that was marshalled for a different thread."
      Case #DISP_E_EXCEPTION
        *error\lastError$ = dispError$
        If *error\lastError$ = ""
          *error\lastError$ = "An exception occurred during the execution of this method/property."
        EndIf
      Default 
        ;Check for a WIN32 facility code.
          If *error\lastErrorCode & $7FFF0000 = $70000
            winError = *error\lastErrorCode&$FFFF
            len = FormatMessage_(#FORMAT_MESSAGE_ALLOCATE_BUFFER|#FORMAT_MESSAGE_FROM_SYSTEM, 0, winError, 0, @*Buffer, 0, 0) 
            If len 
              *error\lastError$ = "(FACILITY_WIN32 error " + Str(winError) + ") " + PeekS(*Buffer, len) 
              LocalFree_(*Buffer) 
            Else
              *error\lastError$ = "(FACILITY_WIN32 error " + Str(winError) + ") Unable to retrieve error description from system!"
            EndIf
          Else
            *error\lastError$ = "Unknown error. (Code : Hex " + Hex(*error\lastErrorCode, #PB_Long) + "). Please report this error code to the author at 'enquiries@nxsoftware.com'"
          EndIf
    EndSelect
  EndIf
EndProcedure
;///////////////////////////////////////////////////////////////////////////////////////////


CompilerIf Defined(COMATE_NOINCLUDEATL, #PB_Constant)=0
;-OUTGOING 'SINK' INTERFACE METHODS.
;------------------------------------------------------------------------

;=================================================================================
;The QueryInterface() method of our COMate sink objects.
Procedure.i COMateSinkClass_QueryInterface(*this._COMateEventSink, *IID.IID, *Object.INTEGER) 
  If CompareMemory(*IID, ?IID_IUnknown, SizeOf(IID)) Or CompareMemory(*IID, ?IID_IDispatch, SizeOf(IID)) Or CompareMemory(*IID, @*this\connIID, SizeOf(IID))
    *Object\i = *this
    *this\refCount + 1
    ProcedureReturn #S_OK
  Else
    *Object\i = 0
    ProcedureReturn #E_NOINTERFACE
  EndIf
EndProcedure
;=================================================================================


;=================================================================================
;The AddRef() method of our COMate sink objects.
Procedure.i COMateSinkClass_AddRef(*this._COMateEventSink)
  *this\refCount + 1
  ProcedureReturn *this\refCount
EndProcedure
;=================================================================================


;=================================================================================
;The Release() method of our COMate sink objects.
Procedure.i COMateSinkClass_Release(*this._COMateEventSink)
  *this\refCount - 1
  If *this\refCount = 0
    If *this\parent
      ;Release all event related window properties added to the ActiveX container.
        If IsWindow_(*this\parent\hWnd)
          EnumPropsEx_(*this\parent\hWnd, @COMate_DelSinkPropsCallback(),#Null)
        EndIf
      *this\parent\eventSink = 0
    EndIf
    *this\typeInfo\Release()
    FreeMemory(*this)
    ProcedureReturn 0
  Else 
    ProcedureReturn *this\refCount
  EndIf
EndProcedure
;=================================================================================


;=================================================================================
;The next 3 methods of the COMate sink interface are possibly not required, but ...
Procedure.i COMateSinkClass_GetTypeInfoCount(*this._COMateEventSink, *pctinfo.INTEGER)
  *pctinfo\i = 1
  ProcedureReturn #S_OK
EndProcedure

Procedure.i COMateSinkClass_GetTypeInfo(*this._COMateEventSink, iTInfo, lcid, *ppTInfo.INTEGER)
  *ppTInfo\i = *this\typeInfo
  *this\typeInfo\AddRef()
  ProcedureReturn #S_OK
EndProcedure

Procedure.i COMateSinkClass_GetIDsOfNames(*this._COMateEventSink, *riid, *rgszNames, *cNames, lcid, *DispID)
  ProcedureReturn DispGetIDsOfNames_(*this\typeInfo, *rgszNames, *cNames, *DispID)
EndProcedure
;=================================================================================


;=================================================================================
;The Invoke() method of our COMate sink objects.
;This is where we call the user's event procedure.
Procedure.i COMateSinkClass_Invoke(*this._COMateEventSink, dispid, *riid, lcid, wflags.w, *Params.DISPPARAMS, *Result.VARIANT, *pExept, *ArgErr)
    Protected result.i = #S_OK, bstrName.i, nameCount, tempParams, eventName$, returnType, address
    Protected callbackNoReturn.COMate_EventCallback_NORETURN, callbackIntegerReturn.COMate_EventCallback_INTEGERRETURN, callbackRealReturn.COMate_EventCallback_REALRETURN, callbackStringReturn.COMate_EventCallback_STRINGRETURN, callbackUnknownReturn.COMate_EventCallback_UNKNOWNRETURN
    Protected intRet.q, realRet.d, stringRet$
    result = *this\TypeInfo\GetNames(dispid, @bstrName, 1, @nameCount)
    If SUCCEEDED(result)
      If bstrName
        tempParams = *this\dispParams
        *this\dispParams = *Params
        eventName$ = PeekS(bstrName, -1, #PB_Unicode)
        SysFreeString_(bstrName)
        ;Call the 'global' #COMate_CatchAllEvents handler if defined.
        If *this\callback
            If *this\returnType = #COMate_OtherReturn
              callbackUnknownReturn = *this\callback
              callbackUnknownReturn(*this\parent, eventName$, *Params\cArgs + *Params\cNamedArgs, *Result)
            Else
              *this\callback(*this\parent, eventName$, *Params\cArgs + *Params\cNamedArgs)
            EndIf
          EndIf
          ;Call any individual handler attached to this event. We need to take into account the return type.
          If *this\parent\hWnd
            address = GetProp_(*this\parent\hWnd, eventName$ + "_COMate")
            If address
              returnType = GetProp_(*this\parent\hWnd, eventName$ + "_RETURN_COMate")
              Select returnType
                Case #COMate_NoReturn
                  callbackNoReturn = address
                  callbackNoReturn(*this\parent, eventName$, *Params\cArgs + *Params\cNamedArgs)
                Case #COMate_IntegerReturn
                  callbackIntegerReturn = address
                  intRet = callbackIntegerReturn(*this\parent, eventName$, *Params\cArgs + *Params\cNamedArgs)
                  If *Result
                    *Result\vt = #VT_I8
                    *Result\llVal = intRet
                  EndIf
                Case #COMate_RealReturn
                  callbackRealReturn = address
                  realRet = callbackRealReturn(*this\parent, eventName$, *Params\cArgs + *Params\cNamedArgs)
                  If *Result
                    *Result\vt = #VT_R8
                    *Result\dblVal = realRet
                  EndIf
                Case #COMate_StringReturn
                  callbackStringReturn = address
                  stringRet$ = callbackStringReturn(*this\parent, eventName$, *Params\cArgs + *Params\cNamedArgs)
                  If *Result
                    *Result\vt = #VT_BSTR
                    *Result\bstrVal = COMate_MakeBSTR(stringRet$)
                  EndIf
                Case #COMate_OtherReturn
                  callbackUnknownReturn = address
                  callbackUnknownReturn(*this\parent, eventName$, *Params\cArgs + *Params\cNamedArgs, *Result)
              EndSelect
            EndIf
          EndIf
        *this\dispParams = tempParams
      Else
        result = #E_OUTOFMEMORY
      EndIf
    EndIf
  ProcedureReturn result
EndProcedure
;=================================================================================


;=================================================================================
;The following callback function is called by windows as a result of the EnumPropsEx_() function
;issued when an outgoing sink object is destroyed.
;We use this to delete the properties we have created.
Procedure.i COMate_DelSinkPropsCallback(hWnd, lpszString, hData,dwData)
  Protected text$
  If lpszString>>16<>0 ;Confirms that this parameter points to a string and is not merely an atom.
    text$ = PeekS(lpszString)
    If Right(PeekS(lpszString),7)="_COMate"
      RemoveProp_(hWnd, lpszString)
    EndIf
  EndIf
ProcedureReturn 1  
EndProcedure
;=================================================================================

CompilerEndIf


;-STATEMENT FUNCTIONS.
;----------------------------------------------

;=================================================================================
;The following function compiles the given command string and if successful, returns a statement handle.
;Returns zero otherwise.
Procedure.i COMate_PrepareStatement(command$)
  Protected errorCode = #S_OK, *hStatement._COMatePLUSStatement
  errorCode = COMate_INTERNAL_PrepareStatement(command$, @*hStatement)
  ;Set any error code.
    CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
      COMateClass_INTERNAL_SetError(0, errorCode)
    CompilerEndIf
  ProcedureReturn *hStatement
EndProcedure
;=================================================================================


;=================================================================================
;Returns, if successful, a direct pointer to the appropriate variant structure. This address will not change for the life of the statement
;and thus need only be retrieved once.
;Index is 1-based.
Procedure.i COMate_GetStatementParameter(*hStatement._COMatePLUSStatement, index)
  Protected errorCode = #E_INVALIDARG, result, i, total
  If index > 0
    ;Track down which sub-object
      For i = 1 To *hStatement\numSubObjects
        total + *hStatement\numArgs[i]
        If index <= total
          ;Adjust the index to reflect the underlying sub-object's number of parameters.
            index = *hStatement\numArgs[i] - total + index
          ;Locate the relevant variant argument.
            result= *hStatement\ptrVarArgs[i] + (#COMate_MAXNUMVARIANTARGS - index)*SizeOf(VARIANT)
          errorCode = #S_OK
          Break
        EndIf
      Next
  EndIf
  ;Set any error code.
    CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
      COMateClass_INTERNAL_SetError(0, errorCode)
    CompilerEndIf
  ProcedureReturn result
EndProcedure
;=================================================================================


;=================================================================================
;The following function frees the specified statement.
Procedure COMate_FreeStatementHandle(*hStatement._COMatePLUSStatement)
  COMate_INTERNAL_FreeStatementHandle(*hStatement)
  ;Set any error code.
    CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
      COMateClass_INTERNAL_SetError(0, #S_OK)
    CompilerEndIf
EndProcedure
;=================================================================================


;-INTERNAL FUNCTIONS.
;------------------------------------------

;=================================================================================
;The following function frees the specified statement but does not set any error.
Procedure COMate_INTERNAL_FreeStatementHandle(*hStatement._COMatePLUSStatement)
  Protected i, j, *varArg.VARIANT
  For i = 1 To *hStatement\numSubObjects
    ;First free any method BSTR.
      If *hStatement\methodName[i]
        SysFreeString_(*hStatement\methodName[i])
      EndIf
    ;Now the variant array.
      If *hStatement\ptrVarArgs[i]
        *varArg = *hStatement\ptrVarArgs[i] + (#COMate_MAXNUMVARIANTARGS - 1) * SizeOf(VARIANT)
        For j = 1 To *hStatement\numArgs[i]
          VariantClear_(*varArg)
          *varArg - SizeOf(VARIANT)
        Next
        FreeMemory(*hStatement\ptrVarArgs[i])
      EndIf
  Next
  FreeMemory(*hStatement)
EndProcedure
;=================================================================================


;=================================================================================
;The following internal function compiles the given command string and if successful, places a statement handle into the buffer
;pointed to by *ptrStatement.
;Returns a HRESULT but does NOT set any error.
Procedure.i COMate_INTERNAL_PrepareStatement(command$, *ptrStatement.INTEGER)
  Protected errorCode = #S_OK, *hStatement._COMatePLUSStatement
  Protected Dim parse._COMateParse(#COMate_MAXNUMSUBOBJECTS), i, subObject
  If command$
    ;Allocate memory for a statement handle.
      *hStatement = AllocateMemory(SizeOf(_COMatePLUSStatement))
    If *hStatement
      ;Tokenise the command string.
        *hStatement\numSubObjects = COMatePLUS_TokeniseCommand(command$, "(),\'= ", parse())
        If *hStatement\numSubObjects
          For subObject = 1 To *hStatement\numSubObjects
            ;We need to parse/compile the tokenised command corresponding to each individual sub-object.
              errorCode = COMatePLUS_CompileSubobjectInvokation(*hStatement, subObject, parse())
              If errorCode <> #S_OK
                COMate_FreeStatementHandle(*hStatement)
                Break
              EndIf
          Next
          If errorCode = #S_OK
            *ptrStatement\i = *hStatement
          EndIf
        Else
          FreeMemory(*hStatement)
          errorCode = #E_INVALIDARG
        EndIf
    Else
      errorCode = #E_OUTOFMEMORY
    EndIf    
  Else
    errorCode = #E_INVALIDARG
  EndIf
  ProcedureReturn errorCode
EndProcedure
;=================================================================================


;///////////////////////////////////////////////////////////////////////////////////////////
;The following function tokenises the given command string.
;Submethod calls (+ associated parameters) are placed into the parse array; 1 element per subobject.
;This is very optimised by avoiding string functions as far as is possible; instead using multiple pointers etc.
;Returns zero if the line cannot be parsed else a count of the number of method calls.
;Error checking is included but is supplemented later on.
Procedure.i COMatePLUS_TokeniseCommand(command$, separator$, Array parse._COMateParse(1))
  Protected length, methodCount=1, numEquals, t1, i, lenSeparator
  Protected *command.CHARACTER, *buffer.CHARACTER, buffer, *ptrString.STRING, *ptrSeparator.CHARACTER, charPos = 1
  length=Len(command$)
  If length
    buffer = AllocateMemory((length+1)*SizeOf(CHARACTER))
    If buffer
      lenSeparator = Len(separator$)
      *ptrString = @buffer  ;Speedy (pointer) access to the contents in the form of a string.
      *buffer = buffer
      parse(methodCount)\numberoftokens=0
      *command = @command$
      Repeat 
        ;Search the separator string looking for this character.
          *ptrSeparator = @separator$
          t1 = #False
          For i = 0 To lenSeparator-1
            If *ptrSeparator\c = *command\c
              t1 = #True
              Break
            EndIf
            *ptrSeparator + SizeOf(CHARACTER)
          Next
        If t1
          If *buffer <> buffer
            parse(methodCount)\tokens$[parse(methodCount)\numberoftokens]=*ptrString\s
            parse(methodCount)\numberoftokens+1
            *buffer = buffer : *buffer\c = 0
          ElseIf *command\c = 39 ;Open quote, buffer empty.
            *buffer\c = *command\c : *buffer + SizeOf(CHARACTER) 
            ;Find closing quote.
              t1 = #False ;Boolean flag to indicate a closing quote.
              While charPos < length
                charPos + 1
                *command + SizeOf(character)
                *buffer\c = *command\c : *buffer + SizeOf(CHARACTER)
                If *command\c = 39
                  t1 = #True
                  *buffer\c = 0 ;Null.
                  Break
                EndIf
              Wend              
            If t1 = #False ;No closing quote.
              methodCount = 0
              Break
            EndIf
            parse(methodCount)\tokens$[parse(methodCount)\numberoftokens]=*ptrString\s
            parse(methodCount)\numberoftokens+1
            charPos+1 : *command + SizeOf(CHARACTER)
            *buffer = buffer : *buffer\c = 0
          ElseIf *command\c <> 32 ;Buffer empty.
            If *command\c = 40 ;"(".
              parse(methodCount)\numOpenBrackets + 1
            ElseIf *command\c = 41 ;")".
              If parse(methodCount)\numOpenBrackets
                parse(methodCount)\numCloseBrackets + 1
              Else
                methodCount = 0
                Break
              EndIf
            ElseIf *command\c = 61 ;"=", buffer empty.
              numEquals+1 ;Only allow 1 equals and then only for setting properties.
              If numEquals > 1 
                methodCount = 0
                Break
              EndIf
            EndIf
            If *command\c = 92 ;"\".
              If methodCount < #COMate_MAXNUMSUBOBJECTS And parse(methodCount)\numOpenBrackets = parse(methodCount)\numCloseBrackets And parse(methodCount)\numOpenBrackets <=1 And parse(methodCount)\numberoftokens And numEquals = 0
                methodCount+1
                parse(methodCount)\numberoftokens=0
                parse(methodCount)\numOpenBrackets = 0
                parse(methodCount)\numCloseBrackets = 0
                charPos+1
                *command + SizeOf(CHARACTER)
                *buffer = buffer : *buffer\c = 0 ;Null.
              Else
                methodCount = 0
                Break
              EndIf
            Else
              *buffer\c = *command\c : *buffer + SizeOf(CHARACTER) : *buffer\c = 0
              parse(methodCount)\tokens$[parse(methodCount)\numberoftokens]=*ptrString\s
              parse(methodCount)\numberoftokens+1
              charPos+1
              *command + SizeOf(CHARACTER)
              *buffer = buffer : *buffer\c = 0 ;Null.
            EndIf
          Else          
            charPos+1
            *command + SizeOf(CHARACTER)
            *buffer = buffer : *buffer\c = 0 ;Null.
          EndIf
        ElseIf charPos = length
          *buffer\c = *command\c : *buffer + SizeOf(CHARACTER)
          *buffer\c = 0 ;Null.
          parse(methodCount)\tokens$[parse(methodCount)\numberoftokens]=*ptrString\s
          parse(methodCount)\numberoftokens+1
          charPos + 1
        Else          
          *buffer\c = *command\c : *buffer + SizeOf(CHARACTER) : *buffer\c = 0 ;Null.
          *command+SizeOf(character)
          charPos + 1
        EndIf
      Until charPos > length Or parse(methodCount)\numberOfTokens = #COMate_MAXNUMSYMBOLSINALINE
      FreeMemory(buffer)    
    EndIf
  EndIf
  If methodCount And (parse(methodCount)\numOpenBrackets <> parse(methodCount)\numCloseBrackets Or parse(methodCount)\numOpenBrackets > 1 Or parse(methodCount)\numberoftokens=0)
    methodCount = 0 ;Error.
  EndIf
  ProcedureReturn methodCount
EndProcedure
;///////////////////////////////////////////////////////////////////////////////////////////


;///////////////////////////////////////////////////////////////////////////////////////////
;The following function compiles the tokenised command corresponding to a sub-object invokation within a command string.
;Returns a HRESULT.
Procedure.i COMatePLUS_CompileSubobjectInvokation(*hStatement._COMatePLUSStatement, subObjectIndex, Array parse._COMateParse(1))
  Protected result = #S_OK, i, *varArg.VARIANT
  Protected parseIndex, currentArg$, blnInsideParanthesis, lastArgType, blnByRef, t1$, vt, *cObject._membersCOMateClass, iDispatch.IDISPATCH
  ;Allocate memory for a variant array to hold the arguments.
    *hStatement\ptrVarArgs[subObjectIndex] = AllocateMemory(#COMate_MAXNUMVARIANTARGS*SizeOf(VARIANT))
  If *hStatement\ptrVarArgs[subObjectIndex]
    ;Set *varArg to point at the last variant in the variant array which is to hold the first parameter,
      *varArg = *hStatement\ptrVarArgs[subObjectIndex] + (#COMate_MAXNUMVARIANTARGS - 1) * SizeOf(VARIANT)
    While parseIndex < parse(subObjectIndex)\numberOfTokens
      currentArg$ = parse(subObjectIndex)\tokens$[parseIndex]
      Select currentArg$
        Case "("
          If parseIndex<>1
            result = #E_INVALIDARG
            Break
          EndIf
          blnInsideParanthesis = #True
          lastArgType = #COMate_OpenParanthesis
        Case ")"
          If lastArgType = #COMate_OpenParanthesis Or lastArgType = #COMate_Operand
            lastArgType = #COMate_CloseParanthesis
            blnInsideParanthesis = #False
          Else
            result = #E_INVALIDARG
            Break
          EndIf
        Case "="
          If (lastArgType = #COMate_CloseParanthesis Or lastArgType = #COMate_Method)
            lastArgType = #COMate_Operator
          Else
            result = #E_INVALIDARG
            Break
          EndIf
        Case ","
          If blnInsideParanthesis And lastArgType = #COMate_Operand
            lastArgType = #COMate_Operator
          Else
            result = #E_INVALIDARG
            Break
          EndIf
        Default ;Method or the beginning of an operand.
          If parseIndex = 0
            lastArgType = #COMate_Method
            *hStatement\methodName[subObjectIndex] = COMate_MakeBSTR(currentArg$)
            If *hStatement\methodName[subObjectIndex] = 0
              result = #E_OUTOFMEMORY
              Break
            EndIf
          ElseIf (lastArgType = #COMate_OpenParanthesis) Or (lastArgType = #COMate_Operator);Cannot have 2 operands together.
            If *varArg < *hStatement\ptrVarArgs[subObjectIndex]
              result = #E_OUTOFMEMORY
              Break
            EndIf
            blnByRef = #False
            lastArgType = #COMate_Operand
            ;We must add the operand to the variant array.
            ;First task is to determine the parameter type.  We first examine the operand and decide on the most likely variant format, creating
            ;a variant argument as appropriate. We then see if the user has supplied a 'type modifier', in which case we use VariantChangeType_() etc.
              *varArg\vt = #VT_BSTR ;Default.
              t1$ = LCase(currentArg$)
              If t1$ = "#nullstring"
                currentArg$ = ""
              EndIf
              If Left(currentArg$,1) = "'" Or currentArg$ = "";BSTR
                currentArg$ = Mid(currentArg$, 2, Len(currentArg$)-2)
                ;We parse the string looking for 'escape' sequences.
                  If currentArg$ 
                    COMate_INTERNAL_EscapeString(@currentArg$)
                  EndIf
              Else
                Select t1$
                  Case "#false"
                    *varArg\vt = #VT_BOOL
                    *varArg\boolVal = #VARIANT_FALSE
                  Case "#true"
                    *varArg\vt = #VT_BOOL
                    *varArg\boolVal = #VARIANT_TRUE
                  Case "#empty", "#optional", "#opt" ;Used for optional parameters.
                    *varArg\vt = #VT_ERROR
                    *varArg\scode = #DISP_E_PARAMNOTFOUND
                  Case "#void"
                    If SizeOf(result) = 4
                      *varArg\vt = #VT_I4
                      *varArg\lval = 0
                    Else
                      *varArg\vt = #VT_I8
                      *varArg\llval = 0
                    EndIf
                  Default ;Here we check for numeric types.
                    If COMate_INTERNAL_CheckNumeric(currentArg$, *varArg) = 0
                      result = #E_INVALIDARG  ;No other type of valid operand.
                      Break
                    EndIf
                EndSelect
              EndIf
              If result = #S_OK And *varArg\vt = #VT_BSTR
                *varArg\bstrVal = COMate_MakeBSTR(currentArg$)
                If *varArg\bstrVal = 0
                  result = #E_OUTOFMEMORY
                  Break
                EndIf
              EndIf
              If parseIndex < parse(subObjectIndex)\numberOfTokens-1 And LCase(parse(subObjectIndex)\tokens$[parseIndex+1]) = "byref"
                blnByRef = #True
                parseIndex+1              
              EndIf
            ;Now check for a 'type modifier' which is signified by the presence of a 'AS <operand type>' etc.
              vt = *varArg\vt
              If parseIndex < parse(subObjectIndex)\numberOfTokens-2 And LCase(parse(subObjectIndex)\tokens$[parseIndex+1]) = "as"
                t1$ = LCase(parse(subObjectIndex)\tokens$[parseIndex+2])
                parseIndex + 2
                Select t1$
                  Case "boolean" : vt = #VT_BOOL
                  Case "string", "bstr" : vt = #VT_BSTR
                  Case "byte" : vt = #VT_I1
                  Case "ubyte" : vt = #VT_UI1
                  Case "word" : vt = #VT_I2
                  Case "uword" : vt = #VT_UI2
                  Case "long", "dword" : vt = #VT_I4
                  Case "ulong", "udword" : vt = #VT_UI4
                  Case "quad", "qword" : vt = #VT_I8
                  Case "uquad", "uqword" : vt = #VT_UI8
                  Case "integer", "int" : vt = #VT_INT
                  Case "uinteger", "uint" : vt = #VT_UINT
                  Case "date" : vt = #VT_DATE
                  Case "object", "idispatch", "comateobject" : vt = #VT_DISPATCH
                  Case "iunknown" : vt = #VT_UNKNOWN
                  Case "float", "single" : vt = #VT_R4
                  Case "double" : vt = #VT_R8
                  Case "variant" : vt = #VT_VARIANT
                  Default
                    result = #E_INVALIDARG
                    Break
                EndSelect
              EndIf
              If parseIndex < parse(subObjectIndex)\numberOfTokens-1 And LCase(parse(subObjectIndex)\tokens$[parseIndex+1]) = "byref"
                blnByRef = #True
                parseIndex+1              
              EndIf
              ;Now modify the underlying parameter depending on it's type and whether it is being passed by reference etc.
              ;Note that objects being passed by reference will NOT have their reference counts increased.
                If blnByRef 
                  Select *varArg\vt
                    Case #VT_I1, #VT_I2, #VT_I4, #VT_I8  ;Only these types (which have already been processed) can hold an address.
                      *varArg\vt = vt | #VT_BYREF
                    Default
                      result = #E_INVALIDARG
                      Break
                  EndSelect
              ;BYVAL.
                ElseIf vt = #VT_DISPATCH
                  Select *varArg\vt
                    Case #VT_I1, #VT_I2, #VT_I4, #VT_I8  ;Only these types (which have already been processed) can hold an address.
                      ;Call the AddRef method manually. A corresponding Release() will ensue when we use VariantClear_() when the underlying statement is freed.
                        If t1$ = "comateobject"
                          *cObject = *varArg\pdispVal
                          *varArg\pdispVal = *cObject\iDisp
                          *cObject\iDisp\AddRef()
                        Else
                          iDispatch = *varArg\pdispVal
                          iDispatch\AddRef()
                        EndIf
                        *varArg\vt = #VT_DISPATCH 
                    Default
                      result = #E_INVALIDARG
                      Break
                  EndSelect
                ElseIf vt = #VT_UNKNOWN
                  Select *varArg\vt
                    Case #VT_I1, #VT_I2, #VT_I4, #VT_I8  ;Only these types (which have already been processed) can hold an address.
                      ;Call the AddRef method manually. A corresponding Release() will ensue when we use VariantClear_() when the underlying statement is freed.
                        iDispatch = *varArg\punkVal
                        iDispatch\AddRef()
                        *varArg\vt = #VT_UNKNOWN 
                    Default
                      result = #E_INVALIDARG
                      Break
                  EndSelect
                ElseIf vt = #VT_VARIANT ;We physically copy the variant into the VarArray().
                  Select *varArg\vt
                    Case #VT_I1, #VT_I2, #VT_I4, #VT_I8  ;Only these types (which have already been processed) can hold an address.
                      If *varArg\llVal
                        result = VariantCopy_(*varArg, *varArg\llVal)
                        If FAILED(result)
                          Break
                        EndIf
                      EndIf
                    Default
                      result = #E_INVALIDARG
                      Break
                  EndSelect
                ElseIf *varArg\vt <> vt
                  result = VariantChangeType_(*varArg, *varArg, 16, vt)
                  If FAILED(result)
                    Break
                  EndIf
                EndIf
              *hStatement\numArgs[subObjectIndex] + 1
              *varArg - SizeOf(VARIANT)
          Else
            result = #E_INVALIDARG
            Break
          EndIf
      EndSelect
      parseIndex+1
    Wend
  Else
    result = #E_OUTOFMEMORY
  EndIf
  ProcedureReturn result
EndProcedure
;///////////////////////////////////////////////////////////////////////////////////////////


;-=======================
;-COMateEnum OBJECT CODE.
;-=======================

;-COMateEnum CLASS METHODS.
;----------------------------------------------

;=================================================================================
;Returns the next object in the underlying enumeration in the form of a COMate object (zero if an error).
;The object should be released as soon as it is no longer required.
;Any HRESULT return value is accessible through the GetLastErrorCode() method of the parent COMate object.
Procedure.i COMateEnumClass_GetNextObject(*this._membersCOMateEnumClass)
  Protected result.i = #S_OK, retVar.VARIANT, *newObject._membersCOMateClass
  result = *this\iEV\Next(1, retVar, 0)
  If result = #S_OK ;Alternative is #S_FALSE.
    If retVar\vt <> #VT_DISPATCH
      result = VariantChangeType_(retVar, retVar, 0, #VT_DISPATCH)
    EndIf
    If SUCCEEDED(result)
      ;We create a new COMate object to house the new object.
        *newObject = AllocateMemory(SizeOf(_membersCOMateClass))
        If *newObject
          *newObject\vTable = ?VTable_COMateClass
          *newObject\iDisp = retVar\pdispVal
          COMate_gNumObjects+1
        Else      
          VariantClear_(retVar)
          result = #E_OUTOFMEMORY
        EndIf
    Else
      VariantClear_(retVar)
    EndIf
  EndIf
  ;Set any error code. iDispatch errors will alreay have been set.
    CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
      COMateClass_INTERNAL_SetError(*this\parent, result)
    CompilerEndIf
  ProcedureReturn *newObject
EndProcedure
;=================================================================================


;=================================================================================
;Returns a pointer to a new variant which represents the next variant in the underlying enumeration (zero if an error).
;The variant should be 'freed' by the user with VariantClear_() etc.
;Any HRESULT return value is accessible through the GetLastErrorCode() method of the parent COMate object.
Procedure.i COMateEnumClass_GetNextVariant(*this._membersCOMateEnumClass)
  Protected result.i = #S_OK, *retVar.VARIANT
  ;Allocate memory for a new variant.
    *retVar = AllocateMemory(SizeOf(VARIANT))
  If *retVar
    result = *this\iEV\Next(1, *retVar, 0)
    If result <> #S_OK ;Alternative is #S_FALSE.
      FreeMemory(*retVar)
      *retVar = 0
    EndIf
  Else
    result = #E_OUTOFMEMORY
  EndIf
  ;Set any error code. iDispatch errors will alreay have been set.
    CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
      COMateClass_INTERNAL_SetError(*this\parent, result)
    CompilerEndIf
  ProcedureReturn *retVar
EndProcedure
;=================================================================================


;=================================================================================
;The following method Resets the enumeration back to the beginning.
;Returns a HRESULT value. #S_OK for no errors.
Procedure.i COMateEnumClass_Reset(*this._membersCOMateEnumClass)
  Protected result.i
  If *this\iEV ;Just in case.
    ;Reset underlying IEnumVARIANT object.
      result = *this\iEV\Reset()
  EndIf
    CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
      COMateClass_INTERNAL_SetError(*this\parent, result)
    CompilerEndIf
  ProcedureReturn result
EndProcedure
;=================================================================================


;=================================================================================
;The following method releases a com object created by any of the functions which return object pointers.
Procedure COMateEnumClass_Release(*this._membersCOMateEnumClass)
  If *this\iEV ;Just in case.
    ;Release underlying IEnumVARIANT object.
      *this\iEV\Release()
  EndIf
  ;Free object.
    FreeMemory(*this)
EndProcedure
;=================================================================================


;-=======================
;-COM (ActiveX) REGISTRATION FUNCTIONS.
;-=======================

;=================================================================================
;The following function allows the user to register a COM server for the duration of an application's run etc.
;Returns a HRESULT value. #S_OK for no errors.
Procedure.i COMate_RegisterCOMServer(dllName$, blnInitCOM = #True)
  Protected result.i = #S_OK, lib.i, fn.i
  If blnInitCOM 
    CoInitialize_(0)
  EndIf
  If FileSize(dllName$) > 0
    lib = OpenLibrary(#PB_Any, dllName$)
    If lib
      fn = GetFunction(lib, "DllRegisterServer")
      If fn
        result = CallFunctionFast(fn)
      Else
        result = #E_FAIL
      EndIf
      CloseLibrary(lib)
    Else
      result = #E_FAIL
    EndIf
  Else
    result = #E_INVALIDARG
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(0, result)
  CompilerEndIf
  ProcedureReturn result
EndProcedure
;=================================================================================


;=================================================================================
;The following function allows the user to unregister a COM server after registering it with COMate_RegisterActiveXServer().
;Returns a HRESULT value. #S_OK for no errors.
Procedure.i COMate_UnRegisterCOMServer(dllName$, blnInitCOM = #True)
  Protected result.i = #S_OK, lib.i, fn.i
  If blnInitCOM 
    CoInitialize_(0)
  EndIf
  If FileSize(dllName$) > 0
    lib = OpenLibrary(#PB_Any, dllName$)
    If lib
      fn = GetFunction(lib, "DllUnregisterServer")
      If fn
        result = CallFunctionFast(fn)
      Else
        result = #E_FAIL
      EndIf
      CloseLibrary(lib)
    Else
      result = #E_FAIL
    EndIf
  Else
    result = #E_INVALIDARG
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(0, result)
  CompilerEndIf
  ProcedureReturn result
EndProcedure
;=================================================================================


;-=======================
;-MISCELLANEOUS FUNCTIONS.
;-=======================

;=================================================================================
;The following function searches the registry for the given textual representation of an interface IID and, if successful, copies the
;actual IID to the specified buffer.
;Returns a HRESULT.
Procedure.i COMate_GetIIDFromName(name$, *iid.IID)
  Protected result = #E_FAIL, error, hKey1, hKey2, enumIndex, subKey, lpcbName = 256, cbData = 256, buffer
  Protected bstr
  If name$ And *iid
    subKey = AllocateMemory(lpcbName)
    If subKey
      buffer = AllocateMemory(cbData)
      If buffer
        If RegOpenKeyEx_(#HKEY_CLASSES_ROOT, "Interface", 0, #KEY_READ, @hKey1) = #ERROR_SUCCESS And hKey1
          enumIndex = 0
          error = RegEnumKeyEx_(hKey1, enumIndex, subKey, @lpcbName, 0, 0, 0, 0)
          While error = #ERROR_SUCCESS
            If RegOpenKeyEx_(hKey1, subKey, 0, #KEY_READ, @hKey2) = #ERROR_SUCCESS And hKey2
              cbData = 256
              If RegQueryValueEx_(hKey2, "", 0, 0, buffer, @cbData) = #ERROR_SUCCESS
                If PeekS(buffer) = name$ ;We have the correct entry.
                  ;Attempt to create an IID from the string representation of the IID.
                    bstr = COMate_MakeBSTR(PeekS(subKey))
                    If bstr
                      result = CLSIDFromString_(bstr, *iid)
                      SysFreeString_(bstr)
                    Else
                      result = #E_OUTOFMEMORY
                    EndIf
                  Break
                EndIf
              EndIf
              RegCloseKey_(hKey2)
            EndIf
            lpcbName = 256
            enumIndex + 1
            error = RegEnumKeyEx_(hKey1, enumIndex, subKey, @lpcbName, 0, 0, 0, 0)
          Wend
          RegCloseKey_(hKey1)
        EndIf
        FreeMemory(buffer)
      EndIf
      FreeMemory(subKey)
    EndIf
  Else
    result = #E_INVALIDARG
  EndIf
  CompilerIf Defined(COMATE_NOERRORREPORTING, #PB_Constant)=0
    COMateClass_INTERNAL_SetError(0, result)
  CompilerEndIf
  ProcedureReturn result
EndProcedure
;=================================================================================


;-=======================
;-ERROR RETRIEVAL FUNCTIONS.
;-=======================

;=================================================================================
;The following function returns the last error HRESULT code recorded by COMate against the underlying thread.
;This is completely threadsafe in that 2 threads using the same COMate object will not overwrite each other's errors.
Procedure.i COMate_GetLastErrorCode()
  Protected *error._COMateThreadErrors
  If COMate_gErrorTLS And COMate_gErrorTLS <> -1
    *error = TlsGetValue_(COMate_gErrorTLS)
    If *error
      ProcedureReturn *error\lastErrorCode
    EndIf
  EndIf
EndProcedure
;=================================================================================


;=================================================================================
;The following function returns a description of the last error recorded by COMate against the underlying thread.
;This is completely threadsafe in that 2 threads using the same COMate object will not overwrite each other's errors.
Procedure.s COMate_GetLastErrorDescription()
  Protected *error._COMateThreadErrors
  If COMate_gErrorTLS And COMate_gErrorTLS <> -1
    *error = TlsGetValue_(COMate_gErrorTLS)
    If *error
      ProcedureReturn *error\lastError$ 
    EndIf
  EndIf
EndProcedure
;=================================================================================



;-=======================
;-UTILITY FUNCTIONS.
;-=======================

;/////////////////////////////////////////////////////////////////////////////////
;The following function converts a string (Ascii or Unicode) to an OLE string.
;We access this through a prototype.
Procedure.i COMateClass_UTILITY_MakeBSTR(value)
  Protected result.i
  result = SysAllocString_(value)
  ProcedureReturn result
EndProcedure
;/////////////////////////////////////////////////////////////////////////////////


DataSection

  VTable_COMateClass:
    Data.i @COMateClass_Invoke() 
    Data.i @COMateClass_Release() 
    Data.i @COMateClass_CreateEnumeration()
    Data.i @COMateClass_GetCOMObject()
    Data.i @COMateClass_GetContainerhWnd()
    Data.i @COMateClass_SetDesignTimeMode()
    Data.i @COMateClass_GetDateProperty()
    Data.i @COMateClass_GetIntegerProperty()
    Data.i @COMateClass_GetObjectProperty()
    Data.i @COMateClass_GetRealProperty()
    Data.i @COMateClass_GetStringProperty()
    Data.i @COMateClass_GetVariantProperty()
    Data.i @COMateClass_SetProperty()
    Data.i @COMateClass_SetPropertyRef()
CompilerIf Defined(COMATE_NOINCLUDEATL, #PB_Constant)=0
    Data.i @COMateClass_SetEventHandler()
    Data.i @COMateClass_GetIntegerEventParam()
    Data.i @COMateClass_GetObjectEventParam()
    Data.i @COMateClass_GetRealEventParam()
    Data.i @COMateClass_GetStringEventParam()
    Data.i @COMateClass_IsEventParamPassedByRef()
CompilerEndIf
  VTable_COMateEnumClass:
    Data.i @COMateEnumClass_GetNextObject()
    Data.i @COMateEnumClass_GetNextVariant()
    Data.i @COMateEnumClass_Reset()
    Data.i @COMateEnumClass_Release() 

CompilerIf Defined(COMATE_NOINCLUDEATL, #PB_Constant)=0
  VTable_COMateEventSink:
    Data.i @COMateSinkClass_QueryInterface()
    Data.i @COMateSinkClass_AddRef()
    Data.i @COMateSinkClass_Release()
    Data.i @COMateSinkClass_GetTypeInfoCount()
    Data.i @COMateSinkClass_GetTypeInfo()
    Data.i @COMateSinkClass_GetIDsOfNames()
    Data.i @COMateSinkClass_Invoke()
CompilerEndIf

  IID_NULL: ; {00000000-0000-0000-0000-000000000000}
    Data.l $00000000
    Data.w $0000, $0000
    Data.b $00, $00, $00, $00, $00, $00, $00, $00

  IID_IUnknown: ; {00000000-0000-0000-C000-000000000046}
    Data.l $00000000
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46

  IID_IDispatch: ; {00020400-0000-0000-C000-000000000046}
    Data.l $00020400
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46

  IID_IClassFactory: ; {00000001-0000-0000-C000-000000000046}
    Data.l $00000001
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46

  IID_IPersistFile: ; {0000010B-0000-0000-C000-000000000046}
    Data.l $0000010B
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46

  IID_IEnumVARIANT: ; {00020404-0000-0000-C000-000000000046}
    Data.l $00020404
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46

  IID_IConnectionPointContainer: ; {B196B284-BAB4-101A-B69C-00AA00341D07}
    Data.l $B196B284
    Data.w $BAB4, $101A
    Data.b $B6, $9C, $00, $AA, $00, $34, $1D, $07

  IID_IAxWinAmbientDispatch: ; {B6EA2051-048A-11D1-82B9-00C04FB9942E}
    Data.l $B6EA2051
    Data.w $048A, $11D1
    Data.b $82, $B9, $00, $C0, $4F, $B9, $94, $2E

EndDataSection

CompilerEndIf

; IDE Options = PureBasic 5.70 LTS (Windows - x86)
; ExecutableFormat = Shared dll
; CursorPosition = 36
; Folding = ----------
; EnableThread
; Executable = nxReportU.dll
; EnableUnicode