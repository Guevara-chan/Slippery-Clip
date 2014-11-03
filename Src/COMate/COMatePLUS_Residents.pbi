;/////////////////////////////////////////////////////////////////////////////////
;***COMatePLUS***  COM OLE automation through iDispatch.
;*===========
;*
;*©nxSoftWare (www.nxSoftware.com) 2009.
;*======================================
;*
;*Header file.
;/////////////////////////////////////////////////////////////////////////////////

  #COMate_UnknownObjectType = 1 ;Use this with the GetObjectProperty() method when the object type does not inherit from iDispatch.
                                ;In such cases, the method will return the interface pointer directly (as opposed to a COMate object!)

  ;The following constant is used with the SetEventHandler() method of the COMateObject class in order to set an optional
  ;'global' handler for all events (useful for examining which events are being sent etc.)
  ;This is in addition to any individual handlers which are called after this global one.
    #COMate_CatchAllEvents = ""
  ;The following enumeration is used with the SetEventHandler() method of the COMateObject class to specify the return type
  ;(if any) of an individual event.
    Enumeration
      #COMate_NoReturn
      #COMate_IntegerReturn
      #COMate_RealReturn
      #COMate_StringReturn
      #COMate_OtherReturn   ;This is a special case; please see the COMate manual for details (SetEventHandler()).
        #COMate_VariantReturn = #COMate_OtherReturn ;Alias!
        #COMate_UnknownReturn = #COMate_OtherReturn ;Alias!
    EndEnumeration


;/////////////////////////////////////////////////////////////////////////////////
;-Declaration of 'Public' functions.
  Declare.i COMate_CreateObject(progID$, hWnd = 0, blnInitCOM = #True)
  Declare.i COMate_GetObject(file$, progID$="", blnInitCOM = #True)
  Declare.i COMate_WrapCOMObject(object.iUnknown)
  ;Statement functions.
    Declare.i COMate_PrepareStatement(command$)               ;Returns a statement handle or zero if an error.
    Declare.i COMate_GetStatementParameter(hStatement, index) ;Returns, if successful, a direct pointer to the appropriate variant structure.
                                                               ;Index is 1-based.
    Declare COMate_FreeStatementHandle(hStatement)
  ;The two error retrieval functions are completely threadsafe in that, for example, 2 threads could be working with the same COMate object
  ;and any resuting errors will be stored separately so that one thread's errors will not conflict with another's etc.
    Declare.i COMate_GetLastErrorCode()
    Declare.s COMate_GetLastErrorDescription()
  ;OCX (ActiveX) functions.
    Declare.i COMate_RegisterCOMServer(dllName$, blnInitCOM = #True)
    Declare.i COMate_UnRegisterCOMServer(dllName$, blnInitCOM = #True)
    Declare.i COMate_CreateActiveXControl(x, y, width, height, progID$, blnInitCOM = #True)
  ;Miscellaneous.
    Declare.i COMate_GetIIDFromName(name$, *iid.IID)
;/////////////////////////////////////////////////////////////////////////////////


;/////////////////////////////////////////////////////////////////////////////////
;-Class interfaces.

  ;The following interface details the class methods for COMateObject type objects; the main object type for COMate.
    Interface COMateObject
      ;General methods.
      ;=================================
        Invoke.i(command$, *hStatement=0)
                                        ;Returns a HRESULT value. #S_OK for no errors.
        Release()                       ;DO NOT use this until all enumeration objects attached to this object have been freed.
        CreateEnumeration.i(command$, *hStatement=0)
                                        ;Returns an object of type COMateEnumObject (see below) or zero if an error.
        GetCOMObject.i()                ;Returns the COMate object's underlying iDispatch object pointer. AddRef() is called on this object
                                        ;so the developer must call Release() at some point.
        GetContainerhWnd.i(returnCtrlID=0)
                                        ;In the case of an ActiveX control, this methods returns either the handle of the container used
                                        ;to house the control or the Purebasic gadget#. Returning the gadget# is only viable if using COMate
                                        ;as a source code include (or a Tailbitten library!)
        SetDesignTimeMode.i(state=#True);In the case of an ActiveX control, this methods attempts to set a design time mode.
                                        ;Returns a HRESULT value. #S_OK for no errors.
      ;Get property methods.
      ;=================================
        GetDateProperty.i(command$, *hStatement=0)
                                        ;Returns a PB date value. Of course you can always retrieve a data in string form using GetStringProperty() etc.
        GetIntegerProperty.q(command$, *hStatement=0)
                                        ;Returns a signed quad which can of course be placed into any integer variable; byte, word etc.
        GetObjectProperty.i(command$, *hStatement=0, objectType = #VT_DISPATCH)
                                        ;Returns a COMate object or an iUnknown interface pointer depending on the value of the 'objectType' parameter.
                                        ;For 'regular' objects based upon iDispatch, leave the optional parameter 'objectType' as it is.
                                        ;Otherwise, for unknown object types set objectType to equal #COMate_UnknownObjectType. In these cases,
                                        ;this method will return the interface pointer directly (as opposed to a COMate object).
                                        ;In either case the object should be released as soon as it is no longer required.
        GetRealProperty.d(command$, *hStatement=0)
                                        ;Returns a double value which can of course be placed into any floating point variable; float or double.
        GetStringProperty.s(command$, *hStatement=0)
        GetVariantProperty.i(command$, *hStatement=0)
                                        ;For those returns which are not directly supported by the COMate functions.
                                        ;The user must use VariantClear_() and FreeMemory() when finished with the variant returned by this method.
      ;Set property methods.
      ;=================================
        SetProperty.i(command$, *hStatement=0)
                                        ;Returns a HRESULT value. #S_OK for no errors.
        SetPropertyRef.i(command$, *hStatement=0)
                                        ;Returns a HRESULT value. #S_OK for no errors.
CompilerIf Defined(COMATE_NOINCLUDEATL, #PB_Constant)=0
      ;Event handler methods.
      ;=================================
        SetEventHandler.i(eventName$, callback, returnType = #COMate_NORETURN, *riid.IID=0)
                                        ;eventName$ = "" to set an optional handler which will receive all events (useful for examining which events are sent).
                                        ;This is in addition to any individual handlers which are called after this 'global' one.
                                        ;Returns a HRESULT value. #S_OK for no errors. Set callback = 0 to remove an existing event handler.
        GetIntegerEventParam.q(index)   ;Only valid to call this during an event handler. Index is a 1-based index.
        GetObjectEventParam.i(index, objectType = #VT_DISPATCH)
                                        ;Only valid to call this during an event handler. Index is a 1-based index.
                                        ;The object returned is NOT in the form of a COMate object. Use COMate_WrapCOMObject() to convert to
                                        ;a COMate object if required.
                                        ;User must call Release() on this object when done.
        GetRealEventParam.d(index)      ;Only valid to call this during an event handler. Index is a 1-based index.
        GetStringEventParam.s(index)    ;Only valid to call this during an event handler. Index is a 1-based index.
        IsEventParamPassedByRef.i(index, *ptrParameter.INTEGER=0, *ptrVariant.INTEGER=0)
                                        ;Returns zero or a variant #VT_... type constant (minus the #VT_BYREF modifier).
                                        ;In the latter case, and if *ptrParameter is non-zero, then the address of the underlying parameter
                                        ;is placed into this buffer, enabling the client application to modify the parameter etc.
CompilerEndIf
    EndInterface


  ;The following interface details the class methods for COMateEnumObject type objects.
  ;Instances of these objects are used to enumerate collections of objects (or variants) exposed by a COM object.
  ;These objects are created through the CreateEnumeration() method of the COMateObject class (see above).
    Interface COMateEnumObject
      GetNextObject.i()                 ;Returns a 'COMateObject' object (or zero if the enumeration is complete).
                                        ;The user must use the Release() method on this object when it is no longer required.
      GetNextVariant.i()                ;Returns a pointer to a variant (or zero if there are no more).
                                        ;The user must use VariantClear_() and FreeMemory() when finished with the variant returned by this method.
      Reset.i()                         ;Resets the enumeration back to the beginning. NOTE that there is no guarantee that a second run
                                        ;through the underlying collection will produce the same results. Imagine a collection of files in a
                                        ;folder for example in which some can be deleted between enumerations etc.
                                        ;Returns a HRESULT value. #S_OK for no errors.
      Release()
    EndInterface
;/////////////////////////////////////////////////////////////////////////////////

; IDE Options = PureBasic 4.31 Beta 1 (Windows - x86)
; ExecutableFormat = Shared Dll
; CursorPosition = 75
; FirstLine = 52
; EnableUnicode
; EnableThread
; Executable = nxReportU.dll