; -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
; IRichEditOleCallback v0.6 (Beta)
; Adopted in 2010 by Guevara-chan.
; -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Structure RichEditOle 
   *pIntf.IRicheditOle 
   Refcount.l 
   hwnd.l 
EndStructure 

Global NewMap RichComObject.RichEditOle()

Procedure.l RichEdit_SetInterface(hWnd) 
AddMapElement(RichComObject(),Hex(hWnd)) 
RichComObject()\pIntf = ?VTable 
RichComObject()\hwnd=hwnd 
SendMessage_(hWnd, #EM_SETOLECALLBACK, 0, RichComObject()) 
ProcedureReturn RichComObject() 
EndProcedure 

Procedure.l RichEdit_QueryInterface(*pObject.RichEditOle, REFIID, *ppvObj.LONG) 
  Protected *pointeur.IRicheditOle 
  *pointeur=*pObject 
  If CompareMemory(REFIID, ?IID_IUnknown2, 16)=1 Or CompareMemory(REFIID, ?IID_IRichEditOleCallback, 16)=1 
    Debug "QueryInterface" 
    *ppvObj\l = *pObject 
    *pointeur\AddRef() 
    ProcedureReturn #S_OK 
  Else 
    *ppvObject=0 
    ProcedureReturn #E_NOINTERFACE 
  EndIf 
EndProcedure 

Procedure.l RichEdit_AddRef(*pObject.RichEditOle)
  *pObject\Refcount+1
  ProcedureReturn *pObject\Refcount
EndProcedure

Procedure.l RichEdit_Release(*pObject.RichEditOle)
  *pObject\Refcount-1
  If *pObject\Refcount > 0
    ProcedureReturn *pObject\Refcount
  Else
;Remove entry in map.
    DeleteMapElement(RichComObject(), Hex(*pObject\hWnd))
      *pObject=0
  EndIf
EndProcedure

Procedure.l RichEdit_GetInPlaceContext(*pObject.RichEditOle, lplpFrame, lplpDoc, lpFrameInfo)
Debug 1
  ProcedureReturn #E_NOTIMPL
EndProcedure

Procedure.l RichEdit_ShowContainerUI(*pObject.RichEditOle, fShow)
  ProcedureReturn #E_NOTIMPL
EndProcedure

Procedure.l RichEdit_QueryInsertObject(*pObject.RichEditOle, lpclsid, lpstg, cp)
    ProcedureReturn #S_OK
EndProcedure

Procedure.l RichEdit_DeleteObject(*pObject.RichEditOle, lpoleobj)
  ProcedureReturn #E_NOTIMPL
EndProcedure

Procedure.l RichEdit_QueryAcceptData(*pObject.RichEditOle, lpdataobj, lpcfFormat, reco, fReally, hMetaPict)
    ProcedureReturn #S_OK
EndProcedure

Procedure.l RichEdit_ContextSensitiveHelp(*pObject.RichEditOle, fEnterMode)
  ProcedureReturn #E_NOTIMPL
EndProcedure

Procedure.l RichEdit_GetClipboardData(*pObject.RichEditOle, lpchrg, reco, lplpdataobj)
  ProcedureReturn #E_NOTIMPL
EndProcedure

Procedure.l RichEdit_GetDragDropEffect(*pObject.RichEditOle, fDrag, grfKeyState, pdwEffect)
;PokeL(pdwEffect,0) ;Uncomment this to prevent dropping to the editor gadget.
  ProcedureReturn #E_NOTIMPL
EndProcedure

Procedure.l RichEdit_GetContextMenu(*pObject.RichEditOle, seltype.w, lpoleobj, lpchrg, lphmenu)
  ProcedureReturn #E_NOTIMPL
EndProcedure


;The following function does the main work!
Procedure.l RichEdit_GetNewStorage(*pObject.RichEditOle, lplpstg)
  Protected sc, lpLockBytes, t.ILockBytes
;Attempt to create a byte array object which acts as the 'foundation' for the upcoming compound file.
  sc=CreateILockBytesOnHGlobal_(#Null, #True, @lpLockBytes) 
  If sc ;This means that the allocation failed.
    ProcedureReturn sc
  EndIf
;Allocation succeeded so we now attempt to create a compound file storage object.
  sc=StgCreateDocfileOnILockBytes_(lpLockBytes, #STGM_SHARE_EXCLUSIVE|#STGM_READWRITE|#STGM_CREATE, 0, lplpstg)
  t = lpLockBytes
  t\Release()  
EndProcedure
;***********************************************************************************************



DataSection
VTable: 
      Data.l @RichEdit_QueryInterface(), @RichEdit_AddRef(), @RichEdit_Release(), @RichEdit_GetNewStorage()
      Data.l @RichEdit_GetInPlaceContext(), @RichEdit_ShowContainerUI(), @RichEdit_QueryInsertObject()
      Data.l @RichEdit_DeleteObject(), @RichEdit_QueryAcceptData(), @RichEdit_ContextSensitiveHelp(), @RichEdit_GetClipboardData()
      Data.l @RichEdit_GetDragDropEffect(), @RichEdit_GetContextMenu()

IID_IRichEditOleCallback: ;" 0x00020D03, 0, 0, 0xC0,0,0,0,0,0,0,0x46" 
Data.l $00020D03 
Data.w $0000,$0000 
Data.b $C0,$00,$00,$00,$00,$00,$00,$46  

IID_IUnknown2:   ;"{00000000-0000-0000-C000-000000000046}" 
Data.l $00000000 
Data.w $0000,$0000 
Data.b $C0,$00,$00,$00,$00,$00,$00,$46  

EndDataSection
; IDE Options = PureBasic 5.21 LTS (Windows - x86)
; Folding = ---
; EnableXP