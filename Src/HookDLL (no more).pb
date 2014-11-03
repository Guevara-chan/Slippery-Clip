; -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
; Slippery clip|board hooking v0.9
; Adopted in 2010 by Guevara-chan.
; Not used anymore, yet I love this.
; -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Structure Hook_Data
OriBytes.b[6]
NewBytes.b[6]
ProcAddr.l
EndStructure

Global ThisHook.Hook_Data
Global HookMessage

Procedure.i Setup_Hook(Libname.s, FuncName.s, RedFunAddr.l)
dwAddr = GetProcAddress_(GetModuleHandle_(LibName), FuncName) ;Get the Function Memory Address
OriginalAdress = dwAddr
ThisHook\ProcAddr = dwAddr ;This for Fast Unhooking
If OriginalAdress > 0
;Save the old api Bytes
If ReadProcessMemory_(-1, dwAddr, @ThisHook\OriBytes[0], 6, 0)
Dim a.b(6)
a(0)=$e9 ;JMP
a(5)=$C3 ;RET
dwCalc = RedFunAddr - dwAddr - 5; 
CopyMemory(@dwCalc, @a(1), 4)
;Save the New Bytes
CopyMemory(@a(0), @ThisHook\NewBytes[0], 6)
;Get ready for new hook
If WriteProcessMemory_(-1, dwAddr, @a(0), 6, 0)
ProcedureReturn ThisHook
EndIf
EndIf
EndIf
ProcedureReturn -1 ;Bad didn't work
EndProcedure

Procedure.b Fast_ReHook()
ProcedureReturn WriteProcessMemory_(-1, ThisHook\ProcAddr, @ThisHook\NewBytes[0], 6, 0)
EndProcedure

Procedure.b ReHook(LibName.s, FuncName.s)
dwAddr = GetProcAddress_(GetModuleHandle_(LibName), FuncName) ;Get the Memory Address
ProcedureReturn WriteProcessMemory_(-1, dwAddr, @ThisHook\NewBytes[0], 6, 0)
EndProcedure

Procedure.b Fast_UnHook()
ProcedureReturn WriteProcessMemory_(-1, ThisHook\ProcAddr, @ThisHook\OriBytes[0], 6, 0)
EndProcedure

Procedure.b UnHook(LibName.s, FuncName.s)
dwAddr = GetProcAddress_(GetModuleHandle_(LibName), FuncName) ;Get the Memory Address
ProcedureReturn WriteProcessMemory_(-1, dwAddr, @ThisHook\OriBytes[0], 6, 0) ;Restore Old Bytes
EndProcedure

Procedure.l ClipHook(Format, *DataPtr)
Fast_UnHook()
SetClipboardData_(Format, *DataPtr)
PostMessage_(#HWND_BROADCAST, HookMessage, Format, *DataPtr)
Fast_ReHook()
ProcedureReturn 1
EndProcedure

ProcedureDLL AttachProcess(Instance)
HookMessage = RegisterWindowMessage_("SlipperyClip_Hook")
Setup_Hook("User32.dll", "SetClipboardData", @ClipHook())
EndProcedure

ProcedureDLL DetachProcess(Instance)
UnHook("User32.dll", "SetClipboardData")
EndProcedure

ProcedureDLL MyHookProcedure(nCode.l, wParam.l, *lParam.CWPSTRUCT) ; Dummy hook.
ProcedureReturn CallNextHookEx_(@MyHookProcedure(), nCode, wParam, *lParam)
EndProcedure
; IDE Options = PureBasic 4.51 (Windows - x86)
; ExecutableFormat = Shared Dll
; CursorPosition = 3
; Folding = --
; EnableXP
; Executable = Hook.dll