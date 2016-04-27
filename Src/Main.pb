; *-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-*
; Slippery Clip|board manager v1.21
; Developed in 2010 by Guevara-chan.
; *-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-*

;{ =[TO.DO]=
; TO.DO Визуализировать блокировку входящих.
; TO.DO Добавить всплывающее меню для вставки.
; TO.DO Перенести больше функций на обработку с клавиатуры.
; -----------------------------------------------------------------------------
; TO.DO Вернуть подсветку найденного HTML (некорректно считается смещение).
; TO.DO Исправить неполадки с автопролистыванием неполностью отображенного нода.
; TO.DO Оптимизировать заполнение окна предпросмотра.
; TO.DO Решить вопрос со второй строкой кнопок управления.
; TO.DO Реализовать сохранение истории поисковых запросов.
; TO.DO Добавить режимы поиска бинарной информации.
; TO.DO Улучшить распознание окон-источников.
; TO.DO Улучшить общую стабильность программы.
; TO.DO Переделать обработку событий на Bind Event.
; TO.DO Унифицировать процедуры обработки входящих данных.
; TO.DO Реализовать кеширование данных.
; TO.DO Решить вопрос с дублированием данных при переносе из Word'а.
; TO.DO Оптимизировать поиск начала HTML-кода.
; TO.DO Улучшить поддержку листингов переноса файлов.
; TO.DO Улучшить фильтры и ограничители принимаемых данных.
; TO.DO Улучшить распознавание выделения Paint'а.
; TO.DO Улучшить поддержку Drag\Drop.
; TO.DO Добавить опции преобразования текста в иные типы.
; TO.DO Доделать сериализацию опций.
; TO.DO GUI: Поправить смещение текста при просмотре.
; TO.DO GUI: Поправить потерю крусора при нагрузке.
; TO.DO GUI: Разобраться со средней кнопкой мыши в главном окне.
; TO.DO GUI: Разобраться с мерцанием окон просмотра на +/-.
;} {End/TO.DO}

; --Preparations--
CompilerIf #PB_Compiler_Unicode
IncludeFile "OLEdit.pbi"
IncludeFile "COMate\COMatePLUS.pbi"
EnableExplicit ; Essential.
IncludeFile "SettingsGUI.pbi"
UsePNGImageEncoder()      ; For data saving.
UseJPEGImageEncoder()     ; For data saving.
UseJPEG2000ImageEncoder() ; For data saving.
UseBriefLZPacker()        ; For data saving.
CompilerElse : CompilerError "No. Just no. Try it Unicode next time."
CompilerEndIf ; ^To notify that you done things tremendously wrong^

CompilerIf #PB_Compiler_Version => 540
UseCRC32Fingerprint() ; Legacy codec.
UseMD5Fingerprint()   ; Legacy codec.
CompilerEndIf

;{ [Definitions]
;{ --Constants--
#Kb            = 1024
#UsedNode      = "->"
#UnusedNode    = -1
#BringKey      = 'T'
#SwapKey       = 'Q'
#SaveKey       = 'S'
#Deployment    = 'X'
#Unlimited     = '0'
#HotKeys       = 9
#IniFile       = "Settings.ini"
#DumpFile      = "Board.db"
#MinWidth      = 345
#MinHeight     = 360
#DumpTimer     = 1
#Second        = 1000
#TrayIcon      = 0
#Minute        = #Second * 60
#TitleDecor    = "="
#TitleRaw      = "Slippery Clip"
#Title         = #TitleDecor + #TitleRaw + #TitleDecor
#FullAlpha     = 255
#GlassAlpha    = #FullAlpha / 2
#SlotPrefix    = "Slot #"
#CF_RichText   = -1
#CF_HTML       = -2
#TempImage     = 0
#AllFormats    = 6
#VPEntropy     = 20
#NoLife        = -1
#Hype          = 777
#DirDelim      = " || "
#IDelim        = " \\ "
#KBDigi        = 1
#MaxLineChars     = 1000
#VPortMinWidth    = 200
#VPortMinHeight   = 200
#PackLimit        = 10 * #Kb
#LayoutTable      = #MAXWORD
#CharSize         = SizeOf(Unicode)
#LayoutEdge       = #LayoutTable * #CharSize
#DirDragDelim     = #LF
#DumpSig   = 1564693339 ; '[SC]'
#ReconMsg  = #CR$ + "Procedure aborted. Reconsider list limitations."
#WinFlags  = #PB_Window_SystemMenu|#PB_Window_Tool|#PB_Window_SizeGadget|#PB_Window_Invisible
#VPFlags   = #WinFlags
#DefFlags  = #PB_Font_Bold|#PB_Font_Italic|#PB_Font_HighQuality
#DragOff   = #PB_Drag_Copy|#PB_Drag_Link|#PB_Drag_Move
#CtrlShift = #MOD_CONTROL | #MOD_SHIFT
#SinglerMx = "=[SlipperyClip]="
#ForceMode = '-1'
#CritOff   = -2
#BitDepth  = 24
#Postf     = ": "
#PreHTML   = "<html><body>" 
#PostHTML  = "</body></html>"
#MetaMode  = -7
#ListingExtra = 2
#RunTime      = 100
#SnapSpace    = 30
#VProp         = "ViewportID"
#VertSFlag     = "Top\Bottom"
#HorSFlag      = "Left\Right"
#NoSnap        = "RecentMove"
#SnapTime      = 100
#CacheQuant    = 5 ; Minutes.
#QuantMul      = 2
#SearchLimit   = 1024
#KDownParam    = $1
#KUpParam      = $C0000001
#ErrHdr        = ":IAmError:"
#SSuccess      = -1
#FlagSeparator = "/"
#SearchRXP     = 0
#HotBound      = 1
#HotUnBound    = 2
#MainWindow    = 77
#SeqCode       = '`'
#SeqChar       = Chr(#SeqCode)
#RepCR         = '|'
#RepLF         = '~'
#Denial        = "Deny *"
#RemPref       = "; «"
#InputMsg      = "Input remark string (one line) for "
#NoiseImg      = 0
#FlagMask      = #MAXBYTE
#DeparseStorage = #SearchLimit * 4
; -GUI hard- 
#ButtonHeight  = 20
#ButtonOffset  = 5
#ButtonSpace   = 5
#ButtonBar   = 4
#InfHeight   = 21
#Walley      = 2
#InfBottom   = #Walley + #InfHeight
#SearchHeight  = 17
#SBHeight      = #SearchHeight
#SBWidth       = #SBHeight - 1
#SearchUOffset = 4
#SearchBOffset = #SearchUOffset + #SearchHeight
#SearchLOffset = 5
#SearchROffset = #SearchLOffset * 2 + (#SBWidth - 1) * 2
#ClipLOffset   = 5
#ClipUOffset   = #SearchUOffset + #SearchBOffset
#ClipROffset   = #ClipLOffset * 2
#ClipInformerB = #ButtonHeight + 5 + #InfBottom + #Walley
#ClipBOffset   = #ClipInformerB + #Walley + #ClipUOffset
#UniVPOffset   = 5
#VALOffset     = #UniVPOffset
#VAUOffset     = #UniVPOffset
#VAROffset     = #VALOffset * 2
#VABOffset     = #VAUOffset + #InfBottom + #Walley
#EFOffset      = 2
#EFSize        = #EFOffset * 2
#WinShade      = $C0C0C0
#ShiftStep     = 3
;}
;{ --Enumerations--
Enumeration 77 ; Gadgets
#ClipList
#Button_Clear
#Button_Switch
#Button_Ocular
#Button_Options
#Button_Terminate
#EyeFrame
#VoidGadget
#RTFParser
#HTMLParser
#SearchBar
#LurkBack
#LurkForth
EndEnumeration

Enumeration ; TrayMenu
#tShowWindow
#tOptions
#tTerminate
#tSwitch
#tClearList
#tUseNode ; Anchor.
EndEnumeration

Enumeration #PB_Compiler_EnumerationValue + #HotKeys ; Menu items.
#cMoveUp
#cMoveDown
#cThrowUp
#cThrowDown
#cQWERTYSwap
#cSetComment
#cViewData
#cFlatten
#cRender
#cSaveAs
#cRemove
#cListerate
#cFindNext   ; Virtual item.
#cFindPrev   ; Virtual item.
#cGoSearch   ; Virtual item.
#cCopy       ; Virtual item.
#cOptions    ; Virtual item.
#cPanopticum ; Virtual item.
#cDelete   ; Shared virtual item.
#cReturn   ; Shared virtual item.
#cCtrlA    ; Shared virtual item.
#cCtrlV    ; Shared virtual item.
#cBindData ; Anchor.
EndEnumeration

Enumeration #PB_Compiler_EnumerationValue + #HotKeys ; More menu items.
#vCenterWin
#vMaximize
#vReturnSize
#vSwitchSize
#vSizeUp
#vSizeDown
#vSaveAs
#vSaveSnap
#vCopy
#vCopyMD
#vCopyRaw
#vSelectAll
#vHighlight
#vWordWrap
EndEnumeration

Enumeration #PB_Compiler_EnumerationValue ; Even more menu items.
#sSearchMenuEdge ; Purely virtual item
#sCut : #sCopy : #sPaste : #sClear : #sSelectAll
; -------------------------- ;
#SfOffsetEdge    ; Purely virtual item
#sfText
#sfSrc
#sfRem
#sfID
#sfList
; -------------------------- ;
#sfNoSTR
#sfNoBMP
#sfNoDIR
#sfNoRTF
#sfNoHTML
#sfNoMETA
#sfUnDeny
; -------------------------- ;
#sfCase
#sfNocase
#sfRegular
#sfWhole
#sfOpen
#sfNot
#sfSel
#sfNoSel
#sfHKBound
#sfHKUnbound
EndEnumeration

Enumeration #PB_Compiler_EnumerationValue ; Extra menu items.
#dDerapserEdge
#dCR : #dLF : #dTAB
EndEnumeration

Enumeration ; Timers.
#tMainInfoTimer
#tBackupTimer
#tSearchTimer
#tNoiseTimer
#tSnapTimer
#tCollector
#tRunTimer
EndEnumeration

Enumeration 667 ; Fonts.
#fInFont
#fAlterFont
#fSBarFont
#fButtonFont
#fSButtonFont
#fTmpListFont
#fListFont
#fBoxedFont
#fOcularFont
EndEnumeration

Enumeration ; Search targets.
#sNone
#sPlainText
#sWSource
#sRemark
#sDictID
#sListLine
EndEnumeration

Enumeration ; Menus
#mListMenu
#mVPMenu
#mSearchMenu
#mTrayMenu
#mDeployment
EndEnumeration
;}
;{ --Structures--
Structure EventData
Window.i
Type.i
SubType.i
Gadget.i
EndStructure

Structure CharExtra
StructureUnion
U.U : Lense.S{1}
EndStructureUnion
EndStructure

Structure UniAsc
StructureUnion : U.U : A.A : EndStructureUnion
EndStructure

Structure COMBOBOXINFO
cbSize.l
rcItem.RECT
rcButton.RECT
stateButton.l
hwndCombo.i
hwndItem.i
hwndList.i
EndStructure

CompilerIf #PB_Compiler_Version < 530 ; In case of outdated (as in, LTS) compiler...
Structure DragDataFormat
Format.i
*Buffer
Size.i
EndStructure
CompilerEndIf

Structure ClipData
HotKey.i   ; Associated hotkey.
DictID.s   ; Identifier in dict.
DataType.i ; Type identifier.
DataSize.i ; Sizing of embedded data.
*BinData   ; Stored data pointer.
MenuSlot.i ; Index of slot in tray menu.
CmpSize.i  ; Compressed data' size.
Comment.s  ; Asocciated remark string.
Flattable.i; Flag for OfferFlatten optimization.
TimeStamp.i; Data registration date and time.
WSource.s  ; Name of window, active at time of WM.
TextData.s   ; Associated text (if any).
Sizing.Point ; To clip metafile sizing.
*ViewPort.ViewPort   ; Associated viewport.
*CacheData.CacheData ; Associated cache entry.
VPSizing.Rect        ; Size params for associated viewport.
StructureUnion       ; ===>Special flags goes here:
RandFlag.i   ; Abstract interface for flag access.
WrapFlag.i   ; Flag for view area being wrapped.
EndStructureUnion
EndStructure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Structure CacheData
TextData.s   ; Cache text.
DataSize.i   ; Size of cached data (in bytes).
Expiration.i ; Time to uncache stored data.
*ClipData.ClipData ; Associated clip entry.
EndStructure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Structure UnFlicker
*Container ; Container gadget to manage clipping outside of window scope.
*Splitter  ; Spiltter gadget to manage sizing operations.
*Voider    ; Pseudo-gadget to setup splitting relativity.
EndStructure

Structure InfoBlock
InfoSum.s   ; Summarized data line.
InfoShift.i ; Shift factor for looped outupt.
*Informer   ; Handle to main informer line.
*InBox      ; Handle to main informer combobox.
EndStructure

Structure Options
ListMax.i    ; Maximal number of entries in list.
UseDump.i    ; Save stored data into dump file (flag).
SaveImages.i ; Save stored images into dump file (flag).
SaveLists.i  ; Save stored images into dump file (flag).
UseTray.i    ; Show tray icon while hidden (flag).
GlassWin.i   ; Transparent window while unfocused.
HideStart.i  ; Hide window on startup (flag).
AltPasting.i ; Use alternative pasting method (flag).
SizeLimit.i  ; Data size limitation (in kilobytes).
MimicList.i  ; Use data list layout to view plain text (flag).
FixPaint.i   ; Auto-fixing for Paint's selection (flag).
HPreserve.i  ; Preserve hotkeyed nodes while cleaning (flag).
FixedTray.i  ; Fix application icon onto tray bar (flag).
WipeWarn.i   ; Show requester to confirm datalist clearing (flag).
AutoReboot.i ; Automatically restart application after err (flag).
SnapWin.i    ; Automatically readjust windows to align (flag).
EndStructure

Structure ViewPort
*WindowID   ; Pointer to associated window.
*ViewArea   ; Pointer to associated gadget.
MimicList.i ; Shows if viewport should mimic data list.
*TempMeta   ; Temporary metafile for callback.
PrevIndex.s ; String representation of former VP' index.
*Frame      ; Artifical frame around view area.
*Img        ; Pointer to associated image, if applicable.
WasVisible.i; Flag for viewport being visible before hiding self.
*WebObject.COMateObject ; For complex browser controlling.
*WinCom.COMateObject    ; For complex browser controlling.
*DataNode.ClipData      ; Pointer to associated node.
Informator.InfoBlock    ; Information summarizing system.
Stabilizer.UnFlicker    ; Output stabilization system.
EndStructure

Structure OptionFlag
IniName.s     ; String identifier for INI file.
Gadget.i      ; Index of corresponding gadget.
DefVal.i      ; Default value for resetting.
*Data.Integer ; Pointer to data on actual Options structure.
EndStructure

Structure SearchEngine
Target.i      ; Type of search being conducted.
CaseFlag.i    ; Flag to force case-sensitive search.
OpenFlag.i    ; Flag to force openinf viewport once target found.
DenialFlag.i  ; Flag to invert results of search.
RegFlag.i     ; Flag for regular expression match.
WholeFlag.i   ; Flag for test text for strict equation to request.
BindingFlag.i ; Flag to filter results by hotkey affinity.
HLFlag.i      ; Flag to highlight found results.
BinMark.i     ; Binary data, being in search.
TextMark.s    ; Text data, being in search.
Direction.i   ; Going up or down, yes.
; ---------------
DenySTR.i     ; Exclusion flag for plain text data
DenyBMP.i     ; Exclusion flag for bitmap data.
DenyDIR.i     ; Exclusion flag for directory listing data.
DenyRTF.i     ; Exclusion flag for rich text data.
DenyMETA.i    ; Exclusion flag for metafile data.
DenyHTML.i    ; Exclusion flag for hyper text.
DenyNot.i     ; Flag for inverting denial filtration.
; ---------------
ToGadget.i       ; Flag to selected gadget after search.
*Anchor.ClipData ; Node, which was under cursor while it's all started.
*Lense.ClipData  ; Node, that is currently under inspection.
EndStructure

Structure SearchFlag
Represent.s   ; Textual flag repesenatation (lcase).
MenuLink.i    ; Link to corresponding.
*Data.Integer ; Pointer to actual flag variable.
FlagVal.i     ; Appointed value.
EndStructure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Structure OpticSys Extends ViewPort
*BitmapArea ; BMP data viewer.
*PlainArea  ; TXT/DIR data viewer.
*MetaArea   ; META data viewer.
*HTMLArea   ; HTML data viewer.
*RTFArea    ; RTF dataviewer
*NoiseArea  ; Pseudo-vewer to show noise.
Actual.i    ; Flag to check actuality of preview window's content.
EndStructure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Structure SystemData
*MainWindow   ; System pointer to main window.
*SetupWindow  ; System pointer to setup window.
*NextWindow   ; System pointer to next window in clipchain.
*OwnHandle    ; Pointer to our thread.
*AppIcon      ; Pointer to app's icon
ClipRTF.i     ; Clipboard format for RTF
ClipHTML.i    ; Clipboard format for HTML
RasterType.i  ; Default format for raster images.
AcceptNew.i   ; Accept new entries (flag).
PostFlag.i    ; New data was posted to clip (flag).
*UpdateMsg    ; Message for viewports to adapt for latest options.
*CloseMsg     ; Message for viewports to close on place.
*HideMsg      ; Message for viewports to hide on screen.
*RenumMsg     ; Message for viewports to renumerate indexes.
*RestoreMsg   ; Message for viewports to restore on screen.
ListToolTip.s ; Current tooltip, dispalyed by data list.
*ListID       ; System pointer to ClipList gadget.
*VoidGadget   ; System pointer to 'void' gadget.
RenderBack.l  ; Background color for rendering metafiles.
Delita.i      ; Factor for wiped out data.
DelState.i    ; Accumulator for deletion position.
WipeShift.i   ; Shift position for cyclical deletion.
DragFlag.i    ; Your best D/D indicator nowadays.
DropFlag.i    ; Your auxilary D/D incdicator, yes.
*RTFParser    ; ID of hidden parser gadget for RTF.
*DupMutex     ; Handle to mutex for checking SC' presence.
*LastSrc      ; Handle of last known clipboard owner.
*LastAnalyzed ; Pointer to last analyzed node.
ForceDrop.i   ; Flag about spacing info menu.
Gravitation.i ; Attraction force between windows.
TotalCached.i ; Total size of cached data (in bytes).
ActualReq.s   ; Actual data, supposed to be in search bar.
*SearchBar    ; Gadget ID for search string field.
HoldField.i   ; Index of field to hold selection on.
XPLegacy.i    ; GUI flag for my beloved OS.
SizeMsg.i     ; Resizer's message.
*ListFont     ; All data for list' font.
Bullet.s{1}   ; OS-dependant character.
SnapShift.i   ; OS-dependant corrector for snapping procedure.
; Former indexators:
*LockedNode.ClipData ; Pointer to data node, locked for dialog.
*UsedNode.ClipDat    ; Pointer to currently selected node.
; Complex data structures:
ClipFormats.i[#AllFormats]     ; Array of supported formats.
CursorPos.Point                ; Cursor position on icon.
Options.Options                ; Application settings.
Panopticum.OpticSys            ; Data preview system.
GUIEvent.EventData             ; Common event accumulator.
Lurker.SearchEngine            ; Data for performing search.
FlickBuf.UnFlicker             ; Accumulator for anti-flicking measures.
Informator.InfoBlock           ; Information summarizing system.
VoidBorders.Rect               ; Null proxy rectangle for initializng new VPorts.
*ViewPort.ViewPort             ; Pointer to current viewport in use.
*WebObject.IWebBrowser2        ; HTML parser interface.
*ComBrowser.COMateObject       ; COM parser auxilary object.
Map *Dictionary.ClipData()     ; Hashtable to disallow copies.
Map PrefixTable.s()            ; Serialized storage for data prefixi.
Map Flagi.OptionFlag()         ; Serialized storage for settings.
Map Emitter.SearchFlag()       ; Serialized list of search flags.
Map *Reflector.SearchFlag()    ; Hashtable to optimize search parsing.
List ClipList.ClipData()       ; List of stored data nodes.
List CacheList.CacheData()     ; List of cached data.
List ViewPorts.ViewPort()      ; List of viewport windows.
List CriticalStack.i()         ; Pseudostack for crit. sections.
LayOut.u[#LayoutTable+1]       ; XLat table for layout switching.
*HotNodes.ClipData[#HotKeys+1] ; Array of associated hotkeys.
Array DragOut.DragDataFormat(0); Array of drag/drop descriptors.
EndStructure
;}
;{ --Variables--
Global I, System.SystemData
;}
;} {End/Definitions}

;{ [Procedures]
;{ @WayMarks@
; Panopticum - окно предпросмотра. Ищется по #cPanopticum
; Обработчики пунктов можно найти по отметке !MenuHandler
; Проверка списка на перегруженность - CheckMaximum()
; (re)storeData - процедуры схоронения и восстановки буфера.
; Сохранение и загрузка дампов - SaveDump/RestoreDump()
; За размерностями главного окна - в GUISizer.
; Для оптимизации Drag\Drop функциональная часть вынесена в макросы Register%format%Guts.
; EncodeData/DecodeData - наши новые друзья.
; ClearData - удаление узла. 
; Feeder информационной строки - GatherNodeInfo.
;}
;{ --Low level--
Macro CheckSize(Size) ; Pseudo-procedure.
(System\Options\SizeLimit = 0 Or Size < System\Options\SizeLimit * #Kb)
EndMacro

Procedure GetClipboardData(Format, MetaOverride = #False) ; Low-level.
Define *Buffer  = GetClipboardData_(Format)
Define DataSize = GlobalSize_(*Buffer)
If DataSize            ; Если есть, что копировать...
Define *DataPtr = GlobalLock_(*Buffer)
Define *Copy    = AllocateMemory(DataSize)
CopyMemory(*DataPtr, *Copy, DataSize)
GlobalUnlock_(*DataPtr) ; Разрешаем использование памяти.
ElseIf MetaOverride And *Buffer ; Спец. вариант для метафайла.
GlobalLock_(*Buffer) : *Copy = CopyEnhMetaFile_(*Buffer, 0) : GlobalUnlock_(*Buffer)
EndIf : CloseClipboard_() ; Закрываем.
ProcedureReturn *Copy
EndProcedure

Procedure WaitOpen(*Window) ; Pseudo-procedure.
While OpenClipboard_(*Window) = #False : Delay(10) : Wend
EndProcedure

Procedure FindClipOwner()
If GetActiveWindow() <> -1 : ProcedureReturn WindowID(GetActiveWindow()) : Else : ProcedureReturn System\MainWindow : EndIf
EndProcedure

Procedure SetClipboardData(*DataPtr, DataSize, Format, Mode = 0) ; Low-level.
If Mode <> 2 ; Если требуется открыть буфер...
WaitOpen(FindClipOwner())                       ; Открываем буфер. 
EmptyClipboard_()                               ; Предварительная очистка.
EndIf
If Mode <> #MetaMode
Define *Buffer  = GlobalAlloc_(#GHND, DataSize) ; Выделяем.
Define *Copy    = GlobalLock_(*Buffer)          ; Блокируем.
CopyMemory(*DataPtr, *Copy, DataSize)           ; Копируем.
GlobalUnlock_(*Buffer)                          ; Разрешаем использование памяти.
SetClipboardData_(Format, *Buffer)              ; Вписываем.
Else : SetClipboardData_(Format, *DataPtr)      ; ...Иначе - тупо пишем, с метафайлами так положено.
EndIf
If Mode <> 1 : CloseClipboard_() : EndIf        ; Если требуется закрыть буффер.
ProcedureReturn *Copy                           ; Возвращаем указатель.
EndProcedure

Macro Rnd(Min, Max) ; Pseudo-procedure.
(Random(Max-Min) + Min)
EndMacro

Procedure XLat(StrPtr, TablePtr)
EnableASM
MOV EAX, StrPtr
MOV EBX, TablePtr
SUB ECX, ECX
MOV CX, word [EAX]
FillOut:
SHL ECX, 1
MOV CX, word [ECX+EBX]
MOV word [EAX], CX
ADD EAX, #CharSize
MOV CX, word [EAX]
Or CX, CX
JNZ ll_xlat_fillout
DisableASM
EndProcedure

Procedure Min(A, B)
If A < B : ProcedureReturn A : Else : ProcedureReturn B : EndIf
EndProcedure

Macro DelayReturn() ; Partializer.
Define __ReturnFlag = #True
EndMacro

Macro ReviseReturn(RetVal =) ; Partializer.
If __ReturnFlag : ProcedureReturn RetVal : EndIf
EndMacro

Procedure EncodeData(*Ptr, Size)
If Size < 20 : ProcedureReturn #False : EndIf ; А вот дабы было не повадно.
Define *OutBuf = AllocateMemory(Size), CSize = CompressMemory(*Ptr, Size, *OutBuf, Size, #PB_PackerPlugin_BriefLZ)
If CSize : *OutBuf = ReAllocateMemory(*OutBuf, CSize) : Else : FreeMemory(*OutBuf) : *OutBuf = #Null : EndIf
ProcedureReturn *OutBuf
EndProcedure

Procedure DecodeData(*Ptr, CmpSize, OriSize, *OutBuf = #Null)
If *OutBuf = #Null : *OutBuf = AllocateMemory(OriSize) : EndIf
Define CSize = UncompressMemory(*Ptr, CmpSize, *OutBuf, OriSize, #PB_PackerPlugin_BriefLZ)
If CSize = #Null : FreeMemory(*OutBuf) : *OutBuf = #Null : EndIf : ProcedureReturn *OutBuf
EndProcedure

Macro OutputSize() ; Partializer.
(DrawingBufferPitch() * OutputHeight())
EndMacro

Procedure.s IIFS(Bool.i, Variant1.s, Variant2.s)
If Bool : ProcedureReturn Variant1 : Else : ProcedureReturn Variant2 : EndIf
EndProcedure

Procedure.s GetWinText(*hWnd)
Define WinText.s = Space(GetWindowTextLength_(*hWnd)) : GetWindowText_(*hWnd, @WinText, Len(WinText) + 1) : ProcedureReturn WinText
EndProcedure

Procedure.s FlattenText(Text.s)
Define *Result = AllocateMemory(StringByteLength(Text) + #CharSize), *In.Unicode = @Text, *Out.Unicode = *Result
While *In\U : If *In\U <> #CR And *in\U <> #LF : *Out\U = *In\U : *Out + SizeOf(UnicodE) : EndIf 
*In + SizeOf(Unicode) : Wend : ProcedureReturn PeekS(*Result)
EndProcedure

Macro HIWORD(Value) ; Pseudo-procedure.
(Value >> 16) & $FFFF
EndMacro

Procedure GageString(*Ptr.Unicode)
Define *Start = *Ptr
While *Ptr\U : *Ptr + SizeOf(Unicode) : Wend
ProcedureReturn (*Ptr - *Start) >> 1 ; Посчитали ? Возвращаем.
EndProcedure

CompilerIf #PB_Compiler_Version => 540 ; Не было печали - апдейтов накачали !
Procedure.l CRC32Fingerprint(*DataPtr, DataSize.i) ; Legacy fingerprint.
ProcedureReturn Val("$" + Fingerprint(*DataPtr, DataSize, #PB_Cipher_CRC32))
EndProcedure

Macro MD5Fingerprint(DataPtr, DataSize) ; Pseudo-procedure.
Fingerprint(DataPtr, DataSize, #PB_Cipher_MD5)
EndMacro
CompilerEndIf
;}
;{ --GUI management--
Procedure Node2Index(*DataNode.ClipData)
If *DataNode : ChangeCurrentElement(System\ClipList(), *DataNode) : ProcedureReturn ListIndex(System\ClipList())
Else : ProcedureReturn #UnusedNode : EndIf
EndProcedure

Procedure Index2Node(Idx)
Define *DN = GetGadgetItemData(#ClipList, Idx)
If *DN <> -1 : ProcedureReturn *DN : EndIf
EndProcedure

Macro UsedNodeIdx() ; Partializer.
Node2Index(System\UsedNode)
EndMacro

Procedure SelectNode(NodeIdx)
Define Idx.i = UsedNodeIdx()
SetGadgetItemText(#ClipList, Idx, Mid(GetGadgetItemText(#ClipList, Idx), Len(#UsedNode) + 1))
SetGadgetItemText(#ClipList, NodeIdx, #UsedNode + GetGadgetItemText(#ClipList, NodeIdx))
System\UsedNode = Index2Node(NodeIdx)   ; Выставляем новое выделение.
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure.s LineCutter(Ln.s)
If Len(Ln) > #MaxLineChars : ProcedureReturn Left(Ln, #MaxLineChars) + "'/[cut]" : EndIf
ProcedureReturn Ln
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure Add2List(Text.s, Link, SelectME = #True)
Define ThisNode = CountGadgetItems(#ClipList)
AddGadgetItem(#ClipList, ThisNode, LineCutter(Text))
SetGadgetItemData(#ClipList, ThisNode, Link)
If SelectME : SelectNode(ThisNode) : EndIf
EndProcedure

Macro UnhideList() ; Partializer.
If IsWindow(#MainWindow) : HideGadget(#ClipList, #False) : EndIf
EndMacro

Macro AtMainNow() ; Pseudo-procedure.
GetActiveWindow() = #MainWindow
EndMacro

Procedure EnterCritical() ; Partializer.
LastElement(System\CriticalStack()) : AddElement(System\CriticalStack())
If ListSize(System\CriticalStack()) : System\CriticalStack() = GetActiveGadget() 
If AtMainNow() : SetActiveGadget(#VoidGadget) : EndIf : EndIf
HideGadget(#ClipList, #True)
EndProcedure

Procedure LeaveCritical() ; Partializer.
UnhideList() : If LastElement(System\CriticalStack())
If ListSize(System\CriticalStack()) = 1 ; Если это последний элемент...
If GetActiveGadget() <> -1 And AtMainNow() : SetActiveGadget(System\CriticalStack()) : EndIf
DeleteElement(System\CriticalStack()) : Else : DeleteElement(System\CriticalStack()) 
EndIf
Else : EndIf
EndProcedure

Procedure ErrorBox(Text.s) ; Pseudo-procedure.
UnhideList() : ProcedureReturn MessageRequester(#Title, "Error: " + Text, #MB_ICONSTOP|#MB_SYSTEMMODAL)
EndProcedure

Procedure WarnBox(Text.s, Buttons = #PB_MessageRequester_YesNo) ; Pseudo-procedure.
UnhideList() : ProcedureReturn MessageRequester(#Title, "Specify: " + Text, #MB_ICONWARNING|#MB_SYSTEMMODAL|Buttons)
EndProcedure

Procedure SwitchStyle(*hWnd, StyleConst, Switch = #True, StyleRank = #GWL_EXSTYLE)
Define Style = GetWindowLongPtr_(*hWnd, StyleRank)
If Switch : Style | StyleConst : Else : Style & ~StyleConst : EndIf
SetWindowLongPtr_(*hWnd, StyleRank, Style)
EndProcedure

Macro AddTray() ; Partializer.
If System\Options\UseTray ; Если трей поддерживается...
#ToolTip = #Title + #CR$ + "-Left click to continue" + #CR$ + "-Right click for menu"
While AddSysTrayIcon(#TrayIcon, System\MainWindow, System\AppIcon) = #False : Delay(100) : Wend
SysTrayIconToolTip(#TrayIcon, #ToolTip) ; Добавляем подсказку.
EndIf
EndMacro

Macro ReturnWindow() ; Pseudo-procedure.
If System\Options\FixedTray=#False:RemoveSysTrayIcon(#TrayIcon):EndIf
HideWindow(#MainWindow, #False) ; Показываем окно.
SendNotifyMessage_(#HWND_BROADCAST, System\RestoreMsg, 0, 0) 
SetForegroundWindow_(System\MainWindow) ; Should be done.
EndMacro

Procedure InstillAlpha(*WindowID, Level)
SetLayeredWindowAttributes_(*WindowID, #Null, Level, #LWA_ALPHA)
EndProcedure

Macro SetOpacity(Value = #FullAlpha, Window = System\MainWindow) ; Pseudo-procedure.
InstillAlpha(Window, Value) : Define *PanID = WindowID(System\Panopticum\WindowID)
If Window = *PanID : InstillAlpha(System\MainWindow, Value) : ElseIf Window = System\MainWindow : InstillAlpha(*PanID, Value) : EndIf
EndMacro

Macro MakeTransparent(Window = System\MainWindow) ; Pseudo-procedure.
If System\Options\Glasswin : SetOpacity(#GlassAlpha, Window) : EndIf
EndMacro

Macro HadleFirstOpen() ; Partializer.
If GetActiveWindow() <> #MainWindow : WinCallback(System\MainWindow, #WM_ACTIVATE, #WA_INACTIVE, 0) : EndIf
EndMacro

Macro PanVisibility() ; Pseudo-procedure.
IsWindowVisible_(WindowID(System\Panopticum\WindowID))
EndMacro

Procedure CCHookProc(hDlg, uiMsg, wParam, *lParam.ChooseColor) ; CallBack.
If uiMsg = #WM_INITDIALOG : SetWindowText_(hDlg, *lParam\lCustData) : EndIf
ProcedureReturn 0
EndProcedure

Procedure ColorRequesterEX(*Owner, StartRGB.l, *TextPtr = #Null) ; Not mine.
Define chc.Choosecolor, RGB.s{16 * 2}
chc\LStructSize = SizeOf(choosecolor) 
chc\hwndOwner = *Owner
chc\rgbResult = startRGB 
chc\lpCustColors = @RGB
chc\lCustData = *TextPtr
chc\lpfnHook = @CCHookProc()
chc\flags = #CC_ANYCOLOR | #CC_RGBINIT | #CC_ENABLEHOOK
If ChooseColor_(@chc)  
ProcedureReturn chc\rgbResult 
Else : ProcedureReturn -1 
EndIf 
EndProcedure

Procedure UpdateSwitch()
DisableDebugger
If System\AcceptNew : SetGadgetText(#Button_Switch, "Stop")
Else                : SetGadgetText(#Button_Switch, "Resume") : EndIf
EnableDebugger
EndProcedure

Procedure UpdateOcular()
DisableDebugger
If PanVisibility() : SetGadgetState(#Button_Ocular, #True)
Else               : SetGadgetState(#Button_Ocular, #False) : EndIf
HideGadget(#EyeFrame, 1 ! PanVisibility())
EnableDebugger
EndProcedure

Macro OverLoad(INum) ; Partializer.
INum > System\Options\ListMax
EndMacro

Procedure UpdateTitle() ; Pseudo-procedure.
Define Title.s, Items = CountGadgetItems(#ClipList)
If Items = 0 : SetWindowTitle(#MainWindow, #Title)
Else : Title = #TitleDecor + #TitleRaw + " [" + Str(Items)
If System\Options\ListMax : If OverLoad(Items) : Title + "!" : EndIf
Title + "/" + Str(System\Options\ListMax) : EndIf
SetWindowTitle(#MainWindow, Title + "]" + #TitleDecor)
EndIf
EndProcedure

Macro ChainOldCB(Window = hWnd, Cnt = "OldProc") ; Partialzier.
CallWindowProc_(GetProp_(Window, Cnt), Window, Message, wParam, lParam)
EndMacro

Macro ChangeCB(GID, CB, Cnt = "OldProc") ; Partializer.
SetProp_(GID, Cnt, SetWindowLongPtr_(GID, #GWL_WNDPROC, @CB))
EndMacro

Procedure ComboWidth(*ComboGadget)
Protected pcbi.COMBOBOXINFO
pcbi\cbSize = SizeOf(COMBOBOXINFO)
GetComboBoxInfo_(GadgetID(*ComboGadget), @pcbi)
ProcedureReturn (pcbi\rcButton\left + pcbi\rcButton\right) * -1
EndProcedure

Procedure ComboBoxButtonGadget(GadgetNr, X, Y)
Define Result = ComboBoxGadget(GadgetNr, X, Y, 0, 0, #PB_ComboBox_Editable)
ResizeGadget(Result, #PB_Ignore, #PB_Ignore, ComboWidth(Result), #PB_Ignore)
ProcedureReturn Result
EndProcedure

Procedure ComboBoxButtonListWidth(WindowNr, GadgetNr, Width=#Null)
Protected Count = CountGadgetItems(GadgetNr) - 1, Index, Text.s, TextWidth, TextWidthMax
If Width <= #Null
StartDrawing(WindowOutput(WindowNr))
DrawingFont(GetGadgetFont(GadgetNr))
For Index = 0 To Count
Text = GetGadgetItemText(GadgetNr, Index)
TextWidth = TextWidth(Text)
If TextWidth > TextWidthMax : TextWidthMax = TextWidth : EndIf
Next Index
StopDrawing()
Width = TextWidthMax
EndIf
Width + (GetSystemMetrics_(#SM_CXEDGE) * 2) + 3 ; Not sure where the 3 is coming from.
If Count > 30 : Width + GetSystemMetrics_(#SM_CXVSCROLL) : EndIf
SendMessage_(GadgetID(GadgetNr), #CB_SETDROPPEDWIDTH, Width, 0)
EndProcedure

Macro ExtractWP(WindowID) ; Partializer.
GetProp_(WindowID, #VProp)
EndMacro

Macro BindWP(WindowID, VPort) ; Partializer.
SetProp_(WindowID, #VProp, VPort)
EndMacro

Macro HLLine(NewIdx) ; Partializer.
SetGadgetState(#ClipList, NewIdx)
EndMacro

Macro DefBarButton(IDx, Text, Font = #fButtonFont, SpecialFlags = #Null) ; Partializer
ButtonGadget(IDx, 0, 0, 0, #ButtonHeight, Text, #BS_FLAT | SpecialFlags) : SetGadgetFont(IDX, FontID(Font))
EndMacro

Procedure BoundCoord(Val, Edge, Shift)
If Val < 0 : Val = 0 : ElseIf Val + Shift > Edge : Val = Edge - Shift : EndIf : ProcedureReturn Val
EndProcedure

Procedure PlaceWindowStrict(*WindowID, X, Y)
ExamineDesktops()
ResizeWindow(*WindowID, BoundCoord(X, DesktopWidth(0), WindowWidth(*WindowID, #PB_Window_FrameCoordinate)), 
                        BoundCoord(Y, DesktopHeight(0), WindowHeight(*WindowID, #PB_Window_FrameCoordinate)), #PB_Ignore, #PB_Ignore)
EndProcedure
;}
;{ --Cache management-
Procedure UncacheNode(*DataNode.ClipData)
If *DataNode\CacheData ; Если есть вообще, о чем говорить...
With *DataNode\CacheData ; Обрабатываем, коли уж.
System\TotalCached - \DataSize ; Уменьшаем счетчик заветного места.
\TextData = ""                 ; А вот дабы неповадно было.
ChangeCurrentElement(System\CacheList(), *DataNode\CacheData)
DeleteElement(System\CacheList()) : *DataNode\CacheData = #Null
EndWith
EndIf
EndProcedure

Procedure CacheNodeText(*DataNode.ClipData, Text.s)
If *DataNode\CacheData : UnCacheNode(*DataNode) : EndIf ; Сразу выгружаем, значит.
If Text ; Только в том случае, когда там есть, что кешировать...
LastElement(System\Cachelist()) : *DataNode\CacheData = AddElement(System\CacheList())
With *DataNode\CacheData ; Обрабатываем.
\TextData = Text : \DataSize = StringByteLength(Text)                        ; Записываем текст.
\Expiration = AddDate(Date(), #PB_Date_Minute, #CacheQuant * #QuantMul)      ; Теряет актуальность через 10 минут.
\ClipData = *DataNode : System\TotalCached + \DataSize                       ; ...И все прочее...
EndWith : ProcedureReturn *DataNode\CacheData : EndIf                        ; Возвращаем указатель на тесты.
EndProcedure

Macro GrantQuant(TStart, Mul = 1) ; Pseudo-procedure.
AddDate(TStart, #PB_Date_Minute, #CacheQuant * Mul) 
EndMacro

Procedure.s RequestCachedText(*DataNode.ClipData)
If *DataNode\CacheData ; Если в кеше вообще что-то есть...
Define Prolong = GrantQuant(*DataNode\CacheData\Expiration), Extend = GrantQuant(Date(), #QuantMul) ; Нельзя отложить > на 10 минут.
If Prolong > Extend : *DataNode\CacheData\Expiration = Extend : Else : *DataNode\CacheData\Expiration = Prolong : EndIf
ProcedureReturn *DataNode\CacheData\TextData : EndIf ; Ну и возвращаем итог.
EndProcedure

Procedure Expires(*DataNode.ClipData)
If *DataNode\CacheData And *DataNode\Hotkey = 0 ; Все, что на горячих клавишах - должно выдавать максимально быстро.
ProcedureReturn *DataNode\CacheData\Expiration : EndIf
EndProcedure

Procedure GCollection()
ForEach System\CacheList()
Define *TCache.CacheData = System\CacheList(), *DataNode.ClipData = *TCache\ClipData, ETimer = Expires(*DataNode)
If ETimer <= Date() : Uncachenode(*DataNode) : EndIf ; Так и выкидываем, дабы знали.
Next
EndProcedure
;}
;{ --Clipboard management--
Procedure Shiftae(Flag)
If Flag = #False : ProcedureReturn #ForceMode : Else : ProcedureReturn #False : EndIf
EndProcedure

Procedure UnloadImage(*DN.ClipData)
With *DN
If \CmpSize : FreeMemory(\BinData)
Else        : FreeImage(\BinData) : EndIf
EndWith
EndProcedure

Macro UnloadText(DN) ; Partializer.
If DN\BinData : FreeMemory(DN\BinData) : EndIf
EndMacro

Macro UnloadFiles(DN) ; Partializer.
If DN\CmpSize : FreeMemory(DN\BinData) : EndIf
EndMacro

Macro UnloadMeta(DN) ; Partializer.
If DN\CmpSize : FreeMemory(DN\BinData)
Else          : DeleteEnhMetaFile_(DN\BinData) : EndIf
EndMacro

Procedure DisposeBinary(*DataNode.ClipData)
DisableDebugger 
Select *DataNode\DataType
Case #CF_RichText, #CF_HTML : FreeMemory(*DataNode\BinData)
Case #CF_BITMAP      : UnloadImage(*DataNode)
Case #CF_ENHMETAFILE : UnloadMeta(*DataNode)
Case #CF_TEXT        : UnloadText(*DataNode)
Case #CF_HDROP       : UnloadFiles(*DataNode)
EndSelect            : UncacheNode(*DataNode)
EnableDebugger
EndProcedure
; -------------------------------
Macro Pref(Type) : System\PrefixTable(Hex(Type)) : EndMacro ; One-liner.
; -------------------------------
Procedure UncmpImage(*DN.ClipData)
With *DN
Define Img = CreateImage(#PB_Any, \Sizing\X, \Sizing\Y, #BitDepth)
StartDrawing(ImageOutput(Img))
Define *PixBuf = DecodeData(\BinData, \CmpSize, \DataSize, DrawingBuffer())
StopDrawing()
EndWith
ProcedureReturn Img
EndProcedure

Procedure ExtractImage(*DN.ClipData)
If *DN\CmpSize : ProcedureReturn UncmpImage(*DN) : Else : ProcedureReturn *DN\BinData : EndIf
EndProcedure

Macro CleanUpImage(DN, ImgFeeder) ; Partializer.
If DN\CmpSize And ImgFeeder : FreeImage(ImgFeeder) : ImgFeeder = #Null : EndIf
EndMacro

Macro ImgSelect() ; Partializer.
Define ImgIDx = ExtractImage(*DataNode)
EndMacro

Macro ImgCheckUP() ; Partializer.
CleanUpImage(*DataNode, ImgIDx)
EndMacro

Procedure UncmpMeta(*DataNode.ClipData)
With *DataNode
Define *Buffer = DecodeData(\BinData, \CmpSize, \DataSize)
Define *Meta   = SetEnhMetaFileBits_(\DataSize, *Buffer)
FreeMemory(*Buffer) : ProcedureReturn *Meta
EndWith
EndProcedure

Procedure ExtractMeta(*DN.ClipData)
If *DN\CmpSize : ProcedureReturn UncmpMeta(*DN) : Else : ProcedureReturn *DN\BinData : EndIf
EndProcedure

Macro CleanUpMeta(DN, MetaFeeder) ; Partializer.
If DN\CmpSize And MetaFeeder : DeleteEnhMetaFile_(MetaFeeder) : MetaFeeder = #Null : EndIf
EndMacro

Macro MetaSelect() ; Partializer.
Define *MetaIDx = ExtractMeta(*DataNode)
EndMacro

Macro MetaCheckUP() ; Partializer.
CleanUpMeta(*DataNode, *MetaIDx)
EndMacro

Macro SAG(GadgetID = #ClipList) ; Partializer.
SetActiveGadget(GadgetID)
EndMacro

Macro RenumSignal(NodeIdx) ; Partializer.
SendNotifyMessage_(#HWND_BROADCAST, System\RenumMsg, 0, 0)
EndMacro

Macro SearchStall() ; Pseudo-procedure
System\Lurker\Target = #sNone
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro HandleSearchPtr(SPtr) ; Pseudo-procedure.
If *DataNode = SPtr And Not SearchStall() ; Если удалили обрабатываемый узел...
If ListSize(System\ClipList())
ChangeCurrentElement(System\ClipList(), SPtr) : If PreviousElement(System\ClipList()) = #Null ; Назад.
LastElement(System\ClipList()) : EndIf        : SPtr = System\ClipList()                      ; Переставляем указатель.
Else : System\Lurker\Target = #sNone : EndIf ; Обрываем поиск.
EndIf
EndMacro 

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure ExtractHK(Index)
Define *DN.ClipData = Index2Node(Index)
If *DN : ProcedureReturn *DN\HotKey : EndIf
EndProcedure

Macro WipeAt(Index) ; Partializer.
System\WipeShift = Index
EndMacro

Procedure FindHost(*Host.ClipData, SelectIt = #True)
If SelectIT = #ForceMode Or System\UsedNode <> *Host ; Если так и надо, вновь искать...
I = Node2Index(*Host) ; Вот это - оптимизация, это понимаю. 
If SelectIT <> #ForceMode  : SelectNode(I) : EndIf : If SelectIt : HLLine(I) : EndIf
EndIf : ProcedureReturn I
EndProcedure

Procedure CheckHost(ID.s)
ProcedureReturn System\Dictionary(ID)
EndProcedure

Procedure LinkIDs(*DataNode.ClipData, ID.s)
System\Dictionary(ID) = *DataNode
*DataNode\DictID = ID
EndProcedure

Procedure.s ExtractSourceName(*Variant)
Define FTry.s = GetWinText(*Variant) 
If FTry : ProcedureReturn FTry : EndIf ; Ну так, вдруг...
ProcedureReturn GetWinText(GetForegroundWindow_()) 
EndProcedure

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro CleanAfter(DN, Ptr) ; Pseudo-procedure.
If DN\CmpSize : FreeMemory(Ptr) : EndIf
EndMacro

Procedure ExtractComplexData(*DN.ClipData)
With *DN
If \CmpSize : ProcedureReturn DecodeData(\BinData, \CmpSize, \DataSize) : Else : ProcedureReturn \BinData : EndIf
EndWith
EndProcedure
; ------------------
Macro CheckFieldPtr(Pointer) ; Partializer.
If Pointer = #Null : ProcedureReturn #Null : EndIf
EndMacro

Macro ErrorHTML() ; Partializer.
"<html><body>Incorrect header specified.</body></html>"
EndMacro

Macro CheckFragmentPtr(Pointer) ; Partializer.
If Pointer <= #Null : ProcedureReturn ErrorHTML() : EndIf
EndMacro

Macro OutBound() ; Partializer.
(*SafetyBound And *FragStart + *FragEnd > *SafetyBound)
EndMacro

Procedure SeekCRLF(*Source.Ascii)
While *Source\A : If *Source\A = #CR Or *Source\A = #LF : ProcedureReturn *Source : EndIf
*Source + SizeOf(Ascii) : Wend ; Поиск символа-разделителя.
EndProcedure

Procedure FindMark(*Source.Ascii, Mark.s)
Define *MarkPtr.Unicode = @Mark, *Anchor = @Mark     ; Устанавливаем позиции.
While *Source\A : If *Source\A = *MarkPtr\U          ; Если символы совпадают...
*MarkPtr + #CharSize : If *MarkPtr\U = #Null         ; ...Если они сопадают полностью...
ProcedureReturn *Source + SizeOf(Ascii) : EndIf      ; Возвращаем позицию.
Else : *Source - (*MarkPtr - *Anchor) >> 1           ; Иначе - сдвигаем все обратно.
*MarkPtr = *Anchor : EndIf : *Source + SizeOf(Ascii) : Wend ; Продолжаем поиск.
EndProcedure

Procedure EvalHeaderField(*Src, FieldName.s)
Define *ValStart = FindMark(*Src, FieldName) : CheckFieldPtr(*ValStart)  ; Парсим поле до начала данных.
Define *ValFinish = SeekCRLF(*ValStart)      : CheckFieldPtr(*ValFinish) ; Парсим окончание данных.
If *ValStart < *ValFinish : ProcedureReturn Val(PeekS(*ValStart, *ValFinish - *ValStart, #PB_Ascii)) : EndIf
EndProcedure

Procedure.s FindBody(*Source.Ascii)
Define *SafetyBound = MemorySize(*Source)
If *SafetyBound : *SafetyBound + *Source : EndIf
Define *FragStart = EvalHeaderField(*Source, "StartHTML:")
Define *FragEnd   = EvalHeaderField(*Source, "EndHTML:")  
If *FragStart < 0 Or *FragEnd <= *FragStart Or OutBound() ; Проверяем, все ли хорошо и нельзя ли этим кончить...
*FragStart = EvalHeaderField(*Source, "StartFragment:") : CheckFragmentPtr(*FragStart)
*FragEnd   = EvalHeaderField(*Source, "EndFragment:")   : CheckFragmentPtr(*FragEnd)
If *FragEnd < *FragStart Or *FragStart < *Source Or OutBound() ; Проверяем размерности области чтения.
ProcedureReturn ErrorHTML() : EndIf              ; Возвращаем, во избежание.
EndIf : ProcedureReturn PeekS(*Source + *FragStart, *FragEnd - *FragStart, #PB_UTF8)
EndProcedure

Macro SetHTMLContent(Web, HTML) ; Partializer.
Web\SetProperty("Document\Body\innerHTML = '" + ReplaceString(HTML, "'", "&#39;") + "'")
EndMacro

Macro ParseHTML(Web, Feed) ; Partializer.
SetHTMLContent(Web, FindBody(Feed))
EndMacro
; ------------------

Macro UpholdGRsr() ; Partializer.
If *GRsr = #Null : *GRsr = ExtractComplexData(*DN) : Else : Define Deliberate = #True : EndIf
EndMacro

Macro PeekRTF(DN, Ptr) ; Pseudo-procedure.
PeekS(Ptr, DN\DataSize >> 1)
EndMacro

Procedure DisposeVPData(*VPort.ViewPort)
With *VPort  ; Обрабатываем данные порта вывода.
If \DataNode ; Если прилинкован какой-то нод...
Select \DataNode\DataType ; По типу...
Case #CF_TEXT, #CF_HDROP, #CF_RichText : SetGadgetText(\ViewArea, "")
Case #CF_ENHMETAFILE      : CleanUpMeta(*VPort\DataNode, *VPort\TempMeta)
Case #CF_BITMAP           : CleanUpImage(*VPort\DataNode, *VPort\Img) 
Case #CF_HTML : SetHTMLContent(\WebObject, "") : If \WinCom : \WinCom\Release() : \WinCom = #Null : EndIf
EndSelect : EndIf
EndWith
EndProcedure

Macro WebDisposal(VP = *VPort) ; Partializer.
If VP\WebObject : VP\WebObject\Release() : VP\WebObject = #Null : EndIf 
EndMacro

Procedure CloseViewPort(*VPort.ViewPort) ; Does not belong here, but...
HideWindow(*VPort\WindowID, #True) : DisposeVPData(*VPort) : WebDisposal() ; Скрываем и удаляем данные.
With *VPort\DataNode\VPSizing         ; Сжохраняем данные местоположения, просто на всякий случай.
\Left  = WindowX(*VPort\WindowID)     : \Top    = WindowY(*VPort\WindowID)
\Right = WindowWidth(*VPort\WindowID) : \Bottom = WindowHeight(*VPort\WindowID)
EndWith ; Теперь все закрываем и все уничтожаем из GUI:
CloseWindow(*VPort\WindowID) : *VPort\DataNode\ViewPort = #Null : *VPort\WindowID = #Null 
ChangeCurrentElement(System\ViewPorts(), *Vport) : DeleteElement(System\ViewPorts())
SetForegroundWindow_(System\MainWindow) : SAG()
EndProcedure

Procedure.s ExtractTextGuts(*DN.ClipData, *GRsr = #Null)
With *DN
Select \DataType  ; Выбираем по типу данных.
Case #CF_RichText : UpholdGRsr() : SetGadgetText(#RTFParser, PeekRTF(*DN, *GRsr))  ; Rich text format.
If Not Deliberate : CleanAfter(*DN, *GRsr) : EndIf : ProcedureReturn GetGadgetText(#RTFParser)
Case #CF_HTML     : UpholdGRsr() : ParseHTML(System\ComBrowser, *GRsr) ; Hyper-Text Markup language.
If Not Deliberate : CleanAfter(*DN, *GRsr) : EndIf : ProcedureReturn System\ComBrowser\GetStringProperty("Document\Body\innerText")
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Default      ; For all of those other types.
If \CmpSize : Define Result.s = Space(\Sizing\X)
DecodeData(\BinData, \CmpSize, \DataSize, @Result)
ProcedureReturn Result : Else : ProcedureReturn \TextData
EndIf
EndSelect
EndWith
EndProcedure

Procedure.s ExtractVPText(*VP.ViewPort)
Select *VP\DataNode\DataType ; Выбор по типу.
Case #CF_HTML ; Извлекаем так.
ProcedureReturn *VP\WebObject\GetStringProperty("Document\Body\innerText")
Default : ProcedureReturn GetGadgetText(*VP\ViewArea) ; Стандартное извлечение.
EndSelect
EndProcedure
; ------------------------
Procedure IsTextNode(*DN.ClipData, CheckHTML = #False)
Select *DN\DataType : Case #CF_TEXT, #CF_RichText, #CF_HDROP : ProcedureReturn #True 
											Case #CF_HTML : ProcedureReturn CheckHTML
EndSelect
EndProcedure
; ------------------------
Procedure.s ExtractText(*DN.ClipData, *GRsr = #Null)
If IsTextNode(*DN, #True) = #False : ProcedureReturn "" : EndIf ; Заглушка.
Define Plain.s = RequestCachedText(*DN)
If Plain = "" ; Если в кеше ничего нет, либо там пустая строка...
If *DN\ViewPort : Plain = ExtractVPText(*DN\ViewPort) ; Извлекаем из области просмотра.
ElseIf *DN = System\Panopticum\DataNode And System\Panopticum\Actual : Plain = ExtractVPText(System\Panopticum)
Else : Plain = ExtractTextGuts(*DN, *GRsr) : SetGadgetText(#HTMLPArser, "") : SetGadgetText(#RTFParser, "") : EndIf
CacheNodeText(*DN, Plain) : EndIf ; Ну и флаг сжимаемости, заодно:
If Plain : *DN\Flattable = #True : EndIf
ProcedureReturn Plain
EndProcedure

Macro FreeCompression() ; Partializer
If *CMem : FreeMemory(*CMem) : EndIf
EndMacro

Procedure ClearData(NodeIdx, Preserve = #False) ; Legacy way.
If NodeIdx <> #UnusedNode ; Если есть, что удалять...
Define *DataNode.ClipData = Index2Node(NodeIdx)
If *DataNode = #Null : ProcedureReturn #True : EndIf                        
ChangeCurrentElement(System\ClipList(), *DataNode)                           ; И заодно меняем ставим элемент.
; -------------------------------
If Preserve  ; Если включена защита забинденных нодов...
Define Flag.i, LifeMode = CountGadgetItems(#ClipList) - 1
While *DataNode\Hotkey : Flag = #True                                        ; Ищем первый свободный.
If NextElement(System\ClipList()) = #Null : ProcedureReturn #True : EndIf    ; Сразу выбиваемся.
*DataNode = System\ClipList() : NodeIdx + 1                                  ; Сдвигаем.
If NodeIdx = LifeMode : LifeMode = #NoLife : EndIf                           ; На случай, если можно обойтись без жертв. 
Wend : If Flag : System\WipeShift = NodeIdx : ProcedureReturn #False : EndIf ; Ну, будем считать, что оно - оптимизация. 
EndIf ; Продолжаем удаление:
; -------------------------------
DisposeBinary(*DataNode)             ; Удаляем бинарные данные.
If *DataNode\HotKey : System\HotNodes[*DataNode\HotKey] = #Null : EndIf
DeleteMapElement(System\Dictionary(), *DataNode\DictID)
DeleteElement(System\ClipList())     ; Удаляем элемент.
If System\LockedNode = *DataNode : System\LockedNode = #Null : EndIf           ; Уточняем, что узел более не валиден.
HandleSearchPtr(System\Lurker\Anchor) : HandleSearchPtr(System\Lurker\Lense)   ; Обрабатываем, наконец.
; Проверка выделения...
If *DataNode = System\UsedNode     : SelectNode(#UnusedNode)                   ; Снимаем выделение.
If LifeMode <> #NoLife : ClearClipboard() : SAG() : EndIf                      ; Жизненно важное условие.
EndIf : RemoveGadgetItem(#ClipList, NodeIdx)                                   ; Удаляем запись. Здесь из-за особенностей.
; Последние штришки...
If *DataNode = System\Delita : System\Delita = #Null : EndIf                   ; Корректор при очистке от излишних элементов.
If NextElement(System\ClipList())  : HLLine(Node2Index(System\ClipList())) ; Смена текущего выделения у листа.
ElseIf CountGadgetItems(#ClipList) : LastElement(System\ClipList()) : HLLine(Node2Index(System\ClipList()))
Else : HLLine(#UnusedNode) : SetActiveGadget(-1) : EndIf ; Выделяем ничего, о да.
If *DataNode\ViewPort : CloseViewPort(*DataNode\ViewPort) : EndIf
RenumSignal(NodeIdx)  : UpdateTitle() : EndIf
EndProcedure

Procedure CheckMaximum()
Define Iteration = 1, UNI = UsedNodeIdx(), LI = Node2Index(System\LockedNode), Sel = GetGadgetState(#ClipList) ; Фактор-индексы.
System\WipeShift = 0                                                     ; Выставляем в ноль, вот да.
If System\Options\ListMax > 0                                            ; Если вообще имеет смысл приступать к работе.
While ListSize(System\ClipList()) > System\Options\ListMax               ; Проверям список, пока он больше, нежели хотелось бы.
While Iteration = 1 And (WipeAt(UNI) Or WipeAt(LI) Or WipeAt(Sel) Or ExtractHK(System\WipeShift)) : System\WipeShift + 1 : Wend
If ClearData(System\WipeShift, System\Options\HPreserve) : If Iteration = 1 : Iteration  + 1 : System\WipeShift = 0 ; Защита.
Else : Break : EndIf : EndIf                                             ; Иначе избегаем бесконечности.
Wend
EndIf
EndProcedure

Procedure RegisterNode(ListText.s, Hash.s, SelectNew = #True) ; Partializer.
LastElement(System\ClipList()) : AddElement(System\ClipList())
Define *DataNode.ClipData = System\ClipList()
*DataNode\TimeStamp     = Date() 
*DataNode\DictID        = Hash
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
*DataNode\WSource = EXtractSourceNAme(System\LastSrc)
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Add2List(ListText, System\ClipList(), SelectNew)
LinkIDs(*DataNode, Hash)
CheckMaximum() : UpdateTitle()
ProcedureReturn *DataNode
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Macro TextEntry(Text) ; Pseudo-procedure.
Pref(#CF_TEXT) + "'" + Text + "'"
EndMacro

Macro TextMD5(Text) ; Pseudo-procedure.
Pref(#CF_TEXT) + MD5Fingerprint(@Text, StringByteLength(Text))
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro ImageCRC() ; Partializer.
Hex(CRC32Fingerprint(DrawingBuffer(), OutputSize()))
EndMacro

Macro ImageMD5() ; Partializer.
Pref(#CF_BITMAP) + MD5Fingerprint(DrawingBuffer(), OutputSize())
EndMacro

Macro ImageSizing() ; Partializer.
Str(OutputWidth()) + "x" + Str(OutputHeight())
EndMacro

Macro ImageText(Sizing = ImageSizing(), CRC32 = ImageCRC()) ; Partializer.
Pref(#CF_BITMAP) + Sizing + ", CRC32 = " + CRC32
EndMacro

Macro ComplexText2CRC(DataPtr, Size, Prefix) ; Partializer.
(Prefix + Hex(CRC32Fingerprint(DataPtr, Size)))
EndMacro

Macro ComplexText2MD5(DataPtr, Size, Prefix) ; Partializer.
(Prefix + MD5Fingerprint(DataPtr, Size))
EndMacro

Macro CompoundText(Text) ; Partializer.
", text = '" + Text + "'"
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro FilesID(FList) ; Partializer
(Pref(#CF_HDROP) + MD5Fingerprint(@FList, StringByteLength(FList)))
EndMacro

Procedure.s FormatEntry(Number) ; Partializer
If Number = 1 : ProcedureReturn "entry"
Else          : ProcedureReturn "entries"
EndIf
EndProcedure

Macro FilesText(Files, FList) ; Partializer
(Pref(#CF_HDROP) + Str(Files) + " " + formatEntry(Files) + ", list = '" + CompressListing(@FList) + "'")
EndMacro

Macro DemiLShift(PField) ; Partializer.
*Seeker + SizeOf(*Seeker\PField)
EndMacro

Macro LResultShift() ; Partializer.
*RePtr + SizeOf(Unicode)
EndMacro

Macro FListShift(Pfield) ; Partializer.
DemiLShift(PField) : LResultShift()
EndMacro
; ------------------------------------------------------
Macro ListParser(PField) ; Partializer.
Define Reparse.s = Space((*ListSize >> (SizeOf(*Seeker\Pfield) - 1)) - #ListingExtra)
Define *RePtr.Unicode = @Reparse ; Указатель для формирования результата.
Repeat : If *Seeker\PField : *RePtr\U = *Seeker\PField : FListShift(PField)   ; Если есть какой символ - копируем.
Else   : *RePtr\U = #DirDragDelim : *CountPtr\I + 1    : DemiLShift(PField)   ; Если же символа нет, то инкрементируем.
If *Seeker\PField = #Null : *RePtr\U = #Null : Break : EndIf : LResultShift() ; Сразу выходим, двойной Null же.
EndIf  : ForEver
EndMacro
; ------------------------------------------------------
Procedure.s ReparseFiles(*ListPtr.DROPFILES, *CountPtr.Integer)
If *ListPtr ; Если есть вообще о чем говорить...
Define FileList, *Seeker.UniAsc = *ListPtr + *ListPtr\pFiles ; Готовим указатель для обработки.
Define *ListSize = MemorySize(*ListPtr) - *ListPtr\pFiles    ; Считаем, собственно, размер итогов.
If *ListPtr\fWide : ListParser(U) : *ListPtr\fWide = #True
             Else : ListParser(A) : EndIf : ProcedureReturn Reparse
EndIf
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure SeekListingDelim(*StartPtr.Unicode)
While *StartPtr\U ; Ищем до нуля на случай поломанной строки.
If *StartPtr\U = #DirDragDelim : ProcedureReturn *StartPtr : EndIf ; Возвращаем, где у нас разделитель.
*StartPtr + SizeOf(Unicode)    : Wend : ProcedureReturn *StartPtr  ; Одной головной болью меньше. Возвращаем всегда.
EndProcedure

Procedure.s DirCombine(Listing.S, NewEntry.s)
If Listing : ProcedureReturn Listing + #DirDelim + NewEntry : Else : ProcedureReturn NewEntry : EndIf
EndProcedure

Procedure.s CompressListing(*ListingPtr)
Define CPression.s, MainPath.s
Repeat ; Пока не пришли к Null'у.
Define *DChar.Unicode = SeekListingDelim(*ListingPtr), SSize = (*DChar - *ListingPtr) >> 1
If SSize = 0 : Break : EndIf ; Сразу выходим, т.к. формат передачи листингов не предусматривает пустых полей.
Define Entry.s = PeekS(*ListingPtr, SSize), EPath.s = GetPathPart(Entry)             ; Получаем содержимое поля и его адресную часть.
If EPath <> MainPath.s : CPression = DirCombine(CPression, Entry) : MainPath = EPath ; Ставим как новый основной путь.
Else : CPression = DirCombine(CPression, GetFilePart(Entry)) : EndIf                 ; Иначем - просто дописываем имя к списку.
If Len(CPression) > #MaxLineChars Or *DChar\U = #Null : Break : EndIf                ; Выходим, если для описания получается многовато.
*ListingPtr = *DChar + SizeOf(Unicode) : ForEver : ProcedureReturn CPression
EndProcedure

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure.s MetaText(*Header.ENHMETAHEADER, CRC32.s, *Size.Point) ; Partializer.
*Size\X = (*Header\rclBounds\Right-*Header\rclBounds\Left)
*Size\Y = (*Header\rclBounds\Bottom-*Header\rclBounds\Top)
If *Size\X > 0 And *Size\Y > 0 ; Если есть минимальные рамки...
Define Bwidth.s = Str(*Size\X), BHeight.s = Str(*Size\Y) : Else : Bwidth = "0" : BHeight = "0" : EndIf
ProcedureReturn Pref(#CF_ENHMETAFILE) + BWidth + "x" + BHeight + "x" + Str(*Header\nRecords) + ", CRC32 = " + CRC32
EndProcedure

Macro MetaCRC(Ptr, Size) ; Partializer.
Hex(CRC32Fingerprint(Ptr, Size))
EndMacro

Macro MetaMD5(Ptr, Size)) ; Partializer.
Pref(#CF_ENHMETAFILE) + MD5Fingerprint(Ptr, Size)
EndMacro

Procedure AnalyzeMeta(*Meta, *Header.ENHMETAHEADER)
With *Header ; Обрабатываем заголовок.
If (\nRecords = 13 Or \nRecords = 14) And \nHandles = 1 And \nPalEntries = 0
If IsClipboardFormatAvailable_(#CF_BITMAP)        ; Если есть, чем заменить...
ProcedureReturn #True                             ; Значит берем изображение.
EndIf
EndIf
EndWith
EndProcedure

Macro GetMetaHeader(EMF) ; Partializer.
Define Header.ENHMETAHEADER
GetEnhMetaFileHeader_(EMF, SizeOf(Header), @Header) ; Экстракция заголовка.
EndMacro

Macro GetMetaBits(EMF) ; Partializer.
Define *RawData = AllocateMemory(Header\nBytes)   ; Аллокация буфера для хранения итооговых данных.
GetEnhMetaFileBits_(EMF, Header\nBytes, *RawData) ; Получение итоговых данных.
EndMacro

Macro CommonMD5(Prefix, Ptr, Size) ; Partializer.
Prefix + MD5Fingerprint(Ptr, Size)
EndMacro

Macro CommonSize(DataSize)
", size = " + StrD(DataSize / #KB, 1) + "Kb"
EndMacro

Macro PrepareDelita() ; Partializer. 
System\Delita = Index2Node(GetGadgetState(#ClipList))
EndMacro

Macro UseDelita() ; Patializer.
If System\Delita : HLLine(Node2Index(System\Delita)) ; Просто выставляем тот самый элемент.
Else : HLLine(-1) : SetActiveGadget(-1) : EndIf
EndMacro

Procedure CountLDelims(*ListPtr.Unicode)
While *ListPtr\U : If *ListPtr\U = #DirDragDelim : Define Ctr = Ctr + 1 : EndIf : *ListPtr + SizeOf(Unicode) : Wend
ProcedureReturn Ctr
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro RegisterTextGuts(Flag = #True) ; Partializer 2.0
If Text ; Если есть какой-нибудь текст...
Define ID.s = TextMD5(Text)      ; Идентификатор.
Define Entry.s = TextEntry(Text) ; Форматирование.
Define *Host.ClipData = CheckHost(ID)
If *Host = #Null ; Если такого вхождения еще не было...
Define TrueSize = StringByteLength(Text), CS = TrueSize
Define *CMem  = EncodeData(@Text, TrueSize) ; Пробуем жать.
If *CMem : CS = MemorySize(*CMem) : EndIf   ; Корректор.
; --------------
If CheckSize(CS) ; Проверям размер. 
Define *DataNode.ClipData = RegisterNode(Entry, ID, Flag)
*DataNode\DataType = #CF_TEXT
*DataNode\Sizing\X = Len(Text)
*DataNode\DataSize = TrueSize
If *CMem : *DataNode\BinData  = *CMem
           *DataNode\CmpSize  = CS
Else     : *DataNode\TextData = Text
EndIf
Else#FreeCompression() ; Обязательно высвобождаем.
; --------------
Else : FindHost(*Host, Shiftae(Flag))
EndIf
EndIf
EndMacro

Macro RegisterText() ; Partializer
Define Text.s = GetClipboardText()
RegisterTextGuts() ; Да будет так. Везде и всегда
EndMacro

Macro RegisterImageGuts(Flag = #True) ; Partializer 2.0
If Image ; Если вообще есть, о чем говорить...
StartDrawing(ImageOutput(Image)) ; Открываем поверхность.
Define ImgSize.s = ImageSizing(), CRC32.s = ImageCRC(), MD5.s = ImageMD5()
Define *Host.ClipData = CheckHost(MD5)
If *Host = #Null ; Если такого вхождения еще не было...
Define TrueSize = OutputSize(), CS = TrueSize
Define *CMem  = EncodeData(DrawingBuffer(), TrueSize) ; Пробуем жать.
If *CMem : CS = MemorySize(*Cmem) : EndIf ; Корректор.
If CheckSize(TrueSize)           ; Проверяем размерность.
; ---------
Define *DataNode.ClipData = RegisterNode(ImageText(ImgSize, CRC32), MD5, Flag)
*DataNode\DataType = #CF_BITMAP
*DataNode\Sizing\X = ImageWidth(Image)
*DataNode\Sizing\Y = ImageHeight(Image)
*DataNode\DataSize = TrueSize
If *CMem : *DataNode\BinData = *CMem
           *DataNode\CmpSize = CS
           StopDrawing() : FreeImage(Image)
Else     : *DataNode\BinData = Image : EndIf
; ---------
Else : FreeCompression() : FreeImage(Image) : EndIf ; Обязательно высвобождаем.
Else : FindHost(*Host, Shiftae(Flag)) : FreeImage(Image)
EndIf
EndIf : StopDrawing() ; Неизбежное, вот да.
EndMacro

Procedure RegisterImage() ; Partializer
Define Image = GetClipboardImage(#PB_Any, #BitDepth)
RegisterImageGuts(Image)
EndProcedure

Macro RegisterFilesGuts(Flag = #True) ; Partializer 2.0
If Counter                                                    ; Если в списке есть хоть один пусть...
; ------------------------------------
Define Hash.s   = FilesID(FList)                              ; Получаем хэш.
Define *Host.ClipData = CheckHost(Hash)                       ; Проверяем уникальность.
If *Host = #Null                                              ; Если такого вхождения еще не было...
Define TrueSize = StringByteLength(FList), CS = TrueSize      ; Считаем размер.
Define *CMem    = EncodeData(@FList, TrueSize)                ; Пробуем жать.
If *CMem : CS   = MemorySize(*CMem) : EndIf                   ; Корректор.
If CheckSize(CS)                                              ; Проверяем размерность.
Define *DataNode.ClipData = RegisterNode(FilesText(Counter, FList), Hash, Flag) ; Регистрируем вхождение.
*DataNode\DataType = #CF_HDROP  : *DataNode\DataSize = TrueSize ; Формат и размерность.
*DataNode\Sizing\X = Len(FList) : *DataNode\Sizing\Y = Counter  ; Длина и количество файлов.
If *CMem : *DataNode\BinData  = *CMem : *DataNode\CmpSize = CS
    Else : *DataNode\TextData = FList
EndIf : *DataNode\Flattable = #True    ; Сжимаемо. Гарантия.
Else  : FreeCompression() : EndIf
Else  : FindHost(*Host, Shiftae(Flag)) ; Обязательно.
; ------------------------------------
EndIf : EndIf
EndMacro

Macro RegisterFiles() ; Partializer.
Define Counter.i, FList.S, *ListPtr = GetClipboardData(#CF_HDROP)
If *ListPtr : Flist = ReparseFiles(*ListPtr, @Counter) : FreeMemory(*ListPtr)
RegisterFilesGuts() : EndIf
EndMacro

Macro RegisterMetafileGuts(Flag = #True) ; Partializer 2.0
Define MetaSize.point : GetMetaHeader(*Meta)      ; Заголовок метафайла.
If System\Options\FixPaint = #False Or AnalyzeMeta(*Meta, Header) = #False ; Фикс.
If Header\nBytes : GetMetaBits(*Meta)             ; Если данные в норме...
Define Hash.s = MetaMD5(*RawData, Header\nBytes)  ; Получаем ID, например.
Define *Host.ClipData = CheckHost(Hash)           ; Проверяем, нет ли уже такого вхождения.
If *Host = #Null                                  ; Если такого вхождения еще не было...
Define TrueSize = Header\nBytes, CS = TrueSize    ; Считаем размер.
Define *CMem    = EncodeData(*RawData, TrueSize)  ; Пробуем жать.
If *CMem : CS   = MemorySize(*CMem) : EndIf       ; Корректор.
If CheckSize(CS)                                  ; Проверяем размерность.
Define *DataNode.ClipData = RegisterNode(MetaText(Header, MetaCRC(*RawData, Header\nBytes), @MetaSize), Hash, Flag) ; Добавляем в список.
*DataNode\DataType = #CF_ENHMETAFILE : *DataNode\DataSize = TrueSize                    ; Указываем тип и размер.
*DataNode\Sizing = MetaSize                                                             ; Обязательно сохраняем размерность.
; -----------------
If *CMem : *DataNode\BinData = *CMem : *DataNode\CmpSize = CS : DeleteEnhMetaFile_(*Meta)
Else     : *DataNode\BinData = *Meta : EndIf
; -----------------
Else : DeleteEnhMetaFile_(*Meta)  : FreeCompression() : EndIf ; Высвобождаем всю память, значит.
Else : DeleteEnhMetaFile_(*Meta)  : FindHost(*Host, Shiftae(Flag)) : EndIf ; Выделяем прежнее вхождение.
FreeMemory(*RawData)                                       ; Обязательно высвобождаем.
EndIf : Else : DeleteEnhMetaFile_(*Meta) : RegisterImage() ; Значит, регистрируем изображение.
EndIf
EndMacro

Macro RegisterMetafile() ; Partializer.
Define *Meta = GetClipboardData(#CF_ENHMETAFILE, #True)
RegisterMetafileGuts()
EndMacro

Macro RegisterCTGuts(OutFormat, Prefix, Flag = #True) ; Partializer 2.0
If *DataPtr ; Если есть, с чем работать...
Define DataSize  = MemorySize(*DataPtr)
Define CRC32.s   = ComplexText2CRC(*DataPtr, DataSize, Prefix)
Define Hash.s    = ComplexText2MD5(*DataPtr, DataSize, Prefix)
Define *Host.ClipData = CheckHost(Hash)
If *Host = #Null           ; Если такого вхождения еще не было...
Define TrueSize = DataSize, CS = TrueSize         ; Считаем размер.
Define *CMem    = EncodeData(*DataPtr, TrueSize)  ; Пробуем жать.
If *CMem : CS   = MemorySize(*CMem) : EndIf       ; Корректор.
If CheckSize(CS)           ; Проверяем размерность.
Define PseudoNode.ClipData ; Хранилище для экстрактора данных.
PseudoNode\BinData = *DataPtr : PseudoNode\DataType = OutFormat : PseudoNode\DataSize = TrueSize
Define *DataNode.ClipData = RegisterNode(CRC32 + CompoundText(ExtractText(PseudoNode)), Hash, Flag)
*DataNode\DataType  = PseudoNode\DataType : *DataNode\Flattable = PseudoNode\Flattable
*DataNode\DataSize  = PseudoNode\DataSize : *DataNode\CacheData = PseudoNode\CacheData
If *DataNode\CacheData : *DataNode\CacheData\ClipData = *DataNode : EndIf ; Так надо.
If *CMem : *DataNode\BinData = *CMem      : *DataNode\CmpSize  = CS : FreeMemory(*DataPtr) 
Else     : *DataNode\BinData = *DataPtr   : EndIf
; ----
Else : FreeCompression() : FreeMemory(*DataPtr) : EndIf ; Обязательно высобождаем.
Else : FreeMemory(*DataPtr) : FindHost(*Host, Shiftae(Flag))
EndIf
EndIf
EndMacro

Macro RegisterCT(InFormat = System\ClipRTF, OutFormat = #CF_RichText, Prefix = Pref(#CF_RichText)) ; Partializer
Define *DataPtr = GetClipboardData(InFormat)
CloseClipboard_() : RegisterCTGuts(OutFormat, Prefix)
EndMacro

Procedure StoreData() ; !Partializer hub.
DisableDebugger : FreeMenu(#mTrayMenu) : EnableDebugger
With System     ; Работаем с системным объектом.
If System\PostFlag = #False   ; Если мы ничего не постили сами...
EnterCritical()               ; Блюдем быстродействие.
PrepareDelita()               ; Готовимся обрабатывать наш GUI.
Define DataType : SelectNode(#UnusedNode)   ; Убиваем полное выделение нода.
If System\AcceptNew                   ; Если принимаем новые данные...
System\LastSrc = GetClipboardOwner_() ; Получаем здесь окно-хозяина.
If System\LastSrc = System\SearchBar  ; Небольшое исключение для собственной строки поиска.
System\LastSrc = System\MainWindow       ; Quickfix. А использовать все одно лучше внутренние ф-ии.
EndIf : WaitOpen(#Null) ; Открываем буфер обмена (на чтение).
DataType = GetPriorityClipboardFormat_(@\ClipFormats, #Allformats) ; Выбираем формат.
Select DataType ; Выбираем по типу...
Case #CF_TEXT        : RegisterText()     ; Plain text.
Case #CF_BITMAP      : RegisterImage()    ; Bitmap image.
Case #CF_HDROP       : RegisterFiles()    ; File listing.
Case #CF_ENHMETAFILE : RegisterMetafile() ; Metafile picture.
Case \ClipHTML       : RegisterCT(\ClipHTML, #CF_HTML, Pref(#CF_HTML))
Case \ClipRTF        : RegisterCT()       ; Rich text.
Default              : CloseClipboard_()  ; Unknown (throw out)
EndSelect : UseDelita() : EndIf ; Иначе декрементируем флаг:
LeaveCritical()                 ; Продолжаем блюсти !
Else : \PostFlag - 1
EndIf ; Передаем по цепочке:
SendNotifyMessage_(\NextWindow, #WM_DRAWCLIPBOARD, 0, 0)
EndWith
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure ReintegrateList(*StartPtr.Unicode, *PlainFeed.Unicode)
While *PlainFeed\U ; До тех пор, пока что-то читается...
If *PlainFeed\U <> #DirDragDelim : *StartPtr\U = *PlainFeed\U : EndIf
*PlainFeed + SizeOf(Unicode) : *StartPtr + SizeOf(Unicode)
Wend
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro RestoreImageGuts(DN, Action = SetClipboardImage) ; Partializer 2.0
ImgSelect() : Action(ImgIdx) : ImgCheckUP()
EndMacro

Macro RestoreImage(DN) ; Stub.
RestoreImageGuts(DN)
EndMacro

Macro RestoreComplexTextGuts(DataNode, Format) ; Partializer 2.0
Define *DPtr = ExtractComplexData(DataNode)
Define Plain.s = ExtractText(DataNode, *DPtr) ; Вытаскиваем текст.
Define DSize = DataNode\DataSize
EndMacro

Macro RestoreComplexText(DataNode, Format) ; Partializer.
RestoreComplexTextGuts(DataNode, Format)
SetClipboardData(*DPtr, DSize, Format, 1)                                      ; RTF/HTML.
SetClipboardData(@Plain, StringByteLength(Plain)+#CharSize, #CF_UNICODETEXT,2) ; Plain.
CleanAfter(DataNode, *DPtr) ; Upfixing memory.
EndMacro

Macro RestoreFiles(Node) ; Partializer.
Define *List = AllocateMemory(SizeOf(Dropfiles) + Node\DataSize + SizeOf(Unicode) * #ListingExtra), *Header.Dropfiles = *List
*Header\fWide = #True : *Header\pFiles = SizeOf(Dropfiles)
Define Plain.s = ExtractText(Node) : ReintegrateList(*List + *Header\pFiles, @Plain)
SetClipboardData(*List, MemorySize(*list), #CF_HDROP) ; Вписываем и закрываем.
FreeMemory(*List)      ; Высвобождаем подальше.
EndMacro

Macro RestoreMetaGuts(Node, Action = SetClipboardData) ; Partializer 2.0
MetaSelect() : Action(CopyEnhMetaFile_(*MetaIDx, 0), 0, #CF_ENHMETAFILE, #MetaMode) : MetaCheckUP()
EndMacro
; -----------------------------------------
Macro RestoreMeta(Node) ; Partializer.
RestoreMetaGuts(Node)
EndMacro

Procedure RestoreData(NodeIdx, Forced = #False) ; !Partializers hub.
If NodeIdx <> #UnusedNode  ; Если есть, что копировать...
If NodeIdx <> UsedNodeIdx() Or Forced ; Оптимизация.
Define *DataNode.ClipData = Index2Node(NodeIdx)
System\PostFlag + 1 ; Повышаем значение счетчика.
Select *DataNode\DataType ; Выбираем по типу...
Case #CF_TEXT     : SetClipboardText(ExtractText(*DataNode)) ; Plain text.
Case #CF_BITMAP   : RestoreImage(*DataNode)                  ; Bitmap image.
Case #CF_RichText : RestoreComplexText(*DataNode, System\ClipRTF)  ; RTF.
Case #CF_HTML     : RestoreComplexText(*DataNode, System\ClipHTML) ; HTML.
Case #CF_HDROP    : RestoreFiles(*DataNode)                  ; File listing.
Case #CF_ENHMETAFILE : RestoreMeta(*DataNode)                ; Meta-file image.
EndSelect
SelectNode(NodeIdx) ; Showing it at GUI. Now marking slot:
If Forced = #False : HLLine(NodeIdx) : EndIf
EndIf
EndIf
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro TakeFocus()   ; Partializer.
Define Focus = GetActiveWindow()
EndMacro

Macro ResumeFocus() ; Partializer.
If Focus <> -1 : SetForegroundWindow_(WindowID(Focus)) : EndIf
EndMacro

Procedure WipeData(ForceOver.i = #False) ; Legacy way.
TakeFocus()
If ForceOver Or ListSize(System\ClipList()) = 0 Or WarnBox("do you really want to erase all absorbed data ?" + #CR$ +
"This action can not be cancelled after confirmation.") = #PB_MessageRequester_Yes
EnterCritical()
System\LockedNode = #Null ; Сразу сбрасываем.
ForEach System\ClipList():DisposeBinary(System\ClipList()):Next
FillMemory(@System\HotNodes, SizeOf(System\HotNodes))
ClearList(System\ClipList()) : ClearMap(System\Dictionary())
SendNotifyMessage_(#HWND_BROADCAST, System\CloseMsg, 0, 0)
ClearGadgetItems(#ClipList) : System\UsedNode = #Null
System\Lurker\Target = #sNone    ; На всякий пожарный.
ClearClipboard() : UpdateTitle() ; Очистка буффера обмена.
ResumeFocus()
EndIf
EndProcedure

Macro FormatKey(Key) ; Pseudo-procedure.
"[" + Str(Key) + "]@"
EndMacro

Procedure SetKey(*DataNode.ClipData, Key, ListIdx, NoReturn = #False)
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
If *DataNode         ; Если вообще есть, куда ставить.
If Key And *DataNode\Hotkey <> Key ; Если вписывается новая клавиша...
Define TextPos = 1   ; Позиция для вставки.
If *DataNode\Hotkey : SetKey(*DataNode, 0, ListIdx) : EndIf ; This!! Fix.
If *DataNode = System\UsedNode : TextPos + Len(#UsedNode) : EndIf : *DataNode\HotKey = Key : System\HotNodes[Key] = *DataNode
SetGadgetItemText(#ClipList, ListIdx, InsertString(GetGadgetItemText(#ClipList, ListIdx), FormatKey(Key), TextPos))
ElseIf NoReturn = 0  ; Проверка для Shift-клавиш. 
SetGadgetItemText(#ClipList,ListIdx,RemoveString(GetGadgetItemText(#ClipList,ListIdx),FormatKey(*DataNode\Hotkey),0,0,1)) ; Magick !
*DataNode\Hotkey = 0 : System\HotNodes[Key] = #Null : *DataNode\MenuSlot = 0 ; Продолжаем удаление.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
EndIf
EndIf
EndProcedure

Procedure LinkHotkey(NodeIdx, Key, NoReturn = #False)
If NodeIdx <> #UnusedNode ; Если есть, что линковать...
Define *DataNode.ClipData = Index2Node(NodeIdx), *HK = System\HotNodes[Key], DNHK = *DataNode\HotKey
If DNHK And DNHK <> Key : System\HotNodes[DNHK] = #Null : EndIf ; QuickFix.
If *HK And *HK <> *DataNode                                ; Если нужен обмен...
SetKey(*HK, *DataNode\HotKey, Node2Index(*HK))             ; Меняем ключи местами.
EndIf : SetKey(*DataNode, Key, NodeIdx, NoReturn) : EndIf  ; Собственно, оно самое.
EndProcedure

Procedure EmulatePasting()
Define *Window = GetForegroundWindow_()
Define *Victim = GetWindowThreadProcessId_(*Window, 0)
AttachThreadInput_(System\OwnHandle, *Victim, #True)
If System\Options\AltPasting    ; Если используем альтернативу...
Define *ForcusWin = GetFocus_() ; Получаем указатель.
PostMessage_(*ForcusWin, #WM_KEYDOWN, #VK_CONTROL, #KDownParam)
PostMessage_(*ForcusWin, #WM_KEYDOWN, 'V'        , #KDownParam)
PostMessage_(*ForcusWin, #WM_KEYUP  , 'V'        , #KUpParam)
PostMessage_(*ForcusWin, #WM_KEYUP  , #VK_CONTROL, #KUpParam)
Else : SendNotifyMessage_(GetFocus_(), #WM_PASTE, 0, 0)
EndIf : AttachThreadInput_(System\OwnHandle, *Victim, #False)
EndProcedure

Macro SelectKey() ; Partializer.
If GetActiveGadget() <> #SearchBar
Define Key = EventwParam() - '1' + 1
If Key >= 1 And Key <= #HotKeys And System\HotNodes[Key]
Define DIdx = Node2Index(System\HotNodes[Key]) ; Idx
RestoreData(DIdx) ; Выставляем выделение.
HLLine(DIdx) ; Поиск.
EndIf : EndIf
EndMacro

Macro CurrentHdrPos() ; Partializer.
StringByteLength(Header)
EndMacro

Macro Zerorial(Num = "") ; Pseudo-procedure.
RSet(Num, 10, "0")
EndMacro

Macro EndHdrField() ; Partializer.
Header + Zerorial() + #LF$
EndMacro

Procedure WriteHEaderField(*Base, *Offset, *Val)
Define Value.s = Zerorial(Str(*Val)) : CopyMemory(@Value, *Base + *Offset, StringByteLength(Value))
EndProcedure

Macro SpanHTML(HTML) ; Pseudo-procedure.
"<SPAN>" + HTML + "</SPAN>"
EndMacro

Procedure FormatHTML(HTML.S)
If HTML ; Если есть, о чем говорить...
HTML = #PreHTML + SpanHTML(HTML) + #PostHTML                                        ; Обрамляем тегами, вот да.
Define Header.s = "Version:0.9" + #LF$ + "StartHTML:", *HTMLStart = CurrentHdrPos() ; Начало HTML.
Define Header.s = EndHdrField() + "EndHTML:"         , *HTMLEnd   = CurrentHdrPos() ; Конец HTML.
Define Header.s = EndHdrField() + "StartFragment:"   , *FragStart = CurrentHdrPos() ; Начало фрагмента.
Define Header.s = EndHdrField() + "EndFragment:"     , *FragEnd   = CurrentHdrPos() ; Конец фрагмента.
Define Header.s = EndHdrField(), HTMLLen  = StringByteLength(HTML, #PB_UTF8), HDRLen = Len(Header)
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Define *HTMLResult = AllocateMemory(HDRLen + HTMLLen + #CharSize)                ; Аллокация результата.
Define *ActualHTML = HDRLen           : WriteHeaderField(*HTMLStart, @Header, *ActualHTML) ; Вписываем начало HTML.
Define *ActualEnd  = HDRLen + HTMLLen : WriteHeaderField(*HTMLEnd, @Header, *ActualEnd)    ; Вписываем конец HTML.
WriteHeaderField(*FragStart, @Header, *ActualHTML + StringByteLength(#PreHTML))  ; Фрагментарное начало.
WriteHeaderField(*FragEnd  , @Header, *ActualEnd  - StringByteLength(#PostHTML)) ; Фрагментарный конец.
PokeS(*HTMLResult, Header, -1, #PB_Ascii) : PokeS(*HTMLResult + *ActualHTML, HTML, -1, #PB_UTF8) ; Вписка результатов.
ProcedureReturn *HTMLResult
EndIf
EndProcedure

Procedure ClipHTML(HTML.S, Plain.S)
If HTML ; Если есть, о чем говорить...
Define *HTMLResult = FormatHTML(HTML)
If *HTMLResult ; Если Форматирование прошло успешно...
SetClipboardData(*HTMLResult, MemorySize(*HTMLResult), System\ClipHTML, 1)
SetClipboardData(@Plain, StringByteLength(Plain)+#CharSize, #CF_UNICODETEXT, 2)
FreeMemory(*HTMLResult) : ProcedureReturn #True ; Рапортуем весь успех.
EndIf
EndIf
EndProcedure
;}
;{ --Input/Output--
Macro FontDataPref() ; Partializer.
Define hWnd  = GadgetID(#VoidGadget) 
Define hdc   = GetDC_(hWnd) 
*FontID  = SelectObject_(hdc, *FontID) 
Define bsize = GetOutlineTextMetrics_(hdc, 0, 0) 
If bsize = 0 : SelectObject_(hdc, *FontID) : ReleaseDC_(hWnd,hdc) : ProcedureReturn : EndIf 
*otm = AllocateMemory(bsize) : *otm\otmSize = bsize 
GetOutlineTextMetrics_(hdc,bsize,*otm) 
EndMacro

Macro FontDataPost() ; Partializer.
FreeMemory(*otm) : SelectObject_(hdc, *FontID) : ReleaseDC_(hWnd, hdc) 
EndMacro

Procedure.s GetFontNameEX(*FontID) 
Define *otm.OUTLINETEXTMETRIC, FontName.s
FontDataPref() : FontName = PeekS(*otm + *otm\otmpFamilyName) : FontDataPost()
ProcedureReturn FontName 
EndProcedure 

Procedure.s GetFontNameReserve(*FontID)
Protected finfo.LOGFONT : GetObject_(*FontId, SizeOf(LOGFONT), @finfo)
ProcedureReturn PeekS(@finfo\lfFaceName) ; Font name
EndProcedure

Procedure.f GetFontHeight(*FontID)
Protected fontid, finfo.LOGFONT, Height.f, hDC
GetObject_(*FontId, SizeOf(LOGFONT), @finfo)
hDc = StartDrawing(WindowOutput(#MainWindow))
Height = -MulDiv_(finfo\lfheight, 72, GetDeviceCaps_(hdc, #LOGPIXELSY))
StopDrawing() : ProcedureReturn Height
EndProcedure

Procedure GetFontStyle(*FontID)
Protected finfo.LOGFONT, Attrib
GetObject_(*FontId,SizeOf(LOGFONT), @finfo)
With finfo 
If \lfWeight > #FW_NORMAL          : Attrib | #PB_Font_Bold        : EndIf ; Bold.
If \lfItalic                       : Attrib | #PB_Font_Italic      : EndIf ; Italic.
If \lfStrikeOut                    : Attrib | #PB_Font_StrikeOut   : EndIf ; Strike Out.
If \lfUnderline                    : Attrib | #PB_Font_Underline   : EndIf ; Underline.
ProcedureReturn Attrib | #PB_Font_HighQuality ; Font attributes
EndWith
EndProcedure

Procedure.s GetFontName(*FontID)
Define Name.s = GetFontNameEX(*FontID)
If NAme = "" : ProcedureReturn GetFontNameReserve(*FontID) : EndIf
ProcedureReturn NAme
EndProcedure

Macro GetDefaultFont() ; Pseudo-procedure.
GetGadgetFont(#VoidGadget)
EndMacro

Macro OpenIniFile() ; Pseudo-procedure.
OpenPreferences(#IniFile, #PB_Preference_GroupSeparator)
EndMacro

Macro RegBoolOption(KeyName, Fld, GadgetID, Def = #False) ; Parializer.
AddMapElement(System\Flagi(), KeyName) : Define *Flag.OptionFlag = System\Flagi(KeyName)
*Flag\IniName = KeyName : *Flag\Data = @System\Options\Fld : *Flag\DefVal = Def
*Flag\Gadget = #Check#GadgetID
If ReadPreferenceInteger(KeyName, Def) : System\Options\Fld  = #True : EndIf
EndMacro

Macro InvalidateGroup(GName) ; Pseudo-procedure.
RemovePreferenceGroup(GName) : PreferenceGroup(GName)
EndMacro

Macro WritePrefixedInteger(Key, Value, Pref = ".") ; Pseudo-procedure.
CompilerIf Pref <> "." : Define Prefix.s = Pref
CompilerEndIf          : WritePreferenceInteger(Prefix + "." + Key, Value)
EndMacro

Macro WriteWindowPos(WinID, Pref = ".") ; Pseudo-procedure.
WritePrefixedInteger("X", WindowX(WinID), Pref)
WritePrefixedInteger("Y", WindowY(WinID))
EndMacro

Macro WriteWindowRect(WinID, Pref = ".") ; Pseudo-procedure.
WriteWindowPos(WinID, Pref) 
WritePrefixedInteger("Width", WindowWidth(WinID))
WritePrefixedInteger("Height", WindowHeight(WinID))
EndMacro

Procedure WriteAllPrefs()
CloseFile(CreateFile(#PB_Any, #IniFile)) : OpenIniFile() 
InvalidateGroup("GUI.List") ; Открываем -> cоздаем группу.
WritePreferenceInteger("List.BackColor" , GetGadgetColor(#ClipList, #PB_Gadget_BackColor))
WritePreferenceInteger("List.FrontColor", GetGadgetColor(#ClipList, #PB_Gadget_FrontColor))
WritePreferenceString ("List.FontName"  , GetFontName(System\ListFont))
WritePreferenceFloat  ("List.FontSize"  , GetFontHeight(System\ListFont))
WritePreferenceInteger("List.FontFlags" , GetFontStyle(System\ListFont))
With System\Options
InvalidateGroup("Misc") ; Создаем.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
ForEach System\Flagi() : Define *Flag.OptionFlag = System\Flagi()
WritePreferenceInteger(*Flag\IniName, *Flag\Data\I) : Next
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
WritePreferenceInteger("MaxEntries"  , \ListMax)
WritePreferenceInteger("SizeLimit"   , \SizeLimit)
WritePreferenceInteger("PreserveHotkeys", \HPreserve)
EndWith : InvalidateGroup("Volatile") ; Открываем->cоздаем.
WriteWindowRect(#MainWindow, "Window")
WriteWindowPos(#SettingsWindow, "OptWin")
WriteWindowRect(System\Panopticum\WindowID, "PanWin")
WritePreferenceInteger("Raster.DefType", System\RasterType)
WritePreferenceInteger("Data.AcceptNew", System\AcceptNew)
WritePreferenceInteger("Render.BackColor", System\RenderBack)
ClosePreferences()
EndProcedure

Macro ReadPreferenceIntegerEX(Value, Key, InitValue, MaxValue = Dummy) ; Pseudo-procedure.
Value = ReadPreferenceInteger(Key, InitValue) : Define Dummy = Value ; Считываем.
If Value < 0 Or Value > MaxValue : RemovePreferenceKey(Key) : Value = InitValue : EndIf
EndMacro

Macro ReadPreferenceFloatEX(Value, Key, InitValue) ; Pseudo-procedure.
Value = ReadPreferenceFloat(Key, InitValue) ; Считываем.
If Value < 0 : RemovePreferenceKey(Key) : Value = InitValue : EndIf
EndMacro
; --------------------------
Macro ReadWindowPosition(WindowID, Prefix) ; Pseudo-procedure.
PlaceWindowStrict(WindowID, ReadPreferenceInteger(Prefix + ".X", WindowX(WindowID)),
                            ReadPreferenceInteger(Prefix + ".Y", WindowY(WindowID)))
EndMacro

Macro ReadWindowRect(Window, Prefix)
Define Width  : ReadPreferenceIntegerEX(Width , Prefix + ".Width" , #MinWidth)
Define Height : ReadPreferenceIntegerEX(Height, Prefix + ".Height", #MinHeight)
ResizeWindow(Window, #PB_Ignore, #PB_Ignore, Width, Height)
ReadWindowPosition(Window, Prefix)
EndMacro
; --------------------------
Macro RestoreVolatile() ; Partializer.
PreferenceGroup("Volatile") ; Загружаем.
ReadPreferenceIntegerEX(System\RasterType, "Raster.DefType", 0, 3) ; Формат растра.
If ReadPreferenceInteger("Data.AcceptNew", #True) : System\AcceptNew = #True : EndIf
ReadWindowRect(#MainWindow, "Window") : ReadWindowRect(System\Panopticum\WindowID, "PanWin")
ReadWindowPosition(#SettingsWindow, "OptWin")
ReadPreferenceIntegerEX(System\RenderBack, "Render.BackColor", DefaultBackColor(), #White)
EndMacro

Macro LoadHQFont(Index, FontName, FontSize, FontStyle) ; Pseudo-procedure.
LoadFont(Index, FontName, FontSize, FontStyle | #PB_Font_HighQuality)
EndMacro

Macro RestoreList() ; Partializer.
Define Color.l ; Для проверки.
PreferenceGroup("GUI.List") ; Загружаем.
ReadPreferenceIntegerEX(Color, "List.BackColor", DefaultBackColor(), #White)
SetGadgetColor(#ClipList, #PB_Gadget_BackColor, Color) ; Задний фон.
ReadPreferenceIntegerEX(Color, "List.FrontColor", DefaultFrontColor(), #White)
SetGadgetColor(#ClipList, #PB_Gadget_FrontColor, Color) ; Цвет текста.
Define FSize, FName.s = ReadPreferenceString("List.FontName", "")
If FName ; Если доступно имя шрифта...
ReadPreferenceFloatEX(FSize, "List.FontSize", 8.5) ; Читаем размер.
Define FStyle = ReadPreferenceInteger("List.FontFlags", 0)
System\ListFont = LoadHQFont(#fListFont, FName, FSize, FStyle)
Else : System\ListFont = GetDefaultFont() : EndIf : SetGadgetFont(#ClipList, System\ListFont)
EndMacro

Procedure WriteDataProtected(FileNum, *Ptr, Size) ; Replacer.
WriteData(FileNum, *Ptr, Size) ; Самое главное.
WriteInteger(Filenum, CRC32Fingerprint(*Ptr, Size))
EndProcedure

Procedure WriteDataEx(*Pointer, Size) ; Replacer.
DisableDebugger                                        ; Отключаем на всякий.
Define *PackData = EncodeData(*Pointer, Size)
If *PackData : Define PackSize = MemorySize(*PackData) ; Вот получаем размер...
WriteInteger(0, PackSize)                              ; ...Вписываем его...
WriteDataProtected(0, *PackData, PackSize)             ; ...И сами данные.
FreeMemory(*PackData)                                  ; Высвобождаем память.
Else : WriteInteger(0, 0) : WriteDataProtected(0, *Pointer, Size) : EndIf ; Или просто пишем.
EnableDebugger                                         ; Ставим обратно, да-да.
EndProcedure

Macro DumpText(Text) ; Pseudo-procedure.
Define Plain.s = Text       ; Аккумулятор.
WriteInteger(0, Len(Plain)) ; Длина текста.
WriteDataEx(@Plain, StringByteLength(Plain))
EndMacro

Macro SmartWrite(Field = BinData) ; Replacer
WriteDataProtected(0, *DataNode\Field, *DataNode\CmpSize)
EndMacro

Macro DumpSizing(DN)
WriteDataProtected(0, @DN\Sizing, SizeOf(Point))
EndMacro

Macro SaveData(DataNode, Extra = 0) ; Partializer
WriteDataEx(DataNode\BinData, DataNode\DataSize + Extra)
EndMacro

Procedure AllowDumping(Format)
Select Format
Case #CF_BITMAP, #CF_ENHMETAFILE : ProcedureReturn System\Options\SaveImages
Case #CF_HDROP                   : ProcedureReturn System\Options\SaveLists
EndSelect                        : ProcedureReturn #True
EndProcedure

Procedure.s GetListText(*DataNode.ClipData)
Define LText.s = GetGadgetItemText(#ClipList, Node2Index(*DataNode))
ProcedureReturn Mid(LText, FindString(LText, Pref(*DataNode\DataType)))
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro PackText(Txt, Fld = BinData) ; Partializer.
If *DataNode\CmpSize : SmartWrite(Fld)                   ; Пишем запакованные данные текста. 
Else : WriteDataProtected(0, @Txt, StringByteLength(Txt)) : EndIf ; Иначе - уж пишем, как оно там есть.
EndMacro

Macro DumpImage() ; Partializer
If *DataNode\CmpSize = 0                     ; Если предполагается писать напрямую...
StartDrawing(ImageOutput(*DataNode\BinData)) ; Начинаем отрисовку.
WriteDataProtected(0, DrawingBuffer(), DrawingBufferPitch() * *DataNode\Sizing\Y) ; Вписываем целиком.
StopDrawing() : Else : SmartWrite() : EndIf  ; Вписываем сжатый вариант.
EndMacro

Macro PackData(Node, WriteSize = *DataNode\DataSize) ; Partiazlier.
If *DataNode\CmpSize : SmartWrite(BinData) : Else : WriteDataProtected(0, *DataNode\BinData, WriteSize) : EndIf
EndMacro

Macro DumpListing(Node) ; Partializer.
PackText(*DataNode\TextData)
EndMacro

Macro PackEMF(Node) ; Partializer.
If *DataNode\CmpSize = 0
GetMetaHeader(Node\BinData) : GetMetaBits(Node\BinData) : WriteDataProtected(0, *RawData, Node\DataSize) : FreeMemory(*RawData)
Else : SmartWrite(BinData)
EndIf
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure SaveDump() ; !Partializers hub.
If System\Options\UseDump   ; Если требуется сохранение дампа...
If CreateFile(0, #Dumpfile) ; Открываем файл.
SetFileAttributes(#Dumpfile, #PB_FileSystem_Hidden)
Define *DataNode.ClipData, Nodes = ListSize(System\ClipList())
WriteLong(0, #DumpSig) ; Записываем сигнатуру.
WriteInteger(0, 0)     ; Резервируем размер списка.
ForEach System\ClipList() : *DataNode = System\ClipList()
If AllowDumping(*DataNode\DataType) ; Если тип можно дампать...
WriteInteger(0, *DataNode\DataType) ; Вписываем тип.
WriteInteger(0, *DataNode\DataSize) ; Вписываем размер.
WriteInteger(0, *DataNode\Hotkey)   ; Вписываем клавишу.
WriteInteger(0, *DataNode\CmpSize)  ; Вписываем запакованный размер.
WriteInteger(0, *DataNode\Flattable); Вписываем флаг сжимаемости.
WriteInteger(0, *DataNode\TimeStamp); Вписываем время получения.
WriteInteger(0, *DataNode\RandFlag) ; Вписываем типоспецифичные флаги.
WriteDataEx(*DataNode\VPSizing, SizeOf(Rect)) ; Вписываем размерность окна просмотра.
DumpText(*DataNode\WSource)    ; Сохраняем название окна-источника.
DumpText(*DataNode\Comment)    ; Сохраняем комментарий.
DumpText(*DataNode\DictID)     ; Сохраняем ID для словаря.
DumpText(GetListText(*DataNode)) ; Сохраняем текст для списка. Наконец-то.
DumpSizing(*DataNode)            ; Вписываем размерность.
Select *DataNode\DataType ; В зависимости от типа...
Case #CF_TEXT   : PackText(*DataNode\TextData)      ; Сохраняем текст.
Case #CF_BITMAP : DumpImage()                       ; Сохраняем изображение.
Case #CF_HDROP  : DumpListing(*DataNode)            ; Сохраняем листинг.
Case #CF_ENHMETAFILE : PackEMF(*DataNode)           ; Просто вписываем данные.
Case #CF_RichText, #CF_HTML : PackData(*DataNode)   ; RTF/HTML.
EndSelect
Else : Nodes - 1 ; Декрементируем.
EndIf
Next : FileSeek(0, SizeOf(Long)) : WriteInteger(0, Nodes)
CloseFile(0) ; Закрываем файл.
EndIf
EndIf
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro DumpError(NeedsWiping = #True) ; Partializer.
CompilerIf NeedsWiping : WipeData(#True) ; Очищаем списки.
CompilerEndIf : ErrorBox("corrupted dump file !" + #CR$ + "Parsing aborted.")
CompilerIf NeedsWiping : Break ; Выходим из цикла.
CompilerEndIf
EndMacro

Macro ReadDataProtected(Fnum, Ptr, Size)
ReadData(Fnum, Ptr, Size) ; Самое главное.
If CRC32Fingerprint(Ptr, Size) <> ReadInteger(0) : DumpError() : EndIf
EndMacro

Macro ReadDataEx(Pointer, Size) ; Replacer.
If Size     < 0 : DumpError() : EndIf
Define PackSize = ReadInteger(0)
If PackSize < 0 : DumpError() : EndIf
If PackSize ; Если там что-то запаковано...
Define *PackData = AllocateMemory(PackSize)
ReadDataProtected(0, *PackData, PackSize)                     ; Читаем данные.
DecodeData(*PackData, PackSize, Size, Pointer)
FreeMemory(*PackData)                                ; Высвобождаем.
Else : ReadDataProtected(0, Pointer, Size) : EndIf
EndMacro

Macro LoadSizing(DN)
ReadDataProtected(0, @DN\Sizing, SizeOf(Point))
EndMacro

Macro LoadText(Receiver = *DataNode\TextData) ; Partializer.
Define Size = ReadInteger(0) : Receiver = Space(Size) : ReadDataEx(@Receiver, Size << 1)
EndMacro

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro RestoreCmpData(Node) ; Partializer.
Node\BinData = AllocateMemory(Node\CmpSize) : ReadDataProtected(0, Node\BinData, Node\CmpSize)
EndMacro

Macro RestoreDataFast(Node, CPress) ; Replacer.
If CPress : RestoreCmpData(Node)
Else : Node\BinData = AllocateMemory(Node\DataSize) : ReadDataProtected(0, Node\BinData, Node\DataSize) : EndIf
EndMacro

Macro LoadTextSmart(Node, CPress) ; Replacer.
If CPress : RestoreCmpData(Node)
Else : Node\TextData = Space(Node\Sizing\X) : ReadDataProtected(0, @Node\TextData, Node\DataSize) : EndIf
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro LoadTextData() ; Partializer.
LoadTextSmart(*DataNode, *DataNode\CmpSize)
EndMacro

Macro LoadImage() ; Partializer.
If *DataNode\CmpSize : RestoreCmpData(*DataNode)
Else : *DataNode\BinData = CreateImage(#PB_Any, *DataNode\Sizing\X, *DataNode\Sizing\Y, #BitDepth)
StartDrawing(ImageOutput(*DataNode\BinData)) ; Открываем поверхность.
ReadDataProtected(0, DrawingBuffer(), *DataNode\DataSize)
StopDrawing() : EndIf    ; Закрываем поверхность.
EndMacro

Macro LoadComplexText(DataNode) ; Partializer.
RestoreDataFast(DataNode, DataNode\CmpSize)
EndMacro

Macro LoadListing() ; Partializer.
LoadTextSmart(*DataNode, *DataNode\CmpSize)
EndMacro

Macro LoadMeta() ; Partializer.
If *DataNode\CmpSize : RestoreCmpData(*DataNode)
Else : Define *RawData = AllocateMemory(*DataNode\DataSize)
ReadDataProtected(0, *RawData, *DataNode\DataSize)
*DataNode\BinData = SetEnhMetaFileBits_(*DataNode\DataSize, *RawData)
FreeMemory(*RawData)
EndIf 
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure RestoreDump() ; !Partializers hub.
If System\Options\UseDump ; Если требуется восстановления дампа...
Define *DataNode.ClipData, ListLine.s
If ReadFile(0, #Dumpfile) ; Открываем файл.
If ReadLong(0) = #DumpSig ; Проверка сигнатуры.
Define TextSize, Node, NodeCount = ReadInteger(0) ; Читаем кол-во нодов.
If NodeCount < 0 : DumpError(#False) : EndIf ; Сразы выходи при неправильном числе.
Else : DumpError(#False) : EndIf     ; Выводим ошибку.
For Node = 1 To NodeCount : AddElement(System\ClipList())
*DataNode = System\ClipList()        ; Получаем элемент.
*DataNode\DataType = ReadInteger(0)  ; Читаем тип.
*DataNode\DataSize = ReadInteger(0)  ; Читаем размер.
If *DataNode\DataSize < 0 : DumpError() : EndIf ; Проверка размера.
Define HotKey  = ReadInteger(0)      ; Читаем хоткей.
If HotKey<0 Or HotKey>#HotKeys Or (Hotkey And System\HotNodes[HotKey]) :DumpError():EndIf ; Проверка клавиши.
*DataNode\CmpSize = ReadInteger(0)   ; Читаем ужатый размер.
*DataNode\Flattable = ReadInteger(0) ; Читаем флаг сжимаемости.
*DataNode\TimeStamp = ReadInteger(0) ; Читаем дату абсорбции.
*DataNode\RandFlag  = ReadInteger(0) ; Читаем типоспецифичные флаги.
ReadDataEx(*DataNode\VPSizing, SizeOf(Rect)) ; Восстанавливаем размерность окна просмотра.
LoadText(*DataNode\WSource)          ; Читаем окно обретения.
LoadText(*DataNode\Comment)          ; Читаем сам комментарий.
LoadText(*DataNode\DictID)           ; Ну почему я изначально так не сделала ?
LoadText(ListLine.s)                 ; Читаем строчку для дальнейшего листинга.
LoadSizing(*DataNode)                ; Читаем размерность.
Select *DataNode\DataType ; В зависимости от типа...
; ------------------
Case #CF_TEXT     : LoadTextData() ; Грузим текст.
Case #CF_BITMAP   : LoadImage()    ; Загружаем изображение.
Case #CF_HDROP    : LoadListing()  ; DIR.
Case #CF_ENHMETAFILE : LoadMeta()  ; EMF.
Case #CF_RichText, #CF_HTML : LoadComplexText(*DataNode) ; RTF.
Default : DumpError() ; Прекращаем загрузку.
; ------------------
EndSelect : LinkIDs(*DataNode, *DataNode\DictID) : Add2List(ListLine, *DataNode, #False)
SetKey(*DataNode, HotKey, Node - 1)
If *DataNode\BinData : *DataNode\TextData = "" : EndIf ; For great justice. Really.
Next Node : CloseFile(0) ; Закрываем файл.
EndIf
EndIf
EndProcedure

Macro DoBackup() ; Partializer.
SetWindowTitle(#MainWindow, #Title + " [...do not disturb...]")
WriteAllPrefs() ; Сохраняем все настройки.
SaveDump() : UpdaTetitle() ; Дамп данных.
EndMacro

Procedure ErrorHandler() ; Last resort.
Define Diagnose.s
Select ErrorCode() ; Определяем причину.
Case #PB_OnError_InvalidMemory          : Diagnose = "invalid memory access at $"+Hex(ErrorTargetAddress())+" !"
Case #PB_OnError_Floatingpoint          : Diagnose = "floating point error !"
Case #PB_OnError_Breakpoint             : Diagnose = "unknown debugger's breakpoint reached !"
Case #PB_OnError_IllegalInstruction     : Diagnose = "attempt to execute an illegal instruction !"
Case #PB_OnError_PriviledgedInstruction : Diagnose = "attempt to execute a priviledged instruction !"
Case #PB_OnError_DivideByZero           : Diagnose = "integer division by zero !"
EndSelect : If System\NextWindow : ChangeClipboardChain_(System\MainWindow, System\NextWindow) : EndIf
Errorbox(Diagnose + " It's irreparable." + #CR$ + "Application would now " + IIFS(System\Options\AutoReboot, 
"try to restart itself.", "be terminated."))
If System\Options\AutoReboot   ; Если требуется перезапуск с чистого листа...
ReleaseMutex_(System\DupMutex) : RunProgram(ProgramFilename()) ; Рестарт.
EndIf
EndProcedure

Macro PokeLayout(LChar, RChar) ; Partializer.
If System\LayOut[LChar] = LChar : System\LayOut[LChar] = RChar : EndIf
If System\LayOut[RChar] = RChar : System\LayOut[RChar] = LChar : EndIf
EndMacro

Macro InitTable() ; Pseudo-procedure.
EnableASM
; Primary initialization.
LEA EAX, System\Layout 
ADD EAX, #LayoutEdge
MOV ECX, #LayoutTable
FillIn: 
MOV word [EAX], CX
SUB EAX, #CharSize
LOOP l_fillin
DisableASM
; Actual initialization.
Define *LTable.CHArExtra = ?QWERTY_lat, *RTable.CHArExtra = ?QWERTY_rus
While *LTable\U : PokeLayout(*LTable\U, *RTable\U) : *LTable\Lense = UCase(*LTable\Lense)
*RTable\Lense = UCase(*RTable\Lense) : PokeLayout(*LTable\U, *RTable\U)
*LTable + #CharSize : *RTable + #CharSize : Wend
; Tables feeder:
DataSection : QWERTY_lat: :Data.s "~`"+        "@#:;&qwertyuiop{[}]|asdfghjkl$^"+#DQUOTE$+"'zxcvbnm<,>./?"
              QWERTY_rus: :Data.s "Ёё"+#DQUOTE$+"№Жж?йцукенгшщзХхЪъ/фывапролд;:Э"+        "эячсмитьБбЮю.,"
EndDataSection
EndMacro

Procedure TextWidthSmart(Text.s, *Font)
StartDrawing(WindowOutput(#MainWindow)) : DrawingFont(FontID(*Font))
Define Width = TextWidth(Text) : StopDrawing() : ProcedureReturn Width
EndProcedure

Macro FaultyOS()
(Not System\XPLegacy)
EndMacro

Macro SetComposition(Win, Compos) ; Partializer.
If FaultyOS() : SwitchStyle(Win, #WS_EX_COMPOSITED, Compos) : EndIf
EndMacro

Macro SetupBuffers(BufAccum, Compos = #True) ; Pseudo-procedure.
If FaultyOS() : CloseGadgetList() : BufAccum\Voider = TextGadget(#PB_Any, 0, 0, 0, 0 , ":IamError:")
Define *CID = GadgetID(BufAccum\Container) ; Переходим к основному списку, на всякий случай.
SwitchStyle(*CID, #WS_CLIPCHILDREN, #True, #GWL_STYLE): SetComposition(*CID, Compos)
BufAccum\Splitter = SplitterGadget(#PB_Any, 0, 0, 0, 0, BufAccum\Voider, BufAccum\Container, #PB_Splitter_FirstFixed)
ChangeCB(GadgetID(BufAccum\Splitter), SplitterCallback()) ; Вся соль, собственно.
EndIf
EndMacro
;}
;{ --Informer GUI--
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure ReDrawInformer(*IBlock.InfoBlock)
StartDrawing(CanvasOutput(*IBlock\Informer)) ; Начинаем рисовать, да.
Define Width = OutputWidth(), Height = OutputHeight()
Box(0, 0, Width, Height, #Black) : DrawingFont(FontID(#fInFont))
If *IBlock\InfoSum ; Хоть оно и маловероятно, но все же защита от дурака.
DrawingMode(#PB_2DDrawing_Transparent) ; Прозрачный текст, просто на всякий случай.
Define TextHeight = (Height - TextHeight(*IBlock\InfoSum)) >> 1 ; Дабы выводить по центру...
Define TextOffset = *IBlock\InfoShift                ; Получаем смещение для циклической отрисовки.
Define TextWidth = TextWidth(*IBlock\InfoSum)        ; Для оптимизации.
Repeat : DrawText(TextOffset, TextHeight, *IBlock\InfoSum, #White)
TextOffset + TextWidth : Until TextOffset => Width   ; Рисуем по циклу.
EndIf : StopDrawing()
EndProcedure

Macro ResizeInformer(IBlock, LOffset = #UniVPOffset, BOffset = #InfBottom) ; Partializer
ResizeGadget(IBlock\Informer, #PB_Ignore, FullHeight - BOffset, FullWidth - GadgetX(IBlock\Informer) - LOffset, #PB_Ignore)
ResizeGadget(IBlock\InBox, #PB_Ignore, FullHeight - BOffset, #PB_Ignore, #PB_Ignore) : ReDrawInformer(IBlock)
EndMacro

Procedure ShiftFactor(Text.s)
Define InfoWidth = TextWidthSmart(Text, #fInFont), Factor = InfoWidth % #ShiftStep
ProcedureReturn -(InfoWidth - Factor)
EndProcedure

Procedure ShiftInfo(*IBlock.InfoBlock)
*IBlock\InfoShift - #ShiftStep
If *IBlock\InfoShift <= ShiftFactor(*IBlock\InfoSum) : *IBlock\InfoShift = 0 : EndIf ; Зацикливаем.
ReDrawInformer(*IBlock)
EndProcedure

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure Inform_Handler(*GadgetID, Message, wParam, lParam) ; Callback
Select Message
Case #WM_COMMAND ; Как-то так до лучших времен.
If Hiword(wParam) = #EN_SETFOCUS ; Если нам рапортовали о фокусе...
Define *VP.ViewPort = ExtractWP(GetParent_(*GadgetID))
If *VP : SetActiveGadget(*VP\ViewArea) : Else : SAG() : EndIf
EndIf
EndSelect
ProcedureReturn ChainOldCB(*GadgetID)
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Procedure Changer()
Define *Informer = EventGadget(), Idx = GetGadgetState(*Informer) 
If Idx <> -1 : SetClipboardText(Mid(GetGadgetItemText(*Informer, Idx), 3))
SetGadgetState(*Informer, -1) : EndIf ; Убиваем.
EndProcedure

Macro InitInformer(IBlock, Offset = #UniVPOffset) 
IBlock\InBox    = ComboBoxButtonGadget(#PB_Any, Offset, 0) : SetGadgetFont(IBlock\InBox, FontID(#fBoxedFont))
IBlock\Informer = CanvasGadget(#PB_Any, Offset - ComboWidth(IBlock\InBox), 0, 0, #InfHeight, #PB_Canvas_Border)
ChangeCB(GadgetID(IBlock\InBox), Inform_Handler()) : SetProp_(GadgetID(IBlock\InBox), "InBox", IBlock\InBox)
BindGadgetEvent(IBlock\InBox, @Changer(), #PB_EventType_Change)
EndMacro

Procedure.s FormatKB(ByteCount.D)
Define Report.s = StrD(ByteCount / #KB, 1)
If ValD(Report) : ProcedureReturn Report : Else : ProcedureReturn "0.1-" : EndIf
EndProcedure

Procedure CmpRate(Original.D, Compressed.D)
ProcedureReturn (1 - (Compressed / Original.D)) * 100
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro TimeField(Node) ; Field partializer.
"Registered at " + FormatDate("[%dd/%mm/%yyyy...(%hh:%ii:%ss)]", Node\TimeStamp)
EndMacro

Procedure.s ExtractKB(*Node.ClipData, Delim.s = #IDelim) ; Field partializer.
With *Node
Define Result.s = "KB: " + FormatKB(\DataSize)
If \CmpSize : Result + "/" + FormatKB(\CmpSize) + " <" + CmpRate(\DataSize, \CmpSize) + "% compressed>" : EndIf
ProcedureReturn Result + Delim
EndWith
EndProcedure

Macro RemField(Node) ; Field partializer.
"'" + Node\Comment + "'"
EndMacro 

Macro SrcField(Node) ; Field partializer.
"«" + Node\WSource + "»"
EndMacro

Procedure.s GatherNodeInfo(*DataNode.ClipData, Delim.s = #IDelim)
If *DataNode ; Если есть нод на анализ...
With *DataNode
Define GigaLine.s = TimeField(*DataNode) + Delim + "Recognized as " + \DictID + Delim + ExtractKB(*DataNode)
If \Comment : GigaLine + "Remarked by "      + RemField(*DataNode) + Delim : EndIf ; Дописываем комментариий.
If \WSource : Gigaline + "Associated with "  + SrcField(*DataNode) + Delim : EndIf ; Дописываем окно-источник.
ProcedureReturn GigaLine
EndWith
Else : ProcedureReturn "]-No data-[#"
EndIf
EndProcedure

Procedure SetInformerText(*IBlock.InfoBlock, Text.s)
Define TText.s = FlattenText(Text), Shifter = -ShiftFactor(TText) / #ShiftStep
*IBlock\InfoSum = TText : *IBlock\InfoShift = -Random(Shifter) * #ShiftStep : ReDrawInformer(*IBlock)
EndProcedure

Macro AddField(fData, Field = -1) ; Partializer.
AddGadgetItem(*IBlock\InBox, -1, System\Bullet + " " + fData) : Define LastIdx = CountGadgetItems(*IBlock\InBox)
If *IBlock\InBox = System\Informator\InBox : SetGadgetItemData(*IBlock\InBox, LastIdx - 1, Field) ; Ставим флаг для поиска.
If System\HoldField = Field : Hfld = LastIdx : EndIf : EndIf               ; Проверка удержания.
EndMacro

Procedure SetInformerFields(*IBlock.InfoBlock, *Node.ClipData)
Define HFld : ClearGadgetItems(*IBlock\InBox) ; Сразу очищаем....
If *Node : DisableGadget(*IBlock\InBox, #False) : AddField(*Node\DictID, #sDictID) ; ID данных.
If *Node\Comment : AddField("Rem.:" + RemField(*Node), #sRemark)     : EndIf ; Комментарий.
If *Node\WSource : AddField("Source: " + SrcField(*Node), #sWSource) : EndIf ; Окно-источник.
AddField(TimeField(*Node))  : AddField(ExtractKB(*Node, ""))            ; Время и размер.
ComboBoxButtonListWidth(#MainWindow, *IBlock\InBox)                     ; Выставляем размер.
Else : DisableGadget(*IBlock\InBox, #True) ; Отключаем, если там и включать нечего. Так надо.
EndIf : If HFld : SetGadgetState(*IBlock\InBox, HFld - 1) : EndIf       ; Ставим выделение на поля.
EndProcedure

Macro UpdateInformer(IBlock, Node = *DataNode) ; Partializer.
SetInformerText(IBlock, GatherNodeInfo(Node)) : SetInformerFields(IBlock, Node)
EndMacro

Procedure GetDroppedState(*Informer)
ProcedureReturn SendMessage_(GadgetID(*Informer), #CB_GETDROPPEDSTATE, 0, 0)
EndProcedure

Procedure SetDroppedState(*Informer, State)
SendMessage_(GadgetID(*Informer), #CB_SHOWDROPDOWN, State, 0)
EndProcedure

Macro ConnectInformer(iBlock = System\Informator) ; Pseudo-procedure.
SetDroppedState(iBlock\InBox, #False) : UpdateInformer(iBlock, *HLNode)
EndMacro
;}
;{ --Options GUI management--
Macro CheckLimit() ; Partializer.
Define State = GetGadgetState(#CheckLimit) ! 1
DisableGadget(#LimitSpin, State)
DisableGadget(#CheckPreserve, State)
If State : SetGadgetState(#CheckPreserve, #False) : EndIf
EndMacro

Macro CheckSizeLimit() ; Partializer.
DisableGadget(#SizeSpin, GetGadgetState(#CheckMaxSize) ! 1) 
EndMacro

Macro CheckDump() ; Partializer.
Define State = GetGadgetState(#CheckUseDump) ! 1
DisableGadget(#CheckDumpImages, State)
DisableGadget(#CheckDumpLists , State)
If State : SetGadgetState(#CheckDumpImages, #False)
SetGadgetState(#CheckDumpLists, #False) : EndIf
EndMacro

Macro CheckTray() ; Partializer.
Define State = GetGadgetState(#CheckTray) ! 1
DisableGadget(#CheckFixed, State)
If State : SetGadgetState(#CheckFixed, #False) : EndIf
EndMacro

Procedure SetFWText(*Gadget, Text.s)
SetGadgetText(*Gadget, Text)
StartDrawing(CanvasOutput(*Gadget))
DrawingFont(GetGadgetData(*Gadget))
Define Width = OutputWidth(), Height = OutputHeight(), BColor = GetGadgetColor(#Button_Back, #PB_Gadget_BackColor)
Box(0, 0, Width, Height, BColor) : DrawingMode(#PB_2DDrawing_Transparent | #PB_2DDrawing_Outlined)
Box(0, 0, Width + 1, Height, ~BColor & $FFFFFF)
DrawText((Width - TextWidth(Text)) / 2, (Height - TextHeight(Text)) / 2, Text, GetGadgetColor(#Button_Front, #PB_Gadget_BackColor))
StopDrawing()
EndProcedure

Macro RequestColor(ReceiverID, HTitle) ; Pseudo-procedure.
Define NewColor.l = ColorRequesterEx(System\SetupWindow, GetGadgetColor(ReceiverID, #PB_Gadget_BackColor), HTitle)
If NewColor <> -1 : SetGadgetColor(ReceiverID, #PB_Gadget_BackColor, NewColor) 
SetFWText(#FontField, GetGadgetText(#FontField)) : EndIf
EndMacro

Macro GetSimpleSel(Gadget, Min , Max) ; Pseudo-procedure.
SendMessage_(GadgetID(Gadget), #EM_GETSEL, @Min, @Max)
EndMacro

Macro SetSimpleSel(Gadget, Min , Max) ; Pseudo-procedure.
SendMessage_(GadgetID(Gadget), #EM_SETSEL, Min, Max)
EndMacro

Procedure NormalizeSpin(SpinID)
Define Offset, SStart, SEnd, Text.s = GetGadgetText(SpinID)
GetSimpleSel(SpinID, SStart, SEnd) ; Получаем выделение.
Define Try = Val(Text)
If Try < 0 And GetGadgetAttribute(SpinID, #PB_Spin_Minimum) >= 0
SetGadgetState(SpinID, -Try)
Else : SetGadgetState(SpinID, GetGadgetState(SpinID))
EndIf : Offset = (Len(Text) - Len(GetGadgetText(SpinID)))
If Offset > SStart : SStart = 0 : SEnd = 0
Else : SStart - Offset : SEnd - Offset
EndIf : SetSimpleSel(SpinID, SStart, SEnd)
EndProcedure

Procedure ContainerCallback(hWnd, Message, wParam, lParam)   
If Message = #WM_LBUTTONDBLCLK : RequestColor(GetProp_(hWnd, "Gadget"), GetProp_(hWnd, "ChooserTitle")) : EndIf
ProcedureReturn ChainOldCB()
EndProcedure

Macro AcceptOptions() ; Pseudo-procedure.
; -Accepting-
ForEach System\Flagi() : Define *Flag.OptionFlag = System\Flagi()
*Flag\Data\I = GetGadgetState(*Flag\Gadget) : Next
SetGadgetColor(#ClipList, #PB_Gadget_BackColor, GetGadgetColor(#Button_Back, #PB_Gadget_BackColor))
SetGadgetColor(#ClipList, #PB_Gadget_FrontColor, GetGadgetColor(#Button_Front, #PB_Gadget_BackColor))
If GetGadgetState(#CheckMaxSize) : System\Options\SizeLimit = GetGadgetState(#SizeSpin)
Else : System\Options\SizeLimit = 0 ; Ставим лимит в 0, если не указано.
EndIf
If GetGadgetState(#CheckLimit) ; Лимит вхождений.
System\Options\ListMax = GetGadgetState(#LimitSpin) 
System\Options\HPreserve = GetGadgetState(#CheckPreserve)
Define Nodes = ListSize(System\ClipList()) 
EnterCritical()
PrepareDelita()
CheckMaximum() ; Убиваем "лишние" ноды.
UseDelita()
LeaveCritical()
Else : System\Options\ListMax = 0 ; Без лимита...
System\Options\HPreserve = #False ; ...Но и без пресервов.
EndIf : UpdateTitle() :           ; Как-то так...
; Теперь шрифты и самое сложное...
System\ListFont = GetGadgetData(#FontField) ; Сразу ставим так.
System\ListFont = LoadHQFont(#fListFont, GetFontName(System\ListFont), GetFontHeight(System\ListFont), GetFontStyle(System\ListFont))
SetGadgetFont(#ClipList, System\ListFont)
; Иконки-иконочки...
If System\Options\UseTray ; Если используется иконка...
If IsWindowVisible_(System\MainWindow) ; Если окно видимо...
If     IsSysTrayIcon(#TrayIcon) = #False And System\Options\FixedTray : AddTray()
ElseIf IsSysTrayIcon(#TrayIcon) And System\Options\FixedTray = #False : RemoveSysTrayIcon(#TrayIcon)
EndIf : EndIf : DisableDebugger ; Убиваем иконку:
Else : RemoveSysTrayIcon(#TrayIcon) : EndIf
EnableDebugger
SendNotifyMessage_(#HWND_BROADCAST, System\UpdateMsg, 0, 0)
; -Saving-
WriteAllPrefs() ; Схороняем все.
EndMacro

Macro FormatFontString(FName, FSize, FFlags) ; PSeudo-procedure
FName + " : " + StrF(FSize, 1) + " : " + Str(FFlags)
EndMacro

Procedure DefaultFrontColor()
ProcedureReturn GetSysColor_(#COLOR_WINDOWTEXT)
EndProcedure

Procedure DefaultBackColor()
ProcedureReturn GetSysColor_(#COLOR_WINDOW)
EndProcedure

Procedure ShowFont(*VFont)
SetGadgetData(#FontField, *VFont) 
SetFWText(#FontField, FormatFontString(GetFontName(*VFont), GetFontHeight(*VFont), GetFontStyle(*VFont)))
EndProcedure

Macro ResetOptions() ; Partializer.
ForEach System\Flagi() : Define *Flag.OptionFlag = System\Flagi()
SetGadgetState(*Flag\Gadget, *Flag\DefVal) : Next
SetGadgetState(#CheckLimit     , #False) : SetGadgetState(#CheckMaxSize   , #False) 
SetGadgetColor(#Button_Front   , #PB_Gadget_BackColor, DefaultFrontColor())
SetGadgetColor(#Button_Back    , #PB_Gadget_BackColor, DefaultBackColor())
CheckDump() : CheckLimit() : CheckSizeLimit() : CheckTray() : ShowFont(GetDefaultFont())
EndMacro

Macro BanishOptions() ; Partializer
HideWindow(#SettingsWindow, #True)
DisableGadget(#Button_Options, #False)
SetForegroundWindow_(System\MainWindow)
SAG()
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure RequestFontSmart(*OldFont)
ProcedureReturn FontRequester(GetFontName(*OldFont), GetFontHeight(*OldFont), 0, 0, GetFontStyle(*OldFont)) ; Добавить стили !
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro OptionsLoop() ; PArtializer
Select System\GUIEvent\Type
Case #PB_Event_Gadget
; -----------
Select System\GUIEvent\Gadget
Case #LimitSpin, #SizeSpin  ; Проверяем ввод:
If System\GUIEvent\SubType = #PB_EventType_Change : NormalizeSpin(System\GUIEvent\Gadget) : EndIf
Case #Button_Font : If RequestFontSmart(GetGadgetData(#FontField))  ; Запрашиваем.
Define *FID =  LoadHQFont(#fTmpListFont,SelectedFontName() , SelectedFontSize(), SelectedFontStyle())
ShowFont(LoadHQFont(#fTmpListFont,SelectedFontName() , SelectedFontSize(), SelectedFontStyle()))
EndIf
Case #CheckLimit    : CheckLimit()
Case #CheckUseDump  : CheckDump()
Case #Button_Reset  : ResetOptions()
Case #Button_Cancel : BanishOptions()
Case #CheckMaxSize  : CheckSizeLimit()
Case #CheckTray     : CheckTray()
Case #Button_Accept : AcceptOptions() : BanishOptions()
EndSelect
; -----------
Case #PB_Event_Menu ; Реализуем стандарт.
Select EventMenu()  ; Escape/Enter
Case #bCancel       : BanishOptions()
Case #bAccept       : AcceptOptions() : BanishOptions()
EndSelect ; ...Ну и последнее:
Case #PB_Event_CloseWindow : BanishOptions()
EndSelect
EndMacro

Procedure BringOptions() ; Partializer.
; Preparations.
ForEach System\Flagi() : Define *Flag.OptionFlag = System\Flagi()
SetGadgetState(*Flag\Gadget, *Flag\Data\I) : Next
SetGadgetState(#LimitSpin      , System\Options\ListMax)
SetGadgetState(#CheckLimit     , System\Options\ListMax)
SetGadgetState(#SizeSpin       , System\Options\SizeLimit)
SetGadgetState(#CheckMaxSize   , System\Options\SizeLimit)
SetGadgetState(#CheckPreserve  , System\Options\HPreserve)
SetGadgetColor(#Button_Front, #PB_Gadget_BackColor, GetGadgetColor(#ClipList, #PB_Gadget_FrontColor))
SetGadgetColor(#Button_Back, #PB_Gadget_BackColor, GetGadgetColor(#ClipList, #PB_Gadget_BackColor))
; Show-up.
CheckDump() : ShowFont(System\ListFont)
CheckLimit() : CheckSizeLimit()
HideWindow(#SettingsWindow, #False)
DisableGadget(#Button_Options, #True)
SetForegroundWindow_(System\SetupWindow)
EndProcedure
;}
;{ --Viewports management--
Procedure MimicList(*Gadget)
SetGadgetFont(*Gadget, SendMessage_(System\ListID, #WM_GETFONT, 0, 0))
SetGadgetColor(*Gadget, #PB_Gadget_FrontColor, GetGadgetColor(#ClipList, #PB_Gadget_FrontColor))
SetGadgetColor(*Gadget, #PB_Gadget_BackColor, GetGadgetColor(#ClipList, #PB_Gadget_BackColor))
EndProcedure

Procedure ResetMimicry(*Gadget)
SetGadgetFont(*Gadget, GetDefaultFont())
SetGadgetColor(*Gadget, #PB_Gadget_FrontColor, DefaultFrontColor())
SetGadgetColor(*Gadget, #PB_Gadget_BackColor, DefaultBackColor())
EndProcedure

Procedure ScaleViewPort(*VPort.ViewPort, Factor.f) ; Pseudo-procedure.
Define *WID = *VPort\WindowID, Width.f = WindowWidth(*WID), Height.f = WindowHeight(*WID)
ResizeWindow(*WID, #PB_Ignore, #PB_Ignore, Width * Factor, Height * Factor)
Define X = WindowX(*WID) - (WindowWidth(*WID) - Width) / 2, Y = WindowY(*WID) - (WindowHeight(*WID) - Height) / 2
ResizeWindow(*WID, X, Y, #PB_Ignore, #PB_Ignore)
EndProcedure
; -----------------------------------------
Procedure MenuItemEx(ItemIdx, ItemText.s, SCText.s = "")
If SCText : ItemText + #TAB$ + SCText : EndIf : MenuItem(ItemIdx, ItemText)
EndProcedure
; -----------------------------------------
Macro CtrlSC(SC) ; Pseudo-procedure.
IIFS(Bool(SC <> ""), "Ctrl+" + SC, "")
EndMacro

Macro MenuItemCtrl(ItemIdx, ItemText, SC) ; Partializer
MenuItemEx(ItemIdx, ItemText, CtrlSC(SC))
EndMacro

Macro SelAndRaw(SubGroup = "", RawDesc = "Copy->TXT", HotKey = "R") ; Partializer.
MenuItemCtrl(#vSelectAll, "Select All", "A")
CompilerIf SubGroup <> "" : OpenSubMenu(Subgroup) ; Если нужно - пакуем контектное меню.
CompilerEndIf : MenuItemCtrl(#vCopyRaw, RawDesc, HotKey) 
EndMacro

Macro GetSel(Gadget, SelAccum) ; Pseudo-procedure.
SendMessage_(GadgetID(Gadget), #EM_EXGETSEL, 0, @SelAccum)
EndMacro

Macro SetSel(Gadget, SelAccum) ; Pseudo-procedure.
SendMessage_(GadgetID(Gadget), #EM_EXSETSEL, 0, @SelAccum)
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro ActualSelection(SelObj) ; Pseudo-procedure.
SelObj\GetStringProperty("type") <> "None"
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure ViewPortMenu(*VPort.ViewPort)
#SubHMTL = "Copy->{...}" ; Спец. константа
If *VPort And *VPort\DataNode ; Такая вот проверка. Ну, может оно и верно.
Define Type = *VPort\DataNode\DataType, Sel.CHARRANGE
Define State = GetWindowState(*VPort\WindowID)
CreatePopupMenu(#mVPMenu)
If State = #PB_Window_Normal ; Вне полноэкранного размера опций явно больше.
MenuItemEx(#vMaximize,  "Maximize")
MenuItemEx(#vCenterWin, "Win->Center")
Else : MenuItemEx(#vReturnSize,  "Restore Size") ; Унылый возврат прежнего размера.
EndIf : MenuBar() 
Select Type   ; Контекстные пункты...
Case #CF_BITMAP, #CF_ENHMETAFILE      ; Void.
Case #CF_HTML : Define *SelObj.COMateObject = System\ViewPort\WebObject\GetObjectProperty("Document\Selection")
If ActualSelection(*SelObj) : SelAndRaw(#SubHMTL, "Copy->TXT{raw}")  : MenuItemCtrl(#vCopyMD, "Copy->TXT{tags}", "M")
MenuItemCtrl(#vCopy, "Copy->HTML", "Ins") : CloseSubMenu() ; HTML' special.
Else : SelAndRaw("", #SubHMTL, "")  : DisableMenuItem(#mVPMenu, #vCopyRaw, #True)
EndIf : MenuBar() : *SelObj\Release() ; Высвобождаем объект.
Default                               ; --Текстовые данные.
SelAndRaw() : MenuItemCtrl(#vCopy, "Copy->RTF", "Ins")
GetSel(System\ViewPort\ViewArea, sel) ; Получаем тек. выделение.
If Sel\CpMin=Sel\CpMax:DisableMenuItem(#mVPMenu,#vCopyRaw,#True):DisableMenuItem(#mVPMenu,#vCopy,#True):EndIf
MenuBar()
EndSelect ; И теперь оставшиеся (по большей части - общие) пункты:
Select Type ; Спец. случай для переноса слов:
Case #CF_TEXT, #CF_HDROP, #CF_RichText : MenuItemCtrl(#vWordWrap, "Word Wrap", "W") : MenuBar()    ; Перенос по словам.
SetMenuItemState(#mVPMenu, #vWordWrap, Bool(Not *Vport\DataNode\WrapFlag))                         ; Ставим галочку, если вдруг.
EndSelect : If Type = #CF_ENHMETAFILE : MenuItemEx(#vSaveSnap, "Render As...") : MenuBar() : EndIf ; Рендер метафайла.
MenuItemCtrl(#vSaveAs, "Save As...", "S") : MenuItemEx(#vHighLight, "Find Source")
If WindowWidth(*VPort\WindowID) = #VPortMinWidth And WindowHeight(*VPort\WindowID) = #VPortMinHeight
DisableMenuItem(#mVPMenu, #vSizeDown, #True)                              ; Отключаем уменьшение.
EndIf : DisplayPopupMenu(#mVPMenu, WindowID(*VPort\WindowID)) : EndIf     ; Отображаем результатирующее меню.
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro Transparentality(Cond, More) ; Pseudo-procedure.
If Cond : MakeTransparent(*WindowID) ; Ставим прозрачность при потере фокуса.
Else : SetOpacity(#FullAlpha, *WindowID) : More : EndIf
EndMacro

Macro StoreVPIndex(VP) ; Pseudo-procedure.
VP\PrevIndex = Str(Node2Index(VP\DataNode) + 1)
EndMacro

Macro ActualizeWordWrap(VP) ; Pseudo-procedure.
SendMessage_(GadgetID(VP\ViewArea), #EM_SETTARGETDEVICE, #Null, VP\DataNode\WrapFlag)
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure WinCOMProc(object.COMateObject, eventName.s, parameterCount, *returnValue.VARIANT) ; Callback.
Protected ObjEvent.COMateObject
If eventName = "onerror" : *returnValue\vt = #VT_BOOL : *returnValue\boolVal = #VARIANT_TRUE : EndIf 
EndProcedure

Procedure FillVPData(*VP.OpticSys)
With *VP        ; Старательно обратаываем нод.
If *VP = System\Panopticum And *VP\Actual = #True : ProcedureReturn #False : EndIf ; Выходим, если нет смысла.
Define *DataNode.ClipData = \DataNode ; Ускоряем доступ.
If *DataNode    ; Если порт вообще должен что-то показывать...
Select *DataNode\DataType ; Выбираем по типу.
Case #CF_TEXT, #CF_HDROP : SetGadgetText(\ViewArea, ExtractText(\DataNode)) : ActualizeWordWrap(*VP)
Case #CF_BITMAP          : \Img = ExtractImage(*DataNode)
Case #CF_ENHMETAFILE     : \TempMeta = ExtractMeta(*DataNode)
Case #CF_RichText        : Define *DPtr = ExtractComplexData(*DataNode)     : SetGadgetText(\ViewArea, PeekRTF(*DataNode, *DPtr))
CleanAfter(*DataNode, *DPtr)                                                : ActualizeWordWrap(*VP)  
Case #CF_HTML            : Define *HPoint = ExtractComplexData(*DataNode) ; Извлекаем данные для заполнения.
ParseHTML(*VP\WebObject, *HPoint) : CleanAfter(*DataNode, *HPoint)        ; Парсим HTML.
*VP\WinCom = *VP\WebObject\GetObjectProperty("Document\Parentwindow")     ; Получаем родительское окно.
*VP\WinCom\SetEventHandler(#COMate_CatchAllEvents, @WinCOMProc(), #COMate_OtherReturn)
EndSelect : If *VP = System\Panopticum : *VP\Actual = #True : EndIf       ; Ставим флажок актуальности.
EndIf
EndWith
EndProcedure

Macro Splitteraze(FB, Win = #MainWindow) ; Partializer
If FaultyOS() : ResizeGadget(FB\Splitter, #PB_Ignore, #PB_Ignore, WindowWidth(Win), WindowHeight(Win))
Define FullWidth = GadgetWidth(FB\Container), FullHeight = GadgetHeight(FB\Container)
Else : FullWidth = WindowWidth(Win) : FullHeight = WindowHeight(Win) : EndIf
EndMacro

Macro BanePanopticum() ; Partializer.
HideWindow(System\Panopticum\WindowID, #True) : UpdateOcular()
EndMacro

Procedure ViewportCB(*WindowID, Message, wParam = 0, lParam = 0) ; Callback.
Define *VPort.ViewPort = ExtractWP(*WindowID)
If *Vport   ; Если есть, о чем беседовать...
With *VPort ; Работаем с идентификатором.
Select Message ; Анализируем сообщение.
Case System\SizeMsg : Splitteraze(\Stabilizer, \WindowID) ; Сообщение об изменении размера:
If \Frame ; Если у области вывода исскуственная рамка...
ResizeGadget(\Frame, #PB_Ignore, #PB_Ignore, FullWidth - #VAROffset + 2, FullHeight - #VABOffset + 2)
EndIf : ResizeGadget(\ViewArea, #PB_Ignore, #PB_Ignore, FullWidth - #VAROffset, FullHeight - #VABOffset) 
ResizeInformer(\Informator) ; Подбиваем размерность.
Case #WM_ACTIVATE : Transparentality(wParam = #WA_INACTIVE, System\ViewPort = *VPort) ; Сообщение об изменении фокуса.
Case System\CloseMsg, #WM_CLOSE : If *Vport <> System\Panopticum : CloseViewPort(*VPort) : Else : BanePanopticum() : EndIf
Case System\HideMsg    : \WasVisible = IsWindowVisible_(WindowID(\WindowID)) : HideWindow(\WindowID, #True)
Case System\RestoreMsg : If \WasVisible : If *VPort = System\Panopticum : FillVPData(System\Panopticum) : EndIf ; Заполняем.
HideWindow(\WindowID, #False) : EndIf                                             ; Возвращаем, если окно правда скрывалось.
Case System\UpdateMsg  : If *VPort\MimicList ; Если порт вообще обрабатывает это сообщение...
If System\Options\MimicList : MimicList(\ViewArea) : Else : ResetMimicry(\ViewArea) : EndIf
EndIf ; Едем дальше... Теперь - 
; -------------------------------------------
Transparentality(GetActiveWindow_() <> *WindowID And System\Options\GlassWin, lParam = 0)
; -------------------------------------------
Case System\RenumMsg ; Если идет перенумерация.
SetWindowTitle(\WindowID,ReplaceString(GetWindowTitle(\WindowID),\PrevIndex,Str(Node2Index(\DataNode)+1),#PB_String_CaseSensitive,1,1))
StoreVPIndex(*VPort) ; Перенумеровываем обязательно.
EndSelect
EndIf
EndWith 
ProcedureReturn #PB_ProcessPureBasicEvents
EndProcedure

Procedure FitImage(*Area.Rect, IWidth, IHeight)
With *Area ; Обработка области.
Define Aspect.f
If IWidth > IHeight : Aspect  = IWidth  / IHeight
Define OutWidth = \Right, OutHeight = OutWidth / Aspect
If OutHeight > \Bottom : OutHeight = \Bottom : OutWidth = OutHeight * Aspect : EndIf
Else                : Aspect  = IHeight / IWidth
Define OutHeight = \Bottom, OutWidth = OutHeight / Aspect
If OutWidth > \Right : OutWidth = \Right : OutHeight = OutWidth * Aspect : EndIf
EndIf
\Left = (\Right - OutWidth) >> 1 : \Top = (\Bottom - OutHeight) >> 1
\Right = OutWidth : \Bottom = OutHeight ; Ставим размеры
EndWith
EndProcedure

Macro PaintDefinitions() ; Partializer.
Define Client.RECT, Painter.PAINTSTRUCT, *DC = BeginPaint_(*GadgetID, @Painter) ; Получаем контекст.
EndMacro

Procedure BMP_CB(*GadgetID, Message, wParam, lParam) ; Callback.
If Message = #WM_PAINT  ; Если нужно отрисовать изображение...
Define *VPort.ViewPort = ExtractWP(GetParent_(*GadgetID))
If *VPort\Img : PaintDefinitions()           ; Если есть, о чем вообще говорить...
With Client ; Работаем с клиентской областью....
Define *Src = StartDrawing(ImageOutput(*VPort\Img))
GetClientRect_(*GadgetID, @Client) ; Получаем клиентскую область.
FillRect_(*DC, Client, GetStockObject_(#BLACK_BRUSH)) ; Забиваем.
Define IWidth = OutputWidth(), IHeight = OutputHeight() ; Размеры.
FitImage(Client, IWidth, IHeight)  ; Подгоняем и рисуем:
StretchBlt_(*DC, \Left, \Top, \Right, \Bottom, *Src, 0, 0, IWidth, IHeight, #SRCCOPY)
StopDrawing() : EndPaint_(*GadgetID, @Painter) ; Освобождаем контекст.
EndIf : EndIf : ProcedureReturn ChainOldCB(*GadgetID)
EndWith
EndProcedure

Procedure META_CB(*GadgetID, Message, wParam, lParam) ; Callback.
If Message = #WM_PAINT  ; Если нужно отрисовать изображение...
Define *VPort.ViewPort = ExtractWP(GetParent_(*GadgetID))
If *VPort\TempMeta : PaintDefinitions() ; Только если есть, что там именно чертить.
GetClientRect_(*GadgetID, @Client)      ; Получаем клиентскую область.
FillRect_(*DC, Client, GetStockObject_(#WHITE_BRUSH)) ; Очистка.
FitImage(Client, *VPort\DataNode\Sizing\X, *VPort\DataNode\Sizing\Y) ; Вписываем.
Client\Right + Client\Left : Client\Bottom + Client\Top
SelectObject_(*DC, GetStockObject_(#BLACK_PEN)) : SelectObject_(*DC, GetStockObject_(#NULL_BRUSH))
PlayEnhMetaFile_(*DC, *VPort\TempMeta, Client) ; Рисуем.
Rectangle_(*DC, Client\Left, Client\Top, Client\Right, Client\Bottom) ; Рамка.
EndPaint_(*GadgetID, @Painter)          ; Освобождаем контекст.
EndIf : EndIf : ProcedureReturn ChainOldCB(*GadgetID)
EndProcedure

Procedure SplitterCallback(hWnd, Message, wParam, lParam)
Select Message : Case #WM_MOUSEMOVE, 132, 32 : ProcedureReturn 0 : EndSelect
Define Result = ChainOldCB()
Select Message
Case #WM_ERASEBKGND       : Result = 1
Case #WM_WINDOWPOSCHANGED : RedrawWindow_(hWnd, 0, 0, #RDW_UPDATENOW)
EndSelect
ProcedureReturn Result
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure BMPDragger() ; Callback.
Define *VP.ViewPort = ExtractWP(WindowID(EventWindow())) : DragImage(ImageID(*VP\Img))
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure CopyMemoryStringN(Text.s) ; Replacer.
CopyMemoryString(Text) : CopyMemoryString(#CRLF$)
EndProcedure

Macro ShakeRnd() ; Partializer.
Rnd(-#VPEntropy, #VPEntropy)
EndMacro

Procedure.s CF2ID(ClipFormat)
Define ID.s = Pref(ClipFormat)
ProcedureReturn Left(ID, Len(ID) - Len(#Postf))
EndProcedure

Procedure NavigationCB(*Gadget, Url.s) ; Callback.
If GetGadgetAttribute(*Gadget, #PB_Web_Busy) : RunProgram(URL) : EndIf ; What "??"
ProcedureReturn #False ; Запрещаем куда-либо переходить.
EndProcedure

Procedure VoidNavi(*Gadget, Url.s) ; Callback.
ProcedureReturn #False ; Запрещаем куда-либо переходить вообще.
EndProcedure

Macro VASizings() ; Partializer.
#PB_Any, #UniVPOffset, #VAUOffset, 0, 0
EndMacro

Macro Editorial(Accum = *Gadget, VPort = *Vport) ; Partializer.
Accum = EditorGadget(VASizings(),#PB_Editor_ReadOnly)
SetGadgetData(Accum, VPort)
EndMacro

Macro PlainEditorial(Accum = *Gadget, VPort = *Vport) ; Partializer.
Editorial(Accum, VPort) : MakeVPPlain(VPort)
EndMacro

Macro ImagePort(Accum = *Gadget, VPort = *Vport) ; Partializer.
Accum = ImageGadget(VASizings(), #Null, #PB_Image_Border)
SetGadgetData(Accum, VPort)
EndMacro

Macro WebPort(Accum = *Gadget, VPort = *Vport) ; Partializer.
Accum = WebGadget(VASizings() , "") : SetGadgetData(Accum, VPort)
SetGadgetAttribute(Accum, #PB_Web_NavigationCallback, @NavigationCB())   ; Запрещаем навигацию.
SetGadgetAttribute(Accum, #PB_Web_BlockPopupMenu, #True) ; Блокируем к черту стандартное меню.
VPort\Frame = ContainerGadget(#PB_Any, #UniVPOffset-1,#VAUOffset-1, 0, 0, #PB_Frame_Flat) : CloseGadgetList()
VPort\WebObject = COMate_WrapCOMObject(GetWindowLongPtr_(GadgetID(Accum), #GWL_USERDATA))
ViewPortShort(A, #vSelectAll) : ViewPortShort(C, #vCopy) : ViewPortShort(Insert, #vCopy)
EndMacro

Macro BMPort(Accum = *Gadget, VPort = *Vport)
ImagePort(Accum, VPort) : ImplyBMP(Accum)
EndMacro

Macro MetaPort(Accum = *Gadget, VPort = *Vport)
ImagePort(Accum, VPort) : ChangeCB(GadgetID(Accum), META_CB())
EndMacro

Macro MakeTextPlain(hWnd) ; Pseudo-procedure.
SendMessage_(hWnd, #EM_SETTEXTMODE, #TM_PLAINTEXT, 0)
EndMacro

Macro MakeVPPlain(ViewPort)
ViewPort\MimicList = #True : MakeTextPlain(GadgetID(ViewPort\ViewArea))
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro AKS(Wnd, BoundKey, MenuItemEx, Mod) ; Partializer.
AddKeyboardShortcut(Wnd, Mod|#PB_Shortcut_#BoundKey, MenuItemEx)
EndMacro

Macro ViewPortShort(BoundKey, MenuItemEx, Mod = #PB_Shortcut_Control) ; Partializer.
AKS(*Window, BoundKey, MenuItemEx, Mod)
EndMacro

Macro ShortControl(BoundKey, MenuItemEx, Mod = #PB_Shortcut_Control) ; Partializer.
AKS(#MainWindow, BoundKey, MenuItemEx, Mod)
EndMacro

Macro ImplyBMP(Gadget) ; Partializer.
ChangeCB(GadgetID(Gadget), BMP_CB()) : BindGadgetEvent(Gadget, @BMPDragger(), #PB_EventType_DragStart)
EndMacro

Macro RemFormat(RText) ; Partializer.
#RemPref + RText + "»"
EndMacro

Macro SetContain(Stab) ; Pseudo-procedure.
If FaultyOS() : Stab\Container = ContainerGadget(#PB_Any, 0, 0, 0, 0)  : EndIf
EndMacro
;;;;;;;;;;;;;;;;;;;;;;
Macro InitVPWindow(WinAccum, Stab, Sizer = System\VoidBorders) ; Pseudo-procedure.
Define Sizing.Rect = Sizer         ; Аккумулируем полученные данные о границах окна.
If Sizing\Right <= 0 And Sizing\Bottom <= 0 : Sizing\Right = -Sizing\Right : Sizing\Bottom = -Sizing\Bottom ; Инвертируем, да.
Define CFlag = #PB_Window_WindowCentered : Else : CFlag = #False     : EndIf ; Устанавливаем по центру только при инциализации.
If Sizing\Right  < #VPortMinWidth  : Sizing\Right  = #VPortMinWidth  : EndIf ; Проверяем на соответсвие минимальным ширине...
If Sizing\Bottom < #VPortMinHeight : Sizing\Bottom = #VPortMinHeight : EndIf ; ...и высоте...
WinAccum = OpenWindow(#PB_Any, #PB_Ignore, #PB_Ignore, Sizing\Right, Sizing\Bottom, "", #VPFlags|CFlag, System\MainWindow)
SmartWindowRefresh(WinAccum, #True) : SmartWindowRefresh(WinAccum, #True)       ; Добавочная подготовка.
WindowBounds(WinAccum, #VPortMinWidth, #VPortMinHeight, #PB_Ignore, #PB_Ignore) ; Ограничители размера.
If CFlag = 0 : PlaceWindowStrict(WinAccum, Sizing\Left, Sizing\Top)             ; Выставляем на заранее определенной позиции.
Else  : ResizeWindow(WinAccum, WindowX(WinAccum)+ShakeRnd(), WindowY(WinAccum)+ShakeRnd(), #PB_Ignore, #PB_Ignore) ; Иначе трясем.
EndIf : SetContain(Stab)
EndMacro

Macro ActivateVP(VPort) ; Pseudo-procedure.
HideWindow(VPort\WindowID, #False)
SetForegroundWindow_(WindowID(VPort\WindowID)) : System\ViewPort = VPort ; Не забывать !
SetActiveGadget(VPort\ViewArea) ; Допустим....
EndMacro

Macro GoesHTML(VP) ; Partializer.
Bool(VP\DataNode And VP\DataNode\DataType <> #CF_HTML)
EndMacro

Macro ConfigureVP(VPort) ; Pseudo-procedure.
InitInformer(VPort\Informator) : AddWindowTimer(*Window, #tRunTimer, #RunTime)    
BindWP(WindowID(*Window), VPort) : SetWindowCallback(@ViewportCB(), *Window)         ; Указываем CallBack.
SwitchStyle(WindowID(*Window), #WS_EX_LAYERED|System\XPLegacy) ; Выставляем необходимые стили.
SendMessage_(WindowID(*Window), System\UpdateMsg, 0, 0) ; Отправлем сообщение на всякий случай.
AddKeyboardShortcut(*Window,#VK_OEM_PLUS,#vSizeUp) : AddKeyboardShortcut(*Window,#VK_OEM_MINUS,#vSizeDown) ; Shortcuts:
AddKeyboardShortcut(*Window,#VK_OEM_MINUS|#PB_Shortcut_Shift,#vSizeDown)   ; Чисто для комплекта, на самом деле минус.
ViewPortShort(Add, #vSizeUp, 0) : ViewPortShort(Subtract, #vSizeDown, 0)   ; +/- к размерам окна.
ViewPortShort(S, #vSaveAs) : ViewPortShort(R, #vCopyRaw)  ; Схоронение в файл / Копирование в "чистом" виде.
ViewPortShort(M, #vCopyMD) : ViewPortShort(W, #vWordWrap) ; Копирование с разметкой / Включение и отключение переноса по словам.
SetupBuffers(VPort\Stabilizer, GoesHTML(VPort)) : BindWP(GadgetID(VPort\Stabilizer\Container), VPort)
EndMacro
;;;;;;;;;;;;;;;;;;;;;;
Procedure NameVP(*VP.ViewPort, Prefix.s = "Data viewer")
Define *DataNode.ClipData = *VP\DataNode, Title.s = Prefix + ": " : StoreVPIndex(*VP)
If *DataNode : Title + CF2ID(*DataNode\DataType) + "[" + *VP\PrevIndex + "]" ; Аккумулируем тип и индекс.
If *DataNode\Comment : Title + RemFormat(*DataNode\Comment) : EndIf : Else : Title + "NULL[#]" : EndIf
SetWindowTitle(*VP\WindowID, Title) ; Пишем название в окно.
EndProcedure

Procedure OpenViewPort(NodeIdx)
If NodeIdx <> #UnusedNode ; Если есть, что показывать...
Define *Window, *Gadget, *VPort.ViewPort, *DataNode.ClipData = Index2Node(NodeIdx), iid.IID
If *DataNode\Viewport = #Null : UnhideList() ; Если еще не ассоциировано Viewport'а.
AddElement(System\ViewPorts()) : *VPort = System\ViewPorts() : InitVPWindow(*Window, *VPort\Stabilizer, *DataNode\VPSizing) 
*VPort\WindowID = *Window : *VPort\DataNode = *DataNode                       ; Основная линковка.
NameVP(*VPort) : Select *DataNode\DataType                                    ; Выбираем по типу данных...
Case #CF_TEXT, #CF_HDROP : PlainEditorial()                                   ; --Text.
Case #CF_BITMAP      : BMPort()                                               ; --Raster image.
Case #CF_RichText    : Editorial() : RichEdit_SetInterface(GadgetID(*Gadget)) ; --Rich Text Format.
Case #CF_ENHMETAFILE : MetaPort()                                             ; --Vector image.
Case #CF_HTML        : WebPort()                                              ; --HyperText Markup.
EndSelect : *VPort\ViewArea = *Gadget : FillVPData(*VPort) : ConfigureVP(*VPort) ; Вписываем отображаемые данные.
UpdateInformer(*VPort\Informator) : *VPort\DataNode\ViewPort = *VPort            ; Вписываем данные информационной строки.
EndIf : UnhideList() : ActivateVP(*DataNode\ViewPort)  ; Активируем и показываем в любом случае, кстати.
EndIf
EndProcedure

Procedure DataNode2Area(*DataNode.ClipData)
With System\Panopticum ; Обрабатываем окно предпросмотра.
If *DataNode           ; Если нод вообще имел место быть...
Select *DataNode\DataType ; Выбираем поле по типу данных
Case #CF_TEXT, #CF_HDROP  : ProcedureReturn \PlainArea
Case #CF_BITMAP           : ProcedureReturn \BitmapArea
Case #CF_ENHMETAFILE      : ProcedureReturn \MetaArea
Case #CF_HTML             : ProcedureReturn \HTMLArea
Case #CF_RichText         : ProcedureReturn \RTFArea
EndSelect ; ...В противном случае указываем на пустую область
Else : ProcedureReturn \NoiseArea : EndIf ; Только шум, тлен и безысходность.
EndWith
EndProcedure

Procedure LinkOptics(*DataNode.ClipData)
With System\Panopticum   ; Обрабатываем окно предпросмотра.
If \DataNode <> *DataNode ; В том случае, коли линкуемся к новому ноду...
HideGadget(\ViewArea, #True) : DisposeVPData(System\Panopticum) : \Actual = #False ; Скрываем прежний гаджет.
\DataNode = *DataNode : \ViewArea = DataNode2Area(\DataNode)
HideGadget(\ViewArea, #False) : HideGadget(\Frame, Bool(Not (\DataNode And \DataNode\DataType = #CF_HTML)))
SetComposition(GadgetID(\Stabilizer\Container), GoesHTML(System\Panopticum))
Define *WID = WindowID(\WindowID)  ; Получаем, для общего упрощения.
If IsWindowVisible_(*WID) : FillVPData(System\Panopticum) : EndIf   ; Заполняем порт данными.
NameVP(System\Panopticum, "/Panopticum/") : ViewportCB(*WID, System\SizeMsg) ; Для отображения новой области вызова.
Define *HLNode = *DataNode : ConnectInformer(System\Panopticum\Informator) ; Указываем новые данные для.
EndIf
EndWith
EndProcedure

Macro NoiseGarden() ; Partializer.
If PanVisibility() And System\Panopticum\ViewArea = System\Panopticum\NoiseArea
Define *NoiseID = GadgetID(System\Panopticum\NoiseArea), NA.Rect
GetClientRect_(*NoiseID, @NA) : CreateImage(#NoiseImg, NA\Right, NA\Bottom)
Define *NoiseDC = StartDrawing(ImageOutput(#NoiseImg)), *AreaDC = GetDC_(*NoiseID)
RandomData(DrawingBuffer(), DrawingBufferPitch() * OutputHeight())
BitBlt_(*AreaDC, 0, 0, NA\Right, NA\Bottom, *NoiseDC, 0, 0, #SRCCOPY) : StopDrawing()
ReleaseDC_(*NoiseID, *AreaDC) : InvalidateRect_(GadgetID(System\Panopticum\ViewArea), #Null, 0)
EndIf
EndMacro

Macro BringPanopticum() ; Partializer.
FillVPData(System\Panopticum) : ActivateVP(System\Panopticum) : UpdateOcular()
EndMacro
;}
;{ --Menu management--
Procedure ShowTrayMenu()
Define *Node.ClipData, Item.i = #tUseNode
NewList HotSlots()
CreatePopupMenu(#mTrayMenu) ; Создаем контекстное меню для иконки.
MenuItemEx(#tShowWindow, "Show Window")   ; Показать окно.
If IsWindowVisible_(System\MainWindow)  : DisableMenuItem(#mTrayMenu, #tShowWindow, #True) : EndIf
MenuItemEx(#tOptions   , "Show Options")  ; Показать опции.
If IsWindowVisible_(System\SetupWindow) : DisableMenuItem(#mTrayMenu, #tOptions   , #True) : EndIf
MenuBar() ; Обязательный разделитель.
For I = 1 To #HotKeys ; Анализируем горячие клавиши...
If System\HotNodes[I] : AddElement(HotSlots()) : HotSlots() = I : EndIf
Next I ; Обрабатываем список:
If ListSize(HotSlots()) ; Если есть привязаннные данные.
OpenSubMenu("Use Data") ; Создаем подменю:
ForEach HotSlots() : MenuItemEx(Item, #SlotPrefix + Str(HotSlots())) ; Добавляем слот.
Define *DN.ClipData = System\HotNodes[HotSlots()]
If System\UsedNode = *DN : SetMenuItemState(#mTrayMenu, Item, #True) : EndIf
*DN\MenuSlot = Item : Item + 1    ; Маркирум слот.
Next : CloseSubMenu() ; Возвращаемся к основному.
EndIf ; Продолжаем генерацию меню:
MenuItemEx(#tClearList , "Clear List")    ; Очистить список данных.
If ListSize(System\ClipList()) = 0 : DisableMenuItem(#mTrayMenu, #tClearList, #True) : EndIf
If System\AcceptNew : MenuItemEx(#tSwitch, "Stop Tracking")
Else                : MenuItemEx(#tSwitch, "Resume Tracking")
EndIf
MenuBar() ; Обязательный разделитель.
MenuItemEx(#tTerminate , "Terminate")     ; Выйти из программы
DisplayPopupMenu(#mTrayMenu, System\MainWindow)
EndProcedure

Macro OfferFlatten(Node, Prefix = CF2ID(*DataNode\DataType)) ; Partializer.
MenuItemEx(#cFlatten, Prefix + "->STR")  : MenuBar()
If Node\Flattable = #False : DisableMenuItem(#mListMenu, #cFlatten, #True) : EndIf
EndMacro

Macro UsedNode() ; Partializer.
NodeIdx <> #UnusedNode
EndMacro

Macro DataNode() ; Partializer.
*DataNode.ClipData = Index2Node(NodeIdx)
EndMacro

Macro RenderInit(Node, Image, NewImg, SizePtr = #Null) ; Partializer.
Define *ISize.Point, *Render    ; Заголовок метафайла.
Define Area.Rect : MetaSelect() : GetMetaHeader(*MetaIDx)
If SizePtr   : *ISize = SizePtr ; Переписываем указатель размера.
Define Width = *ISize\X, Height = *ISize\Y ; Приказные размеры.
Else :         Width = Node\Sizing\X : Height = Node\Sizing\Y
EndIf ; Теперь строим изображение:
CompilerIf NewImg : Image = CreateImage(#PB_Any, Width, Height, #BitDepth)
CompilerElse : CreateImage(Image, Width, Height, 32) ; Сюда пойдет результат.
CompilerEndIf
*Render = StartDrawing(ImageOutput(Image))     ; Начинаем отрисовку.
Area\Right = Width : Area\Bottom = Height      ; Выставляем область отрисовки.
EndMacro

Procedure RenderMetafile(*DataNode.ClipData, *ImageID, *Forced.Point = #Null)
With *Node                                        ; Обрабатываем нод.
Define *Ptr.Ascii, *ToFix                         ; Данные обработки
RenderInit(*DataNode, *ImageID, #False, *Forced) : PlayEnhMetaFile_(*Render, *MetaIDx, Area) ; Рендерим метафайл.
*Ptr = DrawingBuffer() + #BitDepth / 8 : *ToFix = *Ptr + DrawingBufferPitch() * Height
For *Ptr = *Ptr To *ToFix Step SizeOf(Long) : *Ptr\A ! 255 : Next *Ptr ; Уничтожаем альфу.
StopDrawing() : MetaCheckUP()                     ; Удаляем метафайл.
ProcedureReturn *ImageID                          ; Возвращаем на всякий случай.
EndWith
EndProcedure

Procedure.s Format2Template(Format, *DataNode.ClipData = #Null)
#RasterShared = "PNG image (*.png)|*.png"
#TextPattern  = "Text file (*.txt)|*.txt"
#AllFilez     = "|All files (*.*)|*.*"
Select Format ; Выбираем по формату.
Case #CF_TEXT, #CF_HDROP : ProcedureReturn #TextPattern ; UTF-8 text.
Case #CF_HTML            : ProcedureReturn "HTML page (*.html)|*.html|"    + #TextPattern
Case #CF_RichText        : ProcedureReturn "Rich text file (*.rtf)|*.rtf|" + #TextPattern
Case #CF_ENHMETAFILE     : Define Result.s = "Enchanced meta-file (*.emf)|*.emf"
If *DataNode\Sizing\X > 0 And *DataNode\Sizing\Y > 0 ; Если оно там пригодно к рендеру...
Result + "|" + #RasterShared : EndIf : ProcedureReturn Result ; Возвращаем результат.
Case #CF_BITMAP          ; Most complex one:
ProcedureReturn #RasterShared + "|JPEG image (*.jpg)|*.jpg|JPEG2000 image (*.jp2)|*.jp2|BMP image (*.bmp)|*.bmp"
EndSelect
EndProcedure

Procedure.s Pattern2Extension(Pattern.s, PatternIdx.i)
Define Ext.s = GetExtensionPart(StringField(Pattern, (PatternIdx + 1) * 2, "|"))
If Ext <> "*" : ProcedureReturn Ext : EndIf
EndProcedure

Macro AddExtension(FName, Ext) ; Pseudo-procedure
If "." + LCase(GetExtensionPart(FName)) <> Ext : Fname + IIFS(Bool(Ext), ".", "") + Ext : EndIf
EndMacro

Procedure WriteText(Text.s) ; Replacer
WriteString(0, Text)
EndProcedure

Procedure WriteLine(Text.s) ; Replacer
WriteStringN(0, Text)
EndProcedure

Macro DialogID(NodeType, NodeIndex) ; Partializer.
CF2ID(NodeType) + "[" + Str(NodeIndex + 1) + "]"
EndMacro

Procedure.s PrettyPrinter(*Node.ClipData)
Define Plain.s = ExtractText(*Node)
With *Node
Select \DataType ; Выбираем поправки к печати по типу...
Case #CF_HDROP : ProcedureReturn ReplaceString(Plain, #LF$, #CRLF$)
Default : ProcedureReturn Plain.s ; Возвращаем как есть.
EndSelect
EndWith
EndProcedure

Macro SaveAsText(FName, Node) ; Partializer.
CreateFile(0, FName) : WriteString(0, PrettyPrinter(Node)) : CloseFile(0)
EndMacro

Macro TransformationBegin(Node, NewType) ; Partializer.
Define Marking.s = Pref(\DataType)                ; Запоминаем прежний префикс.
DeleteMapElement(System\Dictionary(), Node\DictID); Стираем запись в словаре...
LinkIDs(*DataNode, NewId)                         ; Вписываем его в словарь
Define OldID.s = CF2ID(Node\DataType)             ; Получаем старый идентификатор.
Define OldType = Node\DataType                    ; Копируем тип данных.
Node\DataType = NewType                           ; Выставляем последнее...
EndMacro

Procedure.s PurifyListing(*FList.CharExtra, *CountPtr.Integer)
; ---------
Define OutList.s, Appendix.s, FCount
While *FList\U : Select *FList\U
Case #CR ; NOP.
Case #LF : If FCount = 0 : Break : Else : FCount = 0 : EndIf ; Корректор, дабы поведение было аутеничнее.
OutList + *FList\Lense : *CountPtr\I + 1
Default  : FCount + 1 : OutList + *FList\Lense
EndSelect : *FList + SizeOf(Unicode) : Wend
*CountPtr\I + 1 ; Финальное повышение.
ProcedureReturn OutList + Appendix
; ---------
EndProcedure

Macro OpticBond()
If *DataNode = System\Panopticum\DataNode And PanVisibility() : System\Panopticum\DataNode = #Null : LinkOptics(*DataNode) : EndIf
EndMacro

Macro NodeHL() ; Partializer. 
OpticBond() : HLLine(NodeIdx)
System\LastAnalyzed = -1
EndMacro

Macro ActualizeClip() ; Partializer.
If System\UsedNode = *DataNode : RestoreData(NodeIdx, #True) : EndIf
EndMacro

Macro SwiftActualize() ; Partializer.
Define NewIdx = FindHost(*Host, #ForceMode)
If System\UsedNode = *DataNode : RestoreData(NewIdx, #True) : HLLine(NewIdx) : EndIf
EndMacro

Macro PressText() ; Partializer.
DisposeBinary(*DataNode) ; Удаляем тут, дабы было удобнее.
Define TS = StringByteLength(Text), *CMem = EncodeData(@Text, TS)
If *CMem : *DataNode\BinData = *CMem : *DataNode\CmpSize = MemorySize(*CMem) : *DataNode\TextData = ""
Else     : *DataNode\BinData = #Null : *DataNode\CmpSize = 0                 : *DataNode\TextData = Text : EndIf
EndMacro

Macro PressImage() ; Partializer.
DisposeBinary(*DataNode) ; Удаляем тут, дабы было удобнее.
Define TS = OutputSize(), *CMem = EncodeData(DrawingBuffer(), TS) ; Пытаемся пождать.
If *CMem : *DataNode\BinData = *CMem  : *DataNode\CmpSize = MemorySize(*CMem) : FreeImage(*Image)
Else     : *DataNode\BinData = *Image : *DataNode\CmpSize = 0                 : EndIf 
EndMacro

Macro RedoLine(Mark, Actual = LineCutter(TextEntry(Text))) ; Partializer.
Define DID.s = GetGadgetItemText(#ClipList, NodeIdx)
SetGadgetItemText(#ClipList, NodeIdx, Left(DID, FindString(DID, Mark) - 1) + Actual)
EndMacro

Macro RenameVP(VP = *DataNode\ViewPort) ; Partializer.
SetWindowTitle(VP\WindowID, ReplaceString(GetWindowTitle(VP\WindowID), OldID, Cf2ID(*DataNode\DataType)))
UpdateInformer(VP\Informator, VP\DataNode)
EndMacro

Macro MorphingTriplet(LAccum = TextEntry(Text)) ; Partializer.
PressText() : ActualizeClip() : RedoLine(Marking, LineCutter(LAccum))
EndMacro 

Macro MoprhingInfix(LAccum = TextEntry(Text)) ; Partializer.
MorphingTriplet(LAccum)
CacheNodeText(*DataNode, Text)               ; Указываем, что в кеше теперь иное.
EndMacro

Macro CheckTHost(TForm, IDAct = TextMD5(Text)) ; Partializer.
Define NewID.s = IDAct
Define *Host.ClipData = CheckHost(NewID)     ; Ищем, ищем...
If *Host : SwiftActualize() : ProcedureReturn #False : EndIf ; Теперь так.
TransformationBegin(*DataNode, TForm)        ; Общее начало.
EndMacro 

Macro MakeContextMenu(NI) ; !Menu main handler!
Define NodeIdx = NI
If UsedNode() ; Если есть, о чем говорить...
Define DataNode()
CreatePopupMenu(#mListMenu)
MenuItemCtrl(#cThrowUp,   "Throw Up"   , "PgUp")
MenuItemCtrl(#cThrowDown, "Throw Down" , "PgDn")
MenuBar() : OpenSubMenu("Bind Node")
For I = 1 To #Hotkeys : MenuItemEx(#cBindData + I, "Ctrl+" + Str(I)) ; Проходимся по всем нодам.
If *DataNode\Hotkey = I : SetMenuItemState(#mListMenu, #cBindData + I, #True) : EndIf
Next I : CloseSubMenu() ; ...А теперь продолжаем нашу рутину:
MenuItemCtrl(#cViewData, "View Data", "V")
MenuItemCtrl(#cSaveAs  , "Save As...", "S")
MenuBar()
Select *DataNode\DataType
Case #CF_RichText, #CF_HTML, #CF_HDROP : OfferFlatten(*DataNode, CF2ID(*DataNode\DataType))
Case #CF_ENHMETAFILE : MenuItemEx(#cRender, "META->BMP") : MenuBar()
If *DataNode\Sizing\X<=0 And *DataNode\Sizing\Y<=0:DisableMenuItem(#mListMenu,#cRender,1):EndIf
Case #CF_TEXT        : MenuItemEx(#cListerate, "STR->DIR") : MenuItemCtrl(#cQWERTYSwap, "Rus<->Lat", "Q") : MenuBar()
EndSelect ; Продолжаем...
MenuItemCtrl(#cSetcomment, "Set Remark", "R") : MenuItemEx(#cRemove, "Remove", "Del")
If NodeIdx = 0 : DisableMenuItem(#mListMenu, #cThrowUp, #True) : EndIf ; Убиваем у первого.
If NodeIdx = CountGadgetItems(#ClipList) - 1            ; Если это последний элемент...
DisableMenuItem(#mListMenu, #cThrowDown, #True)         ; Бросать тоже нельзя.
EndIf : DisplayPopupMenu(#mListMenu, System\MainWindow) ; Отображаем результат.
EndIf
EndMacro

Procedure SaveAs(NodeIdx.i, *ForcedRes.Point = #Null) ; !Menu handler.
If UsedNode() ; Если есть, что сохранять...
Define Idx, FileName.s, DataNode(), NoPNG = #False, *UCData
If *ForcedRes : Define Template.s = #RasterShared + #AllFilez                         ; Шаблон рендера.
Else : Template = Format2Template(*DataNode\DataType, *DataNode) + #AllFilez          ; Получаем необходимый шаблон.
EndIf : If *DataNode\DataType = #CF_ENHMETAFILE And FindString(Template, "PNG", 1)=0 : NoPNG = 10 : EndIf
Define NodeID.s = DialogID(*DataNode\DataType,NodeIdx)                                ; Получаем идентификатор.
If *ForcedRes : NodeID + " @ " + Str(*ForcedRes\X) + "x" + Str(*ForcedRes\Y) : EndIf  ; Форсим размеры, если просят.
If *DataNode\DataType = #CF_BITMAP : Idx = System\RasterType : Else : Idx = 0 : EndIf ; Уточняем, что сохраняется растр.
System\LockedNode = *DataNode          ; Схороняем для валидации.
TakeFocus() : UnhideList()             ; Дабы все отображалось.
FileName = Trim(SaveFileRequester("Locate saving destination for "+NodeID+"..."+#CRLF$,NodeID,Template,Idx))
; ----------------------------------------------------
If FileName : Define Pattern = SelectedFilePattern()         ; Сохраняем позицию о шаблоне расширений.
AddExtension(FileName, Pattern2Extension(Template, Pattern)) ; Добиваем расширение. На проверку, не иначе.
If FileSize(FileName) => 0 And WarnBox("are you sure to overwrite '" + GetFilePart(FileName) + "' ?") = #PB_MessageRequester_No
DelayReturn() : EndIf                  ; Спасение от перезаписи.
Else : DelayReturn() : EndIf           ; Если ничего не выбрали - сходу на выход.
; ----------------------------------------------------
ResumeFocus() : ReviseReturn()         ; Дабы фокус не совсем уж сбивался...
If System\LockedNode <> #Null          ; Если за это время нод не убился циклом....
Define Pattern = SelectedFilePattern() ; Сохраняем позицию.
Select *DataNode\DataType              ; Выбираем по формату...
Case #CF_TEXT        : SaveAsText(FileName, *DataNode)                            ; Текстовый файл.
Case #CF_RichText    : If Pattern <> 1 : *UCData = ExtractComplexData(*DataNode)  ; Файл RTF.
CreateFile(0, FileName) : WriteData(0, *UCData, MemorySize(*UCData)) : CloseFile(0) ; Схороняем.
CleanAfter(*DataNode, *UCData)                                                    ; Убиваем буфер
Else : SaveAsText(FileName, *DataNode) : EndIf                                    ; Сохраняем как простой текст.
Case #CF_HTML        : If Pattern <> 1 : *UCData = ExtractComplexData(*DataNode)  ; Страница HTML.
CreateFile(0, FileName) : WriteString(0, FindBody(*UCData)) : CloseFile(0)        ; Схороняем.
CleanAfter(*DataNode, *UCData)                                                    ; Убиваем буфер
Else : SaveAsText(FileName, *DataNode) : EndIf                                    ; Сохраняем как простой текст.
Case #CF_HDROP       : SaveAsText(FileName, *DataNode)                            ; Текстовый листинг.
;--------------------------------------
Case #CF_ENHMETAFILE : If *ForcedRes = #Null                                      ; Стандартное сохранение.
Select Pattern                                                                    ; Анализируем целевой формат.
Case 1 + NoPng : RenderMetafile(*DataNode, #TempImage)                            ; Рендерим метафайл в BMP.
SaveImage(#TempImage, FileName, #PB_ImagePlugin_PNG) : FreeImage(#TempImage)      ; Записываем полученное изображение.
Default                                                                           ; Сохраняем, что называется, as is.
CreateFile(0, FileName) : MetaSelect()                                            ; Аллокация (мета)файлов.
GetMetaHeader(*MetaIDX) : GetMetaBits(*MetaIDX) : WriteData(0, *RawData, *DataNode\DataSize) ; Вписываем все.
FreeMemory(*RawData) : MetaCheckUP() : CloseFile(0)                               ; Деаллокация (мета)файлов.
EndSelect ; Теперь рендер...
Else : RenderMetafile(*DataNode, #TempImage, *ForcedRes) : SaveImage(#TempImage, FileName, #PB_ImagePlugin_PNG)
FreeImage(#TempImage)                                                             ; Высвобождаем память временного изображения.
EndIf     ; Продолжаем.
;--------------------------------------
Case #CF_BITMAP : If Pattern <> 4 : System\RasterType = Pattern : EndIf ; Raster images.
Define Format  : ImgSelect() ; Теперь так.
Select Pattern ; Вновь анализируем формат и выводим отуда плагин:
Case 0  : Format = #PB_ImagePlugin_PNG
Case 1  : Format = #PB_ImagePlugin_JPEG
Case 2  : Format = #PB_ImagePlugin_JPEG2000
Case 3  : Format = #PB_ImagePlugin_BMP
Default : Format = #PB_ImagePlugin_PNG
EndSelect : SaveImage(ImgIDx, FileName, Format) : ImgCheckUP() ; Будет так, все одно лучше.
EndSelect ; Выодим ошибку, если вдруг за это время уже снесло:
Else : ErrorBox("unable to save data from deleted node !" + #ReconMsg) : ResumeFocus()
EndIf : EndIf
EndProcedure

Macro SaveScaled(ViewPort, Index) ; !Menu handler.
If USedNode() ; Если есть, что сохранять...
Define Scale.Point, Area.rect ; Данные областей.
Define DataNode()
GetClientRect_(GadgetID(ViewPort\ViewArea), Area)
FitImage(Area, *DataNode\Sizing\X, *DataNode\Sizing\Y)
Scale\X = Area\Right : Scale\Y = Area\Bottom
SaveAs(Index, Scale) ; Отправляем на сохранение.
EndIf
EndMacro

Macro SendRenum(VP) ; Pseudo-procedure.
ViewportCB(WindowID(VP\WindowID), System\RenumMsg)
EndMacro

Procedure SwapNodes(Index1, Index2) ; !Menu handler.
Define *Node1.ClipData = Index2Node(Index1) ; First node.
Define *Node2.ClipData = Index2Node(Index2) ; Second node.
If *Node1 = #Null Or *Node2 = #Null : ProcedureReturn : EndIf
Define Text1.s = GetGadgetItemText(#ClipList, Index1) ; First element's text.
Define Text2.s = GetGadgetItemText(#ClipList, Index2) ; Second element's text.
SetGadgetItemData(#ClipList, Index1, *Node2) ; Пишем второй нод в первую строку.
SetGadgetItemData(#ClipList, Index2, *Node1) ; Пишем первый нод во вторую строку.
SetGadgetItemText(#ClipList, Index1, Text2)  ; Текст со второго в первый.
SetGadgetItemText(#ClipList, Index2, Text1)  ; Текст с первого во второй.
SwapElements(System\ClipList(), *Node1, *Node2) : HLLine(Index2) ; Обмен списка и уточням выделение.
If *Node1\ViewPort : SendRenum(*Node1\ViewPort) : EndIf : If *Node2\ViewPort  : SendRenum(*Node2\ViewPort)   : EndIf
If System\Panopticum\DataNode = *Node1 Or System\Panopticum\DataNode = *Node2 : SendRenum(System\Panopticum) : EndIf
EndProcedure

Procedure RecommentVP(*VP.ViewPort, NewRem.s)
Define *WinID = *VP\WindowID, PrefCaption.s = GetWindowTitle(*WinID)
If *VP\DataNode\Comment : PrefCaption = Left(PrefCaption, FindString(PrefCaption, #RemPref) - 1) : EndIf 
If NewRem : PrefCaption + RemFormat(NewRem) : EndIf : SetWindowTitle(*WinID, PrefCaption)
EndProcedure

Macro Reinform(BlockSrc) ; Partializer.
UpdateInformer(BlockSrc\Informator, PseudoNode)
EndMacro

Procedure AskNodeComment(NodeIdx) ; !Menu handler.
If UsedNode() ; Если есть, что комментировать...
Define DataNode() : TakeFocus() : System\LockedNode = *DataNode ; Сохраняем фокус.
UnhideList()                                                    ; Дабы все отображалось.
Define Remark.s = Trim(InputRequester(#Title, #InputMsg + DialogID(*DataNode\DataType, NodeIDx) + ":", *DataNode\Comment))
ResumeFocus()                                                   ; На всякий случай, из-за возможных ошибок.
If System\LockedNode = #Null : ErrorBox("unable to set remark for deleted node !" + #ReconMsg) ; Сообщение об ошибке.
ResumeFocus() : ElseIf Remark <> *DataNode\Comment ; Если необходимо выставить новую ремарку ноду...
;;;;;;;;;;;;;;;;
Define PseudoNode.ClipData : CopyStructure(*DataNode, PseudoNode, ClipData)         : PseudoNode\Comment = Remark ; Выставляем новое значение.
If *DataNode\ViewPort                     : RecommentVP(*DataNode\ViewPort, Remark) : Reinform(*DataNode\ViewPort) : EndIf ; Правим уже.
If System\Panopticum\DataNode = *DataNode : RecommentVP(System\Panopticum, Remark)  : Reinform(System\Panopticum)  : EndIf ; Правим и тут тоже.
CopyStructure(PseudoNode, *DataNode, ClipData) : Reinform(System) ; Выставляем окну исправленный заголовок.
;;;;;;;;;;;;;;;;
EndIf
EndIf
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure ProduceListing(NodeIdx) ; !Menu handler.
If UsedNode() ; Если есть, что преобразовывать...
Define DataNode(), Text.s = ExtractText(*DataNode) 
With *DataNode                               ; Обрабатываем нод.
Define Counter.i : Text = PurifyListing(@Text, @Counter) ; Репарсим и считаем.
CheckTHost(#CF_HDROP, FilesID(Text))         ; Если есть - выходим.
\Sizing\X = Len(Text) : \Sizing\Y = Counter  ; Записываем подсчеты.
Define FTxt.s = FilesText(Counter, Text) : MoprhingInfix(FTxt)
; -------
If *DataNode\ViewPort                        ; Если требуется обновление области вывода...
SetGadgetText(\ViewPort\ViewArea, Text)      ; Для пользователя ничего не меняет, но фактички...
RenameVP() : EndIf : NodeHL()                ; Ставим курсор на этот элемент.
; -------
EndWith
EndIf 
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure SwapLayOut(NodeIdx) ; !Menu handler.
If UsedNode() ; Если есть, что перегонять...
Define DataNode()
With *DataNode                               ; Обрабатываем, вот да.
If \DataType = #CF_TEXT                      ; Придется так, раз уж клавиша горячая.
Define Text.s = ExtractText(*DataNode)       ; "Вытаскиваем" текст из нода.
XLat(@Text, @System\Layout)                  ; Инвертируем раскладку.
CheckTHost(#CF_TEXT)                         ; Если есть - выходим.
MoprhingInfix()       ; Правим стандартные данные.
If *DataNode\ViewPort ; Если требуется обновление области вывода...
SetGadgetText(\ViewPort\ViewArea, Text)      ; Переставляем кодировку. Ну, якобы.
RenameVP() : EndIf : NodeHL() : ProcedureReturn #True : EndIf ; Ставим выделение.
EndWith
EndIf
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure ComplexText2STR(NodeIdx) ; !Menu handler.
If UsedNode() : Define DataNode()            ; Если есть, что преобразовывать...
With *DataNode                               ; Обрабатываем нод.
Define Text.s = PrettyPrinter(*DataNode)     ; "Вытаскиваем" текст из нода.
If Text                                      ; Это важно. Без текста делать тут нечего.
CheckTHost(#CF_TEXT)                         ; Если есть - выходим.
*DataNode\Sizing\X = Len(Text) : *DataNode\DataSize = StringByteLength(Text) ; Размерности.
MorphingTriplet()     ; Правим содержимое буфера обмена.
If *DataNode\ViewPort ; Если требуется обновление области вывода...
Select OldType        ; Анализируем область вывода.
Case #CF_RichText     ; Если использовалась обычная...
RichEdit_Release(RichComObject(Str(GadgetID(\ViewPort\ViewArea)))) ; Освобождаем.
MakeVPPlain(*DataNode\ViewPort)                     ; Убираем поддержку RTF.
SetGadgetText(\ViewPort\ViewArea, Text)             ; Выставляем простой текст.
ViewportCB(WindowID(\ViewPort\WindowID), System\UpdateMsg) ; Уточняем.
Case #CF_HDROP : SetGadgetText(\ViewPort\ViewArea, Text)   ; Переставление не обязательно, но...
Case #CF_HTML  : DisposeVPData(\ViewPort) : WebDisposal(\ViewPort) ; ...Тут чуть сложнее, но тем не менее:
FreeGadget(\ViewPort\ViewArea) : UseGadgetList(GadgetID(\Viewport\Stabilizer\Container))
FreeGadget(\ViewPort\Frame)    : PlainEditorial(\ViewPort\ViewArea, \ViewPort) : \ViewPort\Frame = #Null
SetGadgetText(\ViewPort\ViewArea, Text) : ViewportCB(WindowID(\ViewPort\WindowID), System\UpdateMsg) ; Уточняем.
SetComposition(WindowID(\ViewPort\WindowID), #True) : ViewportCB(WindowID(\ViewPort\WindowID), System\SizeMsg) 
EndSelect                                           ; Ну и, наконец...
RenameVP() : EndIf : NodeHL()                       ; Ставим курсор на этот элемент.
EndIf : EndWith : EndIf
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure MetaFile2BMP(NodeIdx) ; !Menu handler.
If UsedNode() ; Если есть, что преобразовывать...
Define Color.l = ColorRequesterEx(System\MainWindow, System\RenderBack, @"Select rendering background:")
If Color = -1 : ProcedureReturn : EndIf : System\RenderBack = Color ; Никакого рендера без цвета.
Define *Image, DataNode()
With *DataNode                            ; Обрабатываем нод.
RenderInit(*DataNode, *Image, #True)      ; Готовим рендер.
Box(0, 0, Width, Height, Color)           ; Пишем задний фон.
PlayEnhMetaFile_(*Render, *MetaIDx, Area) ; Рендерим метафайл.
CheckTHost(#CF_BITMAP, ImageMD5())        ; Если есть - выходим.
MetaCheckUP()                             ; Уничтожаем прежние данные.
RedoLine(Marking, ImageText()) : PressImage() ; Регистрация.
StopDrawing() : ActualizeClip()               ; Правим содержимое буфера обмена. 
If *DataNode\ViewPort                     ; Если требуется обновление области вывода...
SetWindowLongPtr_(GadgetID(*DataNode\ViewPort\ViewArea), #GWL_WNDPROC, @BMP_CB())
CleanUpMeta(*DataNode\ViewPort\DataNode, *DataNode\ViewPort\TempMeta) ; Выкидываем старый метафайл.
*DataNode\ViewPort\Img = ExtractImage(*DataNode)                      ; Перерисовываем.
*DataNode\ViewPort\TempMeta = #Null                                   ; Ставим на всякий случай.
BMP_CB(GadgetID(*DataNode\ViewPort\ViewArea), #WM_PAINT, 0, 0) ; Принудительная перерисовка.
RenameVP() : EndIf : NodeHL()                                  ; Ставим курсор на этот элемент.
EndWith
EndIf
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure ThrowNode(NodeIdx, Mode = #PB_List_Last, TNode = #Null) ; !Menu handler.
If UsedNode() And ListSize(System\ClipList()) > 1 And (NodeIdx > 0 Or Mode <> #PB_List_First) And
(NodeIdx + 1 < ListSize(System\ClipList()) Or Mode <> #PB_List_Last) ; Если есть, что кидать... Ну и проверка, да.
Define DataNode()
ChangeCurrentElement(System\ClipList(), *DataNode) ; Ставим позицию.
MoveElement(System\ClipList(), Mode, TNode)        ; Просто двигаем.
; Обрабатываем GUI-представление списка нодов:
Define NewIdx = ListIndex(System\ClipList())    ; Новый индекс.
Define NodeText.s = GetGadgetItemText(#ClipList, NodeIdx)
RemoveGadgetItem(#ClipList, NodeIdx)            ; Удаляем элемент списка.
AddGadgetItem(#ClipList, NewIdx, NodeText)      ; Ставим в новое место.
SetGadgetItemData(#ClipList, NewIDx, *DataNode) ; Пишем указатель
; Теперь проверяем просмотрщики:
If *DataNode\ViewPort:ViewportCB(WindowID(*DataNode\ViewPort\WindowID),System\RenumMsg):EndIf
RenumSignal(NodeIdx) : HLLine(NewIdx) ; Ставим выделение.
EndIf
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro NotHTMLView(VP = System\ViewPort) ; Partializer
(VP\DataNode\DataType <> #CF_HTML)
EndMacro

Procedure SelectAllHTML(*VPort.ViewPort)
If Not NotHTMLView() ; Проверяем, дабы было не повадно.
With *Vport
\WebObject\Invoke("Document\body\CreateTextRange()\Select()")
EndWith
EndIf
EndProcedure

Procedure HTMLAsRaw(*VPort.ViewPort)
If Not NotHTMLView() ; Проверяем, дабы было не повадно.
Define Plain.s = GetGadgetItemText(*VPort\ViewArea, #PB_Web_SelectedText)
If Plain : SetClipboardText(Plain) : EndIf
EndIf
EndProcedure

Procedure HTMLasHTML(*VPort.ViewPort, TextForm = #False)
With *Vport
If Not NotHTMLView() ; Проверяем, дабы было не повадно.
Define *SelObject.COMateObject = \WebObject\GetObjectProperty("Document\Selection")
Define *RangeObject.COMateObject = *SelObject\GetObjectProperty("CreateRange()")
; =-----------
If *RangeObject ; Если все прошло успешно.
If ActualSelection(*SelObject) ; Проверяем, вдруг там ничего не выделено.
If TextForm = #False  : ClipHTML(*RangeObject\GetStringProperty("htmlText"), *RangeObject\GetStringProperty("Text"))
Else : SetClipboardText(*RangeObject\GetStringProperty("htmlText")) : EndIf ; Копируем как набор тегов.
EndIf : *RangeObject\Release() : *SelObject\Release() : EndIf : EndIf ; Высвобождаем.
; =-----------
EndWith
EndProcedure

Procedure RTFASRaw(*VPort.ViewPort)
If IsTextNode(*VPort\DataNode) ; Проверяем, вдруг оно не RTF.
Protected marked.CHARRANGE , Txt.s, *RED = GadgetID(*VPort\ViewArea)
SendMessage_(*RED, #EM_EXGETSEL, 0, @marked) 
If (marked\cpMax - marked\cpMin) : Txt = Space(1 + marked\cpMax - marked\cpMin) 
SendMessage_(*RED, #EM_GETSELTEXT, 0, @Txt) : SetClipboardText(txt) : EndIf
EndIf
EndProcedure

Procedure Cursor2Node(*Pos.Point = #Null) ; 2be redone.
Define Area.Rect, Cursor.Point, ItemWidth, Item
With Area ; Обрабатываем область...
GetClientRect_(System\ListID, @Area)
If *Pos : CopyStructure(*Pos, @Cursor, Point)
Else : GetCursorPos_(@Cursor) : ClientToScreen_(System\ListID, @Area) : EndIf
Item      = SendMessage_(System\ListID, #LB_GETTOPINDEX, 0, 0)
ItemWidth = SendMessage_(System\ListID, #LB_GETITEMHEIGHT, Item, 0)
Item + (Cursor\Y - \Top) / ItemWidth
ProcedureReturn Index2Node(Item)
EndWith
EndProcedure

Macro ListTip(NewText = "") ; Pseudo-procedure.
If System\ListTooltip <> NewText : System\ListTooltip = NewText : GadgetToolTip(#ClipList, NewText) : EndIf
EndMacro

Procedure DataListCallback(hWnd, Message, wParam, lParam)
Select Message
Case #WM_RBUTTONDOWN : Message = #WM_LBUTTONDOWN ; Эмуляция левой кнопки.
Case #WM_RBUTTONUP   : Message = #WM_LBUTTONUP : MakeContextMenu(GetGadgetState(#ClipList))
Case #WM_MOUSEMOVE   : Define *Node.ClipData = Cursor2Node() ; Получаем нод под курсором.
If *Node > 0 : ListTip(*Node\Comment) : Else : ListTip() : EndIf ; Выставляем подсказку.
EndSelect ; Вызываем старый обработчик:
ProcedureReturn ChainOldCB()
EndProcedure

Macro HKNode(KI = KeyIdx) ; Partializer.
Node2Index(System\HotNodes[KI])
EndMacro

Macro UpHoldSelection(KI = KeyIdx, NI = State) ; Partializer.
If System\HotNodes[KI] : HLLine(HKNode()) : Else : HLLine(NI) : EndIf 
EndMacro

Macro SelectUsed() ; Partializer.
HLLine(USedNodeIdx())
EndMacro
;;;;;;;;---------
Procedure BindNode(NodeIdx, KeyIdx)
If UsedNode() : LinkHotkey(NodeIdx, KeyIdx)                       ; Линкуем клавишу.
If OverLoad(CountGadgetItems(#ClipList)) : CheckMaximum() : EndIf ; Проверяем перегрузку. Да, так надо.
EndIf
EndProcedure
;;;;;;;;---------
Procedure AnalyzeHotkey(KeyIdx)
Define AG = GetActiveGadget()
EnterCritical()                                            ; Блюдем скорость отрисовки.
Define State = GetGadgetState(#ClipList)
Select KeyIdx                                              ; Анализируем введенную клавишу...
Case #BringKey : UnhideList() : DisableDebugger : ReturnWindow() : EnableDebugger ; Вызываем окно на экран.
Case #SwapKey  : SwapLayout(UsedNodeIdx())                 ; Вроде как, меняем раскладку.
Case #SaveKey  : SelectUsed()                              ; Выделяем-таки.
UnhideList() : InvalidateRect_(System\ListID, #Null, #True) : UpdateWindow_(System\ListID) ; Dirty trick, but yeah.
SaveAs(UsedNodeIdx())                                      ; Сохраняем выделение, например.
Case #Hype+1 To #Hype + #Hotkeys : KeyIdx - #Hype          ; Клавиши горячие, вездесущие. Ну, бывает.
LinkHotkey(UsedNodeIdx(), KeyIdx, #True) : SelectUsed()    ; Линкуем и ставим выделение.
Case 1 To #HotKeys                                         ; Клавиши горячие. То, ради чего все и делалось.
If AG <> #SearchBar                                        ; Ничего не делаем, пока активна строка поиска.
If GetActiveWindow() = #MainWindow : BindNode(GetGadgetState(#ClipList), KeyIdx) ; Перевязываем нод к клавише.
ElseIf System\HotNodes[KeyIdx]                             ; Если там еще есть, что выделять...
RestoreData(HKNode()) : EmulatePasting() : EndIf           ; Восстанавливаем данные.
UpHoldSelection() : EndIf                                  ; Блюдем выделение.
EndSelect : LeaveCritical()                                ; Продолжаем блюсти.
EndProcedure

Macro RegHypeKey(Button) ; Pseudo-procedure.
RegisterHotKey_(0, Button, #CtrlShift, Button)
EndMacro

Macro RegFormat(Priority, Desc, Prefix, ExtraDesc = 0) ; Partializer.
System\Clipformats[Priority] = Desc : EnableGadgetDrop(#ClipList, ExtraDesc, #DragOff) : Pref(Desc) = Prefix + #Postf
EndMacro
;}
;{ --Searchworks--
Procedure GetSearchAnchor()
Define *Anchor = Index2Node(GetGadgetState(#ClipList))
If *Anchor : ProcedureReturn *Anchor : Else : LastElement(System\ClipList()) : ProcedureReturn System\ClipList() : EndIf
EndProcedure

Procedure FinishSearch(SetList.i = #True)
DisableGadget(#SearchBar, #False) : DisableGadget(#LurkBack, #False) : DisableGadget(#LurkForth, #False)
If SetList : If System\Lurker\ToGadget = 0 : SAG() : Else : SAG(System\Lurker\ToGadget) : EndIf : EndIf
System\Lurker\Target = #sNone ; Убиваем флаг поиска. По большому счету - вместе с поиском.
System\Lurker\ToGadget = 0 : System\Lurker\TextMark = "" ; Экономия памяти.
EndProcedure

Procedure CountChar(*Ptr.Unicode, Limit.i, Char.U)
Define *Ender = *Ptr + (Limit << 1), Count
While *Ptr\U : If *Ptr\U = Char : Count + 1 : EndIf ; Считаем указанный символ.
If *Ptr = *Ender : Break : EndIf : *Ptr + SizeOf(Unicode) : Wend
ProcedureReturn Count                               ; Возвращаем рассчеты.
EndProcedure

Procedure SelectVPText(*VP.ViewPort, *Rgsr, Pos)
Define *Node.ClipData = *VP\DataNode ; Ускоряем доступ.
Select *Node\DataType ; Выбираем по типу прилинкованных данных.
Case #CF_HTML ; Для HTML - особые условия.
Define *SelObject.COMateObject = *VP\WebObject\GetObjectProperty("Document\body\CreateTextRange()")
*SelObject\Invoke("move('character', " + Pos + ")")
*SelObject\Invoke("moveEnd('character', " + Len(System\Lurker\TextMark) + " )")
*SelObject\Invoke("select()") : *SelObject\Release() ; Высвобождаем.
Default ; Для стандартной репрезентации.
Define Sel.CHARRANGE, Size = Len(System\Lurker\TextMark)
Define StartOffset = CountChar(*Rgsr, Pos, #CR), EndOffset = CountChar(*Rgsr + (Pos) << 1, size, #CR)
Sel\CpMin = Pos - StartOffset : Sel\CpMax = Sel\CpMin + Size - EndOffset
SetSel(*VP\ViewArea, Sel) ; Собственно, оно самое.
EndSelect
EndProcedure

Procedure HLTextMark(*Node.ClipData, *Rgsr, Pos)
Pos - 1 ; Сразу коррeктируем, благо отсчет везде от нуля.
If *Node\ViewPort           : SelectVPText(*Node\ViewPort, *Rgsr, Pos)    : EndIf ; Если порт открыт...
If System\Panopticum\Actual : SelectVPText(System\Panopticum, *Rgsr, Pos) : EndIf
EndProcedure

Macro EscapeSearch() ; Pseudo-procedure.
(GetAsyncKeyState_(#VK_ESCAPE) And GetAsyncKeyState_(#VK_SHIFT))
EndMacro

Procedure SearchPossible()
ProcedureReturn Bool(ListSize(System\ClipList()) And SearchStall() And System\ActualReq)
EndProcedure

Procedure SetHold(*DN.ClipData, FieldID)
If Index2Node(GetGadgetState(#ClipList)) = *DN        ; Если этот нод уже проаналаизирован...
Define DataCount = CountGadgetItems(System\Informator\InBox) - 1 ; Получаем количество данных в спиcке
For I = 0 To DataCount : If GetGadgetItemData(System\Informator\InBox, I) = FieldID                          ; Ищем элемент.
SetDroppedState(System\Informator\InBox, #True) : SetGadgetState(System\Informator\InBox, I) : Break : EndIf ; Нашли ? Выделяем.
Next I : Else : System\HoldField = FieldID : EndIf    ; Иначе - выделение откладывается.
EndProcedure

Procedure SearchSuccess(*Engine.SearchEngine, *DN.ClipData, Txt.s, Rslt.i, HLField.i = 0) ; Pseudo-procedure.
LinkOptics(*DN) ; Сразу выставляем, дабы можно было подсветить в окне предпросмотра.
If *Engine\OpenFlag : OpenViewPort(Node2Index(*DN)) : *Engine\ToGadget = *DN\ViewPort\ViewArea : EndIf ; Просмотрщик.
If *Engine\HLFlag : If Rslt <> #SSuccess : HLTextMArk(*DN, @Txt, Rslt) ; Выделяем текст, если применимо.
ElseIf HLField : SetHold(*DN, HLField) : EndIf : EndIf      ; Выделяем поле, если применимо.
HLLine(Node2Index(*DN)) : FinishSearch() ; Ставим выделение списку и заканичваем.
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro EvaluateResult(ResAccum, HLField, Engine = System\Lurker) ; Partializer.
If Bool(ResAccum XOr Engine\DenialFlag) : If Engine\DenialFlag Or Engine\Target <> #sPlainText : ResAccum = #SSuccess : EndIf
SearchSuccess(Engine, *DN, Txt, ResAccum, HLField) : Break : EndIf                             ; Собственно, на случай успехов.
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro StringWhole(SPtr, Sptr2, CF) ; Suedo-procedure
-Bool(CompareMemoryString(@SPtr, @SPtr2, CF) = #PB_String_Equal)
EndMacro

Procedure FindInText(*Engine.SearchEngine, Src.s)
With *Engine ; Обрабатываем данные запроса.
Define SStart = 1 :  ; Временно так.
If \RegFlag And \WholeFlag : Dim Resultae.s(0) : ExtractRegularExpression(#SearchRXP, Src, Resultae()) ; Экстракция результатов.
If ArraySize(Resultae()) = 0 : ProcedureReturn StringWhole(Resultae(0), Src, \CaseFlag) : EndIf ; Проверяем на полную эквивалентность.
ElseIf \RegFlag   : ProcedureReturn -MatchRegularExpression(#SearchRXP, Src)   ; Ищем вхождение по регулярному выражению.
ElseIf \WholeFlag : ProcedureReturn StringWhole(Src, \TextMark, \CaseFlag)     ; Проверяем, равна ли строка данному запросу.
Else : ProcedureReturn FindString(Src, \TextMark, SStart, \CaseFlag) : EndIf   ; Иначе - ищем как обычно, текстовое вхождение.
EndWith
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure.s TargetToText(*DN.ClipData, Target)
With *DN
Select Target ; Строго по типу...
Case #sWSource : ProcedureReturn \WSource
Case #sRemark  : ProcedureReturn \Comment
Case #sDictID  : ProcedureReturn \DictID
EndSelect
EndWith
EndProcedure

Macro FindAndEval(HLField = 0) ; Partialzier.
Rslt.i = FindInText(System\Lurker, Txt) : EvaluateResult(Rslt, HLField)
EndMacro
; ----------------------------------
Macro FailSearch() ; Partializer.
HLLine(-1) : FinishSearch() : Break
EndMacro
; ----------------------------------
Procedure ApplyFiltration(*DN.ClipData, *Engine.SearchEngine)
With *Engine                    ; Отбираем преимущественно по флагам.
Define HK = *DN\Hotkey          ; Сохраняем для ускорения доступа и вообще.
If (\BindingFlag = #HotBound And HK = 0) Or (\BindingFlag = #HotUnBound And HK) : ProcedureReturn #False : EndIf
; -
Select *DN\DataType ; Типовая фильтрация. На случае, если тип попадает в один из заданных вариантов:
Case \DenySTR, \DenyBMP, \DenyDIR, \DenyRTF, \DenyHTML, \DenyMETA : ProcedureReturn \DenyNot ; По ситуации.
Default : ProcedureReturn \DenyNot ! #True ; А иначе - с точностью до наоборот.
EndSelect : EndWith ; Пока все привязано к DenyNot, но буду думать.
EndProcedure

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro DecorHeader(Hdr, Override = Offset) ; PSeudo-procedure.
Hex(Override) + ":[" + Hdr + "]"
EndMacro

Procedure SearchErr(Add.s = "")
ErrorBox("incorrect search header specified." + #CR$ + Add) : SAG(#SearchBar)
EndProcedure

Macro HeaderEnd(Hdr) ; Pseudo-procedure.
Val("$" + StringField(Hdr, 1, ":"))
EndMacro
; ------
Procedure CountRelativity(Array Relativity(1), Limiter, Start = 1)
Define *RelPtr.Integer = @Relativity(Start), *Ender = @Relativity(Limiter)
For *RelPtr = *RelPtr To *Ender Step SizeOf(Integer) : Define Offset = Offset + *RelPtr\I : Next *RelPtr 
ProcedureReturn Offset
EndProcedure
; ------
Procedure.s ActualReq(Hdr.s, FullReq.s, Array Offsetae(1))
Define EndOffset = HeaderEnd(Hdr) ; Получаем смещения до конца заголовка.
If Offsetae(0) : EndOffset + CountRelativity(Offsetae(), EndOffset) : EndIf
ProcedureReturn Mid(FullReq, EndOFfset + 1)
EndProcedure

Macro FlagForm(FlagName) ; Pseudo-procedure
#FlagSeparator + FlagName + #FlagSeparator
EndMacro

Macro RepChar(Out) ; Partializer.
CopyMemoryString(Out) : *RelPtr\I - Len(Out) : *RelPtr + SizeOf(Integer)
EndMacro

Macro RepReset() ; Partializer.
Replace = #False
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro CheckCodeSize() ; Partializer.
If CodeSize : CodeSize = 0 : RepChar(Chr(#SeqCode) + CodePref + Code) : Replace = #False : EndIf ; Оптимизировать !
EndMacro

Macro CheckRDelim() ; Partializer.
If Replace : CheckCodeSize() : RepReset() : RepChar(Chr(#SeqCode)) : EndIf
EndMacro

Macro CheckDelimAndCode() ; Partializer.
CheckRDelim() : CheckCodeSize()
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Macro ResAndOut(Char) ; Partializer.
RepReset() : RepChar(Char)
EndMacro

Macro OutThis() ; Partializer.
RepChar(*RPtr\Lense)
EndMacro

Macro TryReplacing(RepChar) ; Partializer.
CheckCodeSize() : If Replace : ResAndOut(RepChar) : Else : OutThis() : EndIf
EndMacro

Macro CodeSizeIF ; Partializer.
CheckCodeSize() : If Replace
EndMacro

Macro TryCodeSizing(CSize) ; Partializer.
CodeSizeIF : RepReset() : CodeSize = CSize : CodePref = *RPtr\Lense : Code = "" : Else : OutThis() : EndIf
EndMacro

Macro DepOut(Out) ; Partializer.
CopyMemoryString(Out) : *RelPtr\I = Len(Out) - 1 : *RelPtr + SizeOf(Integer)
EndMacro

Macro SeqOut(Out) ; Partializer.
DepOut(#SeqChar + Out)
EndMacro

Macro TMapEntry() ; Pseudo-procedure
If Entry = ""  And ErrCheck : ProcedureReturn "Empty flag encountered."     : EndIf ; Ошибка пустого флага.
Entry = LCase(Entry) ; Автоматически меняем регистр.
If TMap(Entry) And ErrCheck : ProcedureReturn "Duplicate flag encountered: " + FlagForm(Entry) : EndIf ; Ошибка дублирования.
TMap(Entry) = #True : Entry = "" ; Записываем вхождние.
EndMacro
;;;;;;;;;;;;;;;;;;;;;
Procedure.s GetSearchHdr(Request.s)
Define *PPtr.CharExtra = @Request, ParseIT, Offset = 1, Hdr.s
While *PPtr\U : Select *PPtr\U ; Анализируем сивольное наполнение.
Case ' ', #TAB ; NOP.
Case '[' : If ParseIT : ProcedureReturn #ErrHdr : Else : ParseIT = #True : EndIf                  ; Начинаем.
Case ']' : If ParseIT : ProcedureReturn DecorHeader(Hdr) : Else : ProcedureReturn #ErrHdr : EndIf ; Заканчиваем.
Default  : If ParseIT : Hdr + *PPtr\Lense : Else : Break : EndIf ; Обязательно проверяем - парсинг только после скобки.
EndSelect : *PPtr + SizeOf(Unicode) : Offset + 1 ; Инкрементируем счетчики.
Wend : If ParseIT : ProcedureReturn #ErrHdr : Else : ProcedureReturn DecorHeader("", 0) : EndIf 
EndProcedure

Procedure.s DisassembleHdr(Hdr.s, Map TMap.a(), ErrCheck = #True)
Define *PPtr.CharExtra = @Hdr, HdrStart, Entry.s
While *PPtr\U : Select *PPtr\U ; Анализируем сивольное наполнение.
Case '[' : HdrStart = #True    ; Рапортуем, что можно начать парсинг.
Case '/' : TMapEntry()         ; Добавляем слово в список.
Case ']' : If MapSize(TMap()) = 0 And Entry = "" : ProcedureReturn "" : Else : TMapEntry() : EndIf ; Финальный флаг.
Default  : If HdrStart : Entry + *PPtr\Lense : EndIf ; Добавляем сивол к текущему вхождению.
EndSelect : *PPtr + SizeOf(Unicode) : Wend
EndProcedure

Procedure.s ReparseRequest(Req.s, Array Relativity.i(1))
Define *RPtr.CharExtra = @Req, OutReq.s{#SearchLimit}, Replace.i, CodeSize.i, Code.s, CodePRef.s, *CStr = @OutReq
ReDim Relativity(#SearchLimit) : Define *RelPtr.Integer = @Relativity(1) ; Карта образования результата.
CopyMemoryString(@"", @*CStr)   ; Инциализация потокового копирования.
While *RPtr\U : *RelPtr\I + 1 : Select *RPtr\U  ; Анализируем символьное наполнение.
Case #SeqCode : CodeSizeIF : ResAndOut(#SeqChar) : Else : Replace = #True : EndIf ; Символом начала последовательности.
Case '#' : TryCodeSizing(2)     ; `#** для вывода символа по ASCII-коду.
Case '$' : TryCodeSizing(4)     ; `$**** для вывода символа по Юникоду. Вот прямо Юникоду, да.
Case #RepCR : TryReplacing(#CR$) ; `| для вывода #CR$
Case #RepLF : TryReplacing(#LF$) ; `~ для вывода #LF$
; -----------
Case '0' To '9', 'a' To 'f', 'A' To 'F' : If CodeSize = 1 And Val(Code) = 0 : CheckCodeSize() : Continue : EndIf ; Ну случай нуля.
CheckRDelim() : If CodeSize : Code + *RPtr\Lense : CodeSize - 1                      ; Если идет парсинг кода...
If CodeSize = 0 : ResAndOut(Chr(Val("$" + Code))) : EndIf : Else : OutThis() : EndIf ; ...Иначе - выводим уже этот несчастный символ.
; -----------
Default : CheckDelimAndCode() : OutThis() ; В самом общем случае - просто перегоняем символ в новый буффер.
EndSelect : *RPtr + SizeOf(Unicode) : Wend : CheckDelimAndCode() : ProcedureReturn OutReq  ; Возврат.
EndProcedure

Procedure.s DeparseString(Text.s, Array Relativity.i(1))
Define *DPtr.CharExtra = @Text, OutText.s{#DeparseStorage}, *CStr = @OutText ; Запасем сразу.
ReDim Relativity(#DeparseStorage) : Define *RelPtr.Integer = @Relativity(1)  ; Карта образования результата.
CopyMemoryString(@"", @*CStr)   ; Инциализация потокового копирования.
While *DPtr\U : Select *DPtr\U  ; Анализируем символьное наполнение.
Case #CR  : SeqOut(Chr(#RepCR)) ; Сменяем #CR$ на `|
Case #LF  : SeqOut(Chr(#RepLF)) ; Сменяем #LF$ на `~
Case #TAB : SeqOut("#09")       ; Сменяем #TAB на его код в HEX.
Default : DepOut(*DPtr\Lense)   ; Просто возвращаем символ. Он не очень интересен.
EndSelect : *DPtr + SizeOf(Unicode) : Wend : ProcedureReturn (OutText) ; Возвращем закодированный результат.
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro ReturnErr(Text) ; Partializer
System\Lurker\Target = #sNone : ProcedureReturn Text
EndMacro

Macro SetFlag(Flag, Val) ; Pseudo-procedure.
If Flag : ReturnErr("Dissonant flag encountered: " + FlagForm(TName)) ; Ошибка противоречащих флагов.
Else : Flag = Val + #FlagMask : EndIf ; Мирно выставляем значение флага.
EndMacro

Macro ProbeFlag(Flag, DefVal = #False) ; Pseudo-procedure.
If Flag = 0 : Flag = DefVal : Else : Flag - #FlagMask : EndIf ; Корректируем флаги.
EndMacro

Macro ArrangeHeader() ; Partializer.
Dim Offsets(0) ; Массив для дальнейшей обработки смещений (коли потребуется).
Define Req.s = ReparseRequest(System\ActualReq, Offsets()) ; Запрашиваем актуальные данные.
Define SHeader.s = GetSearchHdr(Req) ; Получаем заголовочную информацию запроса.
EndMacro

Macro ParseHeaderFast() ; Partializier.
NewMap Tokens.a() : ArrangeHeader() : If SHeader <> #ErrHdr : DisassembleHdr(SHeader, Tokens(), #False) : EndIf
EndMacro

Macro DefSearchFlag(FName, MenuItem, FlagPtr, AVal) ; Partializer.
AddMapElement(System\Emitter(), Hex(MenuItem))  : Define *SF.SearchFlag = System\Emitter()
*SF\Represent = FName : *SF\MenuLink = MenuItem : *SF\Data = @System\Lurker\FlagPtr : *SF\FlagVal = AVal
System\Reflector(FName) = *SF ; Дописываем ссылку.
EndMacro

Macro MIToFlag(MenuIndex) ; Pseudo-procedure.
System\Emitter(Hex(MenuIndex))
EndMacro

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure SearchPulse() ; To be moved.
With System\Lurker   ; Приступаем.
If Not SearchStall() ; Если на данный момент что-то ищется...
Define *Edge, *Start = ElapsedMilliseconds() ; Уточняем, когда начали.
ChangeCurrentElement(System\ClipList(), \Lense)
; =Loop goes here=
Repeat : If \Direction ; Выбираем по направлению. Это не сложно.
      : *Edge = PreviousElement(System\ClipList()) : If *Edge = #Null : *Edge = LastElement(System\ClipList())  : EndIf
Else  : *Edge = NextElement(System\ClipList())     : If *Edge = #Null : *Edge = FirstElement(System\ClipList()) : EndIf
EndIf : If *Edge = #Null Or EscapeSearch() :  FailSearch() : EndIf ; На случай экстренного выхода.
Define *DN.ClipData = System\ClipList() ; Получаем элемент к анализу.
; ------------------------
If ApplyFiltration(*DN, System\Lurker) ; Если нод проходит по заданным критериям поиска...
Select \Target   ; Выбираем по типу запрошенного поиска, да.
Case #sPlainText                   ; ----->Поиск по текстовой репрезентации нода.
Define Txt.s = ExtractText(*DN), RsLt : FindAndEval()
Case #sWSource, #sRemark, #sDictID ; ----->Поиск по вторичным текстовым данным.
Txt = TargetToText(*DN, \Target)      : FindAndEval(\Target)
Case #sListLine                    ; ----->Поиск по строкам в списке нодов.
Txt = GetListText(*DN)                : FindAndEval()
EndSelect : Txt = "" : EndIf : If *DN = \Anchor : FailSearch() : EndIf ; Выходим, коли вернулись на начало.
; ------------------------
Until ElapsedMilliseconds() - *Start >= 1000 ; Ищем не более секунды.
Else : FinishSearch(#False) : EndIf          ; Просто на всякий случай, с чего бы и нет ?
EndWith
EndProcedure
;;;;;;;;;;;;;;;;;;;;;
Procedure.s CrunchHeader(Header.s, FullRequest.s) ; To be moved.
Dim Dummy(0) ; Массив-заглушка для парсеров.
Define Req.s = ActualReq(Header, FullRequest, Dummy()) ; Получаем актуальный запрос.
NewMap Tokens.a() : Define HErr.s = DisassembleHdr(Header, Tokens())
If HErr : ProcedureReturn HErr : EndIf        ; Сразу выходиим с ошибкой первичного разбора.
Define TG = System\Lurker\ToGadget            ; Сохраняем флаг перехода. Он уникален.
ClearStructure(@System\Lurker, SearchEngine)  ; Очищаем, дабы избежать пересечений.
System\Lurker\ToGadget = TG                   ; Возвращаем обратно флаг перехода.
With System\Lurker ; Для пущего удобства.
; -Обрабатываем полученный список флагов.-
ForEach Tokens() : Define TName.s = MapKey(Tokens()), *Standart.SearchFlag = System\Reflector(TName)
If *Standart : SetFlag(*Standart\Data\I, *Standart\FlagVal) ; Если это стандартный случай...
Else : Select TName ; Обрабатываем особые флаги.
Case ""             ; Временно так, да.
Default      : ReturnErr("Unable to recognize flag: " + FlagForm(TName)) ; Флаг не известен.
EndSelect : EndIf : Next ; И теперь подбиваем результаты...
If \Target = #sNone : \Target = #sPlainText : Else : \Target - #FlagMask : EndIf ; Дабы пользователь не путался.
Select \Target      ; Последние подбивки по типам - параметры умолчания.
Case #sPlainText, #sWSource, #sRemark, #sDictID, #sListLine : \TextMark = Req  ; Сразу ставим цель.
ProbeFlag(\CaseFlag, #PB_String_NoCase) ; Ставим значения по умолчанию.
EndSelect           ; ...И теперь ставим общие флаги:
ProbeFlag(\DenySTR)  : ProbeFlag(\DenyBMP) : ProbeFlag(\DenyDIR) : ProbeFlag(\DenyRTF) : ProbeFlag(\DenyHTML) : ProbeFlag(\DenyMETA)
ProbeFlag(\OpenFlag) : ProbeFlag(\DenialFlag) : ProbeFlag(\HLFlag, #True) : ProbeFlag(\BindingFlag) ; Остатки категории Misc flagi.
ProbeFlag(\DenyNot)  ; Флаг инверсии проверяем одним из последних fgj.
Define Factor = Bool(\DenyStr) + Bool(\DenyBMP) + Bool(\DenyDIR) + Bool(\DenyRTF) + Bool(\DenyHTML) + Bool(\DenyMETA)
If Factor = (\DenyNot ! 1) * #AllFormats : ReturnErr("At least one data type should be out of denial.") : EndIf
; -Отдельно - подготовка для регулярных выражений. Уже в самом конце-
If \RegFlag  : If \TextMark ; Если есть, из чего делать выражение..
Define RFlag : If \CaseFlag = #PB_String_NoCase : RFlag | #PB_RegularExpression_NoCase : EndIf ; Нечуствительность к регистру.
If CreateRegularExpression(#SearchRXP, \TextMark, RFlag) = #Null                               ; Если не удалось создать выражение...
ReturnErr("PCRE->" + RegularExpressionError())  : EndIf                                        ; Рапорт PCRE о внутренней ошибке.
ElseIf \WholeFlag : ReturnErr("Regular expression may not be empty, try static matching.")     ; ...Даже если очень хочется - нельзя.
ElseIf Req : ReturnErr("Regular epressions may not match binary data.") ; Не хотеть бинарных поисков по /reg/.
EndIf : EndIf
EndWith
EndProcedure

Procedure InitSearch(Reverse = #False) ; To be moved.
If SearchPossible() ; Если есть, о чем вообще говорить...
With System\Lurker ; Выставляем начальные данные.
; -Первичная обработка запроса-
ArrangeHeader()    ; Получаем актуальный заголовок.
If SHeader = #ErrHdr : SearchErr("Syntax not recognized.") : ProcedureReturn #False : EndIf ; Ошибка, если заголовок не корректен.
Define ParseErr.s = CrunchHeader(SHeader, Req) : If ParseErr : SearchErr(DeparseString(ParseErr, Offsets())) ; Возвращаем ошибку.
ProcedureReturn #False : EndIf ; Выходим с порложенным ситуации нулем.
; --Вторичная обработка приголотовления--
If \TextMark Or \WholeFlag ; Если есть, что вообще искать...
\Direction = Reverse : \Anchor = GetSearchAnchor() : \Lense = \Anchor ; Коориданты направлений поиска.
SAG() ; Ставим выделение туда, где все одно скоро свершится поиск.
DisableGadget(#SearchBar, #True) : DisableGadget(#LurkBack, #True) : DisableGadget(#LurkForth, #True) ; Все выключаем.
SearchPulse() : If Not SearchStall() : ProcedureReturn #True : EndIf ; Рапортуем, что поиск начат и не закончен.
Else : \Target = #sNone : EndIf ; Отменяем поиск, если там назначен.
EndWith : EndIf ; Сразу отправляем на поиск, дабы избежать мерцания.
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; -GUIPathy-
Procedure SetSearchBar(Txt.s)
If Txt <> System\ActualReq ; Проверяем, новый ли хотят вписать текст...
Define Min, Max : Dim Offsets(0) : GetSimpleSel(#SearchBar, Min, Max)
If Len(Txt) > #SearchLimit : Txt = Left(Txt, #SearchLimit) : EndIf
Txt = DeParsestring(Txt, Offsets()) ; Переносы потом могут добавлять кодами, да.
Define MinOff = CountRelativity(Offsets(), Min)          ; Смещение на старт выделения.
Define MaxOff = CountRelativity(Offsets(), Max, Min + 1) ; Смещение на конец выделения.
Min + MinOff : Max + MinOff + MAxOff                     ; Корректируем отступы с учетом неравномерности.
SetGadgetText(#SearchBar, Txt) : SetSimpleSel(#SearchBar, Min, Max)
System\ActualReq = Txt ; Выставляем в аккумулятор реальное значние.
SetGadgetColor(#SearchBar, #PB_Gadget_FrontColor, #Black) ; Автоматически и всегда - там черный.
EndIf
EndProcedure

Procedure SearcherDrop() ; Bindback
SetSearchBar(EventDropText())
EndProcedure

Macro SetSearcherDummy() ; Pseudo-procedure.
If System\ActualReq = "" : SetGadgetText(#SearchBar, "...Search me requests...")
SetGadgetColor(#SearchBar, #PB_Gadget_FrontColor, $707070) : EndIf
EndMacro 

Procedure SearchDaemon() ; Bindback.
Select EventType() 
Case #PB_EventType_LostFocus : SetSearcherDummy()
Case #PB_EventType_Change    : SetSearchBar(GetGadgetText(#SearchBar))
Case #PB_EventType_Focus     : If System\ActualReq = "" : SetGadgetText(#SearchBar, "") : EndIf
EndSelect
EndProcedure

Macro SelectMI(Index, Text, SC, Flag = #False) ; Partializer.
MenuItemEx(Index, Text, Sc) : If Not (AnyText Or Flag) : DisableMenuItem(#mSearchMenu, Index, #True) : EndIf
EndMacro

Macro DisableSI(Index) ; Partializer.
DisableMenuItem(#mSearchMenu, Index, #True)
EndMacro

Macro EmitterItem(ItemIdx, Sign, BoundFlag = 0) ; Partializer.
Define *SF.SearchFlag = MIToFlag(ItemIdx) ; Ищем элемент в таблице.
MenuItemEx(ItemIdx, ReplaceString(ReplaceString(Sign, "*", FlagForm(*SF\Represent)), "~", "")) ; Создаем пункт меню.
If *SF : SetMenuItemState(#mSearchMenu, ItemIdx, Tokens(*SF\Represent)) : EndIf  ; Проверяем и ставим отметку.
If BoundFlag : *SF = MIToFlag(BoundFlag) : If *SF                                ; Провереяем также связанный флаг.
DisableMenuItem(#mSearchMenu, ItemIdx, Tokens(*SF\Represent))                    ; Он есть ? Ну и отключаем тогда.
EndIf : ElseIf Skipper And Skipper <> ItemIdx : DisableMenuItem(#mSearchMenu, ItemIdx, #True) : EndIf ; Обработка категорий.
EndMacro

Macro SetDisFlag(TField) ; Partializer.
ForEach Tokens() : Define *SF.SearchFlag = System\Reflector(MapKey(Tokens()))
If *SF And *SF\Data = @System\Lurker\TField : Define Skipper = *SF\MenuLink : Break : EndIf
Next
EndMacro

Macro AddHdrFlag(Rep) ; Partializer.
OutHeader + Rep + #FlagSeparator
EndMacro

Procedure.s ReconstructHeader(Offer)
ParseHeaderFast() : Offsets(0) = #True ; Настраиваем корректор. Так надо.
Define OutHeader.s = "[", TName.s, *SF.SearchFlag = MIToFlag(Offer), TGT.s = *SF\Represent ; Необходимые данные.
ForEach Tokens() : TName = MapKey(Tokens()) : If TName : If TName <> TGT : AddHdrFlag(TName) : Else : TGT = "" : EndIf : EndIf ; Вписываем.
Next : If TGT : AddHdrFlag(TGT) : EndIf : If Len(OutHeader) > 1 : OutHeader = Left(OutHeader, Len(OutHeader) - 1) + "]"    ; Доформатируем.
Else : OutHeader = "" : EndIf : ProcedureReturn OutHeader + ActualReq(SHeader, System\ActualReq, Offsets()) ; Наконец, возвращаем то самое.
EndProcedure

Macro SelSBar() ; Partializer.
SearchBarCallBack(System\SearchBar, #WM_SETFOCUS, 0, 0)
EndMacro

Procedure SearchBarCallBack(hWnd, Message, wParam, lParam) ; Callback.
Select Message ; Анализируем входящие сигналы...
Case #WM_LBUTTONDBLCLK : SetSimpleSel(#SearchBar, 0, -1) : ProcedureReturn #False ; Выделяем все. Удобно же.
Case #WM_CONTEXTMENU   : CreatePopupMenu(#mSearchMenu) : Define AnyText = Bool(System\ActualReq)
SelectMI(#sCut, "Cut", "Shift+Del") : SelectMI(#sCopy, "Copy", CtrlSC("Ins")) : MenuItemEx(#sPaste, "Paste", "Shift+Ins") 
SelectMI(#sClear, "Delete", "Del") : MenuBar() : SelectMI(#sSelectAll, "Select All", CtrlSC("A")) : MenuBar()
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
ParseHeaderFast() : OpenSubMenu("Emit search flag")   ; Общая категория эмиттера.
OpenSubMenu("Targeting options") : SetDisFlag(Target) ; Точки проведения поиска.
EmitterItem(#sfText, "Plain *")   : EmitterItem(#sfSrc, "Caption of *")  ; Стандартный текстовый поиск.
EmitterItem(#sfRem,  "Bound *ark") : EmitterItem(#sfID, "Dictionary *")  ; Частный случай текстового поиска.
EmitterItem(#sfList, "Node* line")                                       ; Избыточный текстовый поиск.
CloseSubMenu() : Skipper = #False : OpenSubMenu("Typing exclusion")      ; Фильтрация результатов по типу данных.
EmitterItem(#sfNoSTR, #Denial) : EmitterItem(#sfNoRTF , #Denial) : EmitterItem(#sfNoDIR , #Denial) ; Первая партия фильтров.
EmitterItem(#sfNoBMP, #Denial) : EmitterItem(#sfNoHTML, #Denial) : EmitterItem(#sfNoMETA, #Denial) ; Вторая партия фильтров.
EmitterItem(#sfUnDeny, "Denial *ersion") ; Дополнительный флаг для инверсивной фильтрации типов.
CloseSubMenu() : OpenSubMenu("Text analysis")                            ; Все необходимое для поиска по тексту.
EmitterItem(#sfCase, "Sensitive to *", #sfNoCase) : EmitterItem(#sfNoCase, "Sensitive to no *", #sfCase) ; Текстовые флагов.
EmitterItem(#sfWhole , "Equals to * source")      : EmitterItem(#sfRegular, "Match *ular expression")    ; Еще немного флагов.
CloseSubMenu() : OpenSubMenu("Misc flagi") : EmitterItem(#sfOpen, "* findings") : EmitterItem(#sfNot, "* equality") ; Побочные флаги.
EmitterItem(#sfSel, "Auto-*ect results", #sfNoSel) : EmitterItem(#sfNoSel, "Supress *ection", #sfSel)    ; Дополнительные флаги.
EmitterItem(#sfHKBound, "Hotkey-* only", #sfHKUnBound) : EmitterItem(#sfHKUnBound, "Hotkey-un* only", #sfHKBound)   ; По нодам.
CloseSubMenu() : CloseSubMenu() : OpenSubMenu("Emit reparser code") ; Отдельная категория для эмиттера заменяемых символов.
MenuItemEX(#dTAB, "Alt+009 (TAB)") : MenuItemEX(#dLF, "Alt+010 (LF)") : MenuItemEX(#dCR, "Alt+013 (CR)") ; Основной набор.
CloseSubMenu() : MenuBar() ; Финальный разделитель перед остаточной частью меню.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
MenuItemEx(#cFindNext, "Find Next", "F3") : MenuItemEx(#cFindPrev, "Find Previous", "Shift+F3") : Define Min, Max
If Not SearchPossible() : DisableSI(#cFindNext) : DisableSI(#cFindPrev) : EndIf  ; Отключаем, если поиск невозможен.
Define Min, Max : GetSimpleSel(#SearchBar, Min, MAx) : If Min = MAx              ; Если по факту нет выделения...
DisableSI(#sCut) : DisableSI(#sCopy) : DisableSI(#sClear) : EndIf                ; Отключаем работу с ним.
If GetClipboardText() = "" : DisableSI(#sPaste) : EndIf                          ; И вот такое тоже надо бы в стандарт.
DisplayPopupMenu(#mSearchMenu, System\MainWindow) : ProcedureReturn 0            ; Показываем свое меню вместо системного.
EndSelect : ProcedureReturn ChainOldCB(hWnd)
EndProcedure

Procedure EmitSymbol(*SBar, MenuIndex)
; -----------
Select MenuIndex
Case #dTAB : Define OutSeq.s = #TAB$ ; Выдаем табулятор.
Case #dLF  : OutSeq = #LF$   ; Выдаем возврат строки.
Case #dCR  : OutSeq = #CR$   ; Выдаем перенос строки.
EndSelect  : Define Min, Max ; Корректоры.
Dim Offset(0) : GetSimpleSel(*SBar, Min, MAx)
Define RepText.s = ReparseRequest(System\ActualReq, Offset())
Define *RelPtr.Integer = @OFfset(1), CheckSize
While CheckSize < Min : CheckSize + *RelPtr\I + 1 : *RelPtr + SizeOf(Integer) : Wend ; Подсчет корректного смещения.
If CheckSize > Min : Define Factor = CheckSize - Min : Min + Factor : Max + Factor   ; Перерассчет смещения.
SetSimpleSel(*SBar, Min, Max) : EndIf ; Корректируем смещения для корректной вставки. Хоть тут - все.
SelSBar() : SendMessage_(GadgetID(*SBar), #EM_REPLACESEL, 1, @OutSeq) ; Вызов итогов.
; -----------
EndProcedure
;}
;{ --Drag/Drop--
; ---------------------------------------
Procedure ReserveBuffer(*OBuf, OSize)
Define *Result = AllocateMemory(OSize) : CopyMemory(*OBuf, *Result, OSize)
ProcedureReturn *Result
EndProcedure

Macro PrepareDragArraySize(FCount) ; Pseudo-procedure.
ReDim System\DragOut(FCount)
EndMacro

Macro PrepareOSDrag(Fmt, Ptr, DSize) ; Pseudo-procedure.
Define DCount, *OSPtr.DragDataFormat = System\DragOut(DCount) : DCount + 1 ; Инкрементируем счетчик.
*OSPtr\Format = Fmt : *OSPtr\Size = DSize : *OSPtr\Buffer = Ptr ; Пока так.
EndMacro

Macro EngulfDDBuffer()
Define *DataPtr = ReserveBuffer(EventDropBuffer(), EventDropSize())
EndMacro
; ---------------------------------------
Macro DragOutCT(Fmt = System\ClipRTF) ; Partializer.
RestoreComplexTextGuts(*DataNode, Fmt)
PrepareDragArraySize(1) 
PrepareOSDrag(#CF_UNICODETEXT, @Plain, StringByteLength(Plain) + #CharSize)
PrepareOSDrag(Fmt, *DPtr, DSize)
DragOSFormats(System\DragOut(), 2, #DragOff)
CleanAfter(*DataNode,*DPtr) 
EndMacro

Procedure DragImg(Index) ; Replacer.
DragImage(ImageID(Index), #DragOff)
EndProcedure

Procedure DragEMF(*MetaPtr, Void, Void2, Void3) ; Replacer.
PrepareDragArraySize(0)
PrepareOSDrag(#CF_TEXT, *MetaPtr, 0)
DragOSFormats(System\DragOut(), 1, #DragOff)
EndProcedure

Macro DragOutText() ; Partializer.
DragText(ExtractText(*DataNode), #DragOff)
EndMacro

Macro DragOutBMP() ; Partializer.
RestoreImageGuts(*DataNode, DragImg)
EndMacro

Macro DragOutMeta() ; Partializer.
RestoreMetaGuts(*DataNode, DragEMF)
EndMacro

Procedure DropCB(TargetHandle, State, Format, Action, x, y) ; Callback
System\DropFlag = System\DragFlag ; Вписываем флаг, раз уж.
ProcedureReturn #True
EndProcedure

Procedure DragNode(NodeIdx)
If UsedNode() ; Если есть, о чем вообще говорить...
Define DataNode(), Result : System\LockedNode = *DataNode : System\DragFlag = #True
With *DataNode ; Дополнительная обработка для перетаскивания информации во внешние приложения..
Select \DataType
Case #CF_TEXT        : DragOutText()              ; Выбрасываем текст.
Case #CF_BITMAP      : DragOutBMP()               ; Выбрасываем BMP.
Case #CF_RichText    : DragOutCT()                ; Выбрасываем RTF.
Case #CF_HTML        : DragOutCT(System\ClipHTML) ; Выбрасываем HTML.
;Case #CF_ENHMETAFILE : DragOutMeta()             ; Выбрасываем EMF. Пока пусть побудет отключенным.
; Default'ный вариант до полной реализации.
Default : DragPrivate(0, #PB_Drag_Link) 
EndSelect : EndWith : EndIf
; Очищаем массив форматов.
System\DragFlag = #False : PrepareDragArraySize(0) ; Очищаем массив форматов.
EndProcedure
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Macro DropFiles() ; Partializer.
Define FList.s = EventDropFiles() : If FList ; Если в списке нам что-то соизволили дать...
Define Counter = CountLDelims(@FList) + 1 : RegisterFilesGuts(#False) : EndIf
EndMacro
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Procedure ListDrop()
With System
Define DPoint.Point : DPoint\X = EventDropX() : DPoint\Y = EventDropY()  ; Записываем коордианты вброса. Сразу, вот да.
System\LastSrc = GadgetID(#VoidGadget) ; Временное решение, пока я не найду надежный дескриптор.
If \DropFlag : \DropFlag = #False : TakeFocus() : If \LockedNode <> #Null : Define *DataNode.ClipData = \LockedNode
Else : ErrorBox("unable to drag data from deleted node !" + #ReconMsg) : ResumeFocus() : ProcedureReturn : EndIf ; Выходим с ошибкой.
Else : Select EventDropType() ; Иначе обрабатываем входящие данные.
; -------------------------------------
Case #PB_Drop_Text   : Define Text.s = EventDropText()                    : RegisterTextGuts(#False)
Case #PB_Drop_Image  : Define Image  = EventDropImage(#PB_Any, #BitDepth) : RegisterImageGuts(#False)
Case System\ClipRTF  : EngulfDDBuffer() : RegisterCTGuts(#CF_RichText, Pref(#CF_RichText), #False)
Case System\ClipHTML : EngulfDDBuffer() : RegisterCTGuts(#CF_HTML    , Pref(#CF_HTML)    , #False)
Case #PB_Drop_Files  : DropFiles()
; -------------------------------------
EndSelect : If *Host : *DataNode = *Host : EndIf ; Ну, посмотрим, что из оного выйдет...
EndIf : EndWith ; А теперь - перемещаем результат на положенное ему место:
Define *Target.ClipData = Cursor2Node(DPoint), SrcIdx = Node2Index(*DataNode), Mode
If *Target : If Node2Index(*Target) < SrcIdx : Mode = #PB_List_Before : Else : Mode = #PB_List_After : EndIf  ; Выбираем режим переноса.
ThrowNode(SrcIdx, Mode, *Target) : Else : ThrowNode(SrcIdx) : HLLine(CountGadgetItems(#ClipList) - 1) : EndIf ; ...Или кидаем в конец, да.
EndProcedure

Procedure GadgetDrop() ; Bindback.
Select EventGadget()
Case #ClipList  : ListDrop()
Case #SearchBar : SearcherDrop() ; Тут все чуть проще...
EndSelect
EndProcedure
;}
;{ --Snapping management--
Macro SetSnapped(Win, Dir, Val = #True) ; Partializer.
SetProp_(WindowID(Win), Dir, Val)
EndMacro

Macro ResetSnapped(Dir) ; Partializer.
SetSnapped(*Window, Dir, #False)
EndMacro

Macro CheckSnapped(Win, Dir) ; Partializer.
GetProp_(WindowID(Win), Dir)
EndMacro

Procedure DownCheck(*Window, Flag.s)
Define FVal = CheckSnapped(*Window, Flag)
If FVal = 0 : ProcedureReturn #True : EndIf
SetSnapped(*Window, Flag, FVal - 1)
EndProcedure

Procedure FillWinRect(*Window, *Bounds.Rect)
With *Bounds
\Top = WindowY(*Window) : \Left = WindowX(*Window) - System\SnapShift
\Bottom = \Top  + WindowHeight(*Window, #PB_Window_FrameCoordinate)
\Right  = \Left + WindowWidth(*Window, #PB_Window_FrameCoordinate)
EndWith
EndProcedure

Macro InBound(Val, Min, Max) ; Pseudo-procedure.
(Val => Min And Val <= Max)
EndMacro

Macro CheckSpacing(Coord, CoordAlt, LBound, UBound) ; Partializer.
If InBound(Coord, LBound, UBound) Or InBound(Coordalt, LBound, UBound) Or
InBound(LBound, Coord, CoordAlt) Or InBound(UBound, Coord, CoordAlt)
EndMacro

Macro TrickSnap(Target, Dest, Dir, Action) ; Partializer.
If CheckSnapped(*SrcWin, Dir) = #False Or *TargetWin = GetActiveWindow(); ^Если мы еще не клеились по этому направлению^.
Define Spacing = Target - Dest : If Abs(Spacing) <= System\Gravitation ; Если все корректно...
Action : SetSnapped(*SrcWin, Dir) : Break : EndIf ; ...Сдвигаем окно.
EndIf
EndMacro

Procedure CheckSnapDelay(*Win) ; Partializer
Define MS = CheckSnapped(*Win, #NoSnap)
If MS = 0 Or ElapsedMilliseconds() - MS > 200 : ProcedureReturn #True : EndIf
EndProcedure

Procedure TrySnappingTo(*SrcWin, *TargetWin)
If *SrcWin <> *TargetWin And ChecksnapDelay(*SrcWin) And IsWindowVisible_(WindowID(*TargetWin)) ; Если имеет смысл этим заниматься...
Define Src.Rect, Target.Rect : FillWinRect(*SrcWin, @Src)       ; Прямоугольник-источник.
Define After.Rect = Src      : FillWinRect(*TargetWin, @Target) ; Прямоугольник окна, к которому прятягиваемся.
With After ; -Проверяем, есть ли точки соприкосновения...-
Repeat ; Пытаемся притянуться к целевому окну...
CheckSpacing(\Top , \Bottom, Target\Top , Target\Bottom)         ; ---Горизонталь.
; ------------------------
TrickSnap(\Left, Target\Right, #HorSFlag, \Left = Target\Right)  ; Левый край.
TrickSnap(\Right, Target\Left, #HorSFlag, \Left - Spacing)       ; Правый край.
EndIf : CheckSpacing(\Left, \Right , Target\Left, Target\Right)  ; ---Вертикаль.
TrickSnap(\Top, Target\Bottom, #VertSFlag, \Top = Target\Bottom) ; Верхний край.
TrickSnap(\Bottom, Target\Top, #VertSFlag, \Top - Spacing)       ; Нижний край.
; ------------------------
EndIf : Until #True : If \Top <> Src\Top Or \Left <> Src\Left  ; Ну и уж пытаемся...
SetWindowPos_(WindowID(*SrcWin),#Null,\Left,\Top,0,0,#SWP_NOACTIVATE|#SWP_NOZORDER|#SWP_NOSIZE)
EndWith : ProcedureReturn #True 
EndIf : EndIf
EndProcedure

Macro WinSnapping(Dest) ; Partializer.
If *Window <> -1 : TrySnappingTo(*Window, Dest) : Else : SnapWorld(Dest) : EndIf
EndMacro

Procedure SnapWorld(*Window = -1)
If *Window <> -1 : If IsWindowVisible_(WindowID(*Window)) = #False : ProcedureReturn : EndIf      ; Выходим, если окно не видимо.
ResetSnapped(#VertSFlag) : ResetSnapped(#HorSFlag) : EndIf : PushListPosition(System\Viewports()) ; Сохраняем позицию...
ForEach System\Viewports() : WinSnapping(System\Viewports()\WindowID) : Next ; Пытаемся примагнитить.
PopListPosition(System\Viewports())  ; Загружаем позицию...
WinSnapping(System\Panopticum\WindowID)
WinSnapping(#MainWindow)             ; И вот теперь - магнетизм для остатков.
EndProcedure

Procedure SnapVerse()
If System\Options\SnapWin : System\Gravitation = #SnapSpace : SnapWorld() : EndIf
EndProcedure
;}
;{ --Hi-level callbacks--
Procedure ReceiveEvent(*Container.EventData)
With *Container
\Type = WaitWindowEvent()
Select \Type ; Проверка для событий, где этот шедевр инженерной мысли сходит с ума.
Case #WM_RBUTTONUP : \Window = GetActiveWindow_() : ProcedureReturn ; Так мы делаем, дабы всегда вовремя приходило меню.
Case #WM_HOTKEY : \Window = #MainWindow           : ProcedureReturn ; Так - дабы не было проблем с горячими клаившами.
Case #WM_CHAR : If GetActiveWindow_() <> System\MainWindow : \Type = -0 : Else : \Window = #MainWindow : EndIf
ProcedureReturn                                                     ; Так - дабы закончить безумие горячих клавиш.
Default : \Window = EventWindow() : EndSelect                       ; Ну и, якобы, вот так еще.
If IsWindow(\Window) ; Проверяем, не в 5.20+ ли мы случаем...
\SubType = EventType()
If \Type = #PB_Event_Gadget : \Gadget = EventGadget() : Else : \Gadget = #Null : EndIf
Else : \Type = -0 ; Viva la 5.20+ !
EndIf
EndWith
EndProcedure

Procedure GUISizer()
Splitteraze(System\FlickBuf)
Define SBOffset = FullWidth - #SearchROFfset + (#SearchLOFfset - 1)
ResizeGadget(#SearchBar, #PB_Ignore, #PB_Ignore, FullWidth - #SearchROFfset, #PB_Ignore)
ResizeGadget(#LurkBack , SBOffset, #PB_Ignore, #PB_Ignore, #PB_Ignore)
ResizeGadget(#LurkForth, SBOffset + #SBWidth - 1, #PB_Ignore, #PB_Ignore, #PB_Ignore)
ResizeGadget(#ClipList, #PB_Ignore, #PB_Ignore, FullWidth - #ClipROffset, FullHeight - #ClipBOffset)
Define Y         = FullHeight - #ButtonHeight - #ButtonSpace
Define BFactor.f = (FullWidth - #ButtonOffset * (#ButtonBar + 2)) / #ButtonBar
Define Offset.f  = #ButtonOffset
For I = #Button_Clear To #Button_Terminate ; Пересчитываем размер панели.
Define Width.f   = BFactor * (1.0 - (#ButtonBar / 2 - Abs(I - #Button_Ocular)) * 0.25)
ResizeGadget(I, Offset, Y, Width, #PB_Ignore) : Offset + Width + #ButtonOffset
Next I : ResizeInformer(System\Informator, #UniVPOffset, #ClipInformerB)
ResizeGadget(#EyeFrame, GadgetX(#Button_Ocular) - #EFOffset, GadgetY(#Button_Ocular) - #EFOffset, 
GadgetWidth(#Button_Ocular) + #EFSize,#PB_Ignore)
EndProcedure

Procedure WinCallback(hWnd, uMsg, wParam, lParam) 
Define NewPos.Point ; To prevent flickering.
Select uMsg ; Анализируем сообщение.
Case System\SizeMsg : GUISizer() ; Теперь такой плавный....
Case #WM_NCACTIVATE : If wParam = #False And System\Options\FixedTray : GetCursorPos_(NewPos)
If NewPos\X = System\CursorPos\X And NewPos\Y = System\CursorPos\Y ; Сравнение координат.
ProcedureReturn #False ; Предотвращаем мерцание.
EndIf : EndIf ; Получаем координаты на предотвращение:
Case 12501, 0 : GetCursorPos_(System\CursorPos)
Case #WM_ACTIVATE ; Сообщение об изменении фокуса.
If wParam = #WA_INACTIVE : MakeTransparent() : Else : SetOpacity() : EndIf
Case #WM_DRAWCLIPBOARD : StoreData()
Case #WM_CHANGECBCHAIN : System\NextWindow = lParam   ; Получаем следущее звено.
SendNotifyMessage_(System\NextWindow, uMsg, wParam, lParam) ; Высылаем ему сообщение.
ProcedureReturn 0                                     ; Возвращаем, что просили.
Case #WM_QUERYENDSESSION, #WM_ENDSESSION, #WM_DESTROY : ! JMP AfterMath
EndSelect
ProcedureReturn #PB_ProcessPureBasicEvents 
EndProcedure

Procedure SendSBKeys(VK)
SendMessage_(System\SearchBar, #WM_KEYDOWN, VK, #KDownParam)
SendMessage_(System\SearchBar, #WM_KEYUP  , VK, #KUpParam)
EndProcedure

Macro EmuCase(MenuEvent, Key, VK) ; Partializer.
Case MenuEvent         : RemoveKeyboardShortcut(#MainWindow, #PB_Shortcut_#Key) 
SendSBKeys(#VK_DELETE) : AddKeyboardShortcut(#MainWindow, #PB_Shortcut_#Key, MenuEvent) 
EndMacro

Procedure ProcessSearchBar(Bar, MenuItemEx)
If MenuItemEx > #dDerapserEdge : EmitSymbol(Bar, MenuItemEx) : ProcedureReturn #True : EndIf ; Выдаем спец. символ.
If MenuItemEx > #SfOffsetEdge : SetSearchBar(ReconstructHeader(MenuItemEx)) ; Реконструируем заголовк обратно.
Define TSize = Len(System\ActualReq) : SetSimpleSel(Bar, TSize, TSize) : ProcedureReturn #True : EndIf
Select MenuItemEx ; Клавиатурная эмуляция для SearchBar'а.
EmuCase(#cDelete, Delete, #VK_DELETE)                           ; Стандартный функционал Delete.
Case #cReturn : System\Lurker\ToGadget = #SearchBar             ; Уточняем, что надо вернуться к строке.
If InitSearch() = #False : System\Lurker\ToGadget = 0 : EndIf   ; Убираем в 0, дабы не смешивалось.
Case #cCtrlA, #sSelectAll  : SetSimpleSel(#SearchBar, 0, -1)                 ; Выделяем все содержимое. Целиком.
Case #cCtrlV, #sPaste      : SelSBar() : SendMessage_(System\SearchBar, #WM_PASTE, 0, 0) ; Выделяем все содержимое. Целиком.
Case #cCopy , #sCopy       : SendMessage_(System\SearchBar, #WM_COPY , 0, 0) ; Стандарт для клавиш восстановления.
Case #sClear               : SendMessage_(System\SearchBar, #WM_CLEAR, 0, 0) ; Очищаем выделение.
Case #sCut                 : SendMessage_(System\SearchBar, #WM_CUT  , 0, 0) ; Вырезаем выделенное.
Default   : ProcedureReturn #False ; Если ничего не узнали - надо продолжать обработку.
EndSelect : ProcedureReturn #True  ; Узнали, вот да.
EndProcedure
;-----------------------
Procedure SetWordWrap(*VP.ViewPort, WrapFlag.i)
If isTextNode(*VP\DataNode) ; Проверяем, дабы было не повадно.
WrapFlag = Bool(WrapFlag) : *VP\DataNode\WrapFlag = WrapFlag : ActualizeWordWrap(*VP)  ; Выставляем значение флага.
EndIf
EndProcedure
;-----------------------
Procedure UseMenu() ; Bindback.
Define AG = GetActiveGadget(), Item = EventMenu() ; Заранее.
If (AG = #SearchBar Or Item > #sSearchMenuEdge) And ProcessSearchBar(#SearchBar, Item) ; Совсем отдельный случай...
SAG(#SearchBar) : ProcedureReturn : EndIf                                              ; ...Для поисковика.
EnterCritical() ; Блюдем скорость отрисовки.
Define NodeIdx = GetGadgetState(#ClipList), X, Y, sel.CHARRANGE
; Актуализируем пункты меню:
Select Item; Анализируем пункты...
Case #tShowWindow : ReturnWindow() ; Показать окно программы.
Case #tOptions    ; Показать окно настроек.
If IsWindowVisible_(System\SetupWindow) : SetForegroundWindow_(System\SetupWindow) : Else : BringOptions() : EndIf
Case #tTerminate  : ! JMP AfterMath ; Выход из программы.
Case #tClearList  : WipeData()
Case #tSwitch     : System\AcceptNew ! 1 : UpdateSwitch()
Case #tUseNode To #tUseNode + #HotKeys ; Слоты быстрого доступа.
RestoreData(Node2Index(System\HotNodes[Val(Mid(GetMenuItemText(#mTrayMenu, Item), Len(#SlotPrefix) + 1))]))
; Теперь контекстное меню главного окна...
Case #cMoveUp     : SwapNodes(NodeIdx, NodeIdx - 1)
Case #cMoveDown   : SwapNodes(NodeIdx, NodeIdx + 1)
Case #cSaveAs     : SaveAs(NodeIdx)
Case #cFlatten    : ComplexText2STR(NodeIdx)
Case #cRender     : MetaFile2BMP(NodeIdx)
Case #cOptions    : BringOptions()              ; Virtual item.
Case #cPanopticum : BringPanopticum()           ; Virtual item.
Case #cSetComment : AskNodeComment(NodeIdx)
Case #cThrowDown  : ThrowNode(NodeIdx)          ; Кидаем вниз, вот да.
Case #cThrowUp    : ThrowNode(NodeIdx, #PB_List_First)  ; Подкидываем.
Case #cQWERTYSwap : SwapLayout(NodeIdx)         ; Меняем раскладку.
Case #cListerate  : ProduceListing(NodeIdx)     ; Превращаем текст в листинг
Case #cBindData To #cBindData + #HotKeys : Define KeyIdx.i = Item - #cBindData 
BindNode(NodeIdx, KeyIdx) : UpHoldSelection(KeyIdx, NodeIdx) ; Привязываем данные к клавише.
; ...и меню контекста Viewport'а:
Case #vCenterWin  : X = (GetSystemMetrics_(#SM_CXSCREEN)-WindowWidth(System\ViewPort\WindowID)) >> 1
Y = (GetSystemMetrics_(#SM_CYSCREEN) - WindowHeight(System\ViewPort\WindowID)) >> 1
ResizeWindow(System\ViewPort\WindowID, X, Y, #PB_Ignore, #PB_Ignore) ; Центрируем.
Case #vMaximize   : SetWindowState(System\ViewPort\WindowID, #PB_Window_Maximize)
Case #vReturnSize : SetWindowState(System\ViewPort\WindowID, #PB_Window_Normal)
Case #vSizeUp     : ScaleViewPort(System\ViewPort, 1.2)
Case #vSizeDown   : ScaleViewPort(System\ViewPort, 0.8)
Case #vSaveAs     : SaveAs(Node2Index(System\ViewPort\DataNode))
Case #vSaveSnap   : SaveScaled(System\ViewPort, Node2Index(System\ViewPort\DataNode))
Case #vHighLight  : SetGadgetState(#ClipList,Node2Index(System\ViewPort\DataNode)):SetForegroundWindow_(System\MainWindow)
System\CriticalStack() = #ClipList ; Выделяемся в списке.
; Обработка контекстного меню для строки поиска:
Case #cFindNext   : InitSearch()
Case #cFindPrev   : InitSearch(#True)
Case #cGoSearch   : System\CriticalStack() = #SearchBar
Case #cViewData, #cCtrlV  : OpenViewPort(NodeIdx) 
Case #cRemove  , #cDelete : ClearData(NodeIdx)
Case #cCopy    , #cReturn : RestoreData(NodeIdx)
; Спец. обработка для нодов текстового содержания:
Case #vCopy       : If NotHTMLView() : SendMessage_(GadgetID(System\ViewPort\ViewArea), #WM_COPY, 0, 0) : Else ; Прямое копирование.
HTMLasHTML(System\ViewPort) : EndIf
Case #vCopyRaw    : If NotHTMLView() : RTFAsRaw(System\ViewPort) : Else : HTMLAsRaw(System\ViewPort) : EndIf   ; Копирование текста.
Case #vCopyMD     : HTMLAsHTML(System\ViewPort, #True)                                                         ; Копирование разметки.
Case #vSelectAll  : If IsTextNode(System\ViewPort\DataNode, #True) : If NotHTMLView() : Sel\CpMax = -1         ; Выделение всей области.
SetSel(System\ViewPort\ViewArea, sel) : Else : SelectAllHTML(System\ViewPort) : EndIf : EndIf
Case #vWordWrap   : SetWordWrap(System\ViewPort, System\ViewPort\DataNode\WrapFlag ! 1)
EndSelect  ; Дабы не ставить метку.
LeaveCritical() ; Продолжаем блюсти.
EndProcedure
  
Procedure SetAnalyzis()
Define *HLNode = Index2Node(GetGadgetState(#ClipList))
If *HLNode <> System\LastAnalyzed : ConnectInformer() : System\LastAnalyzed = *HLNode : LinkOptics(*HLNode) : EndIf
; А теперь - насильственное включение:
If System\HoldField : SetDroppedState(System\Informator\InBox, #True) : System\HoldField = #False : EndIf ; Автоматически выбрасываем.
If GetAsyncKeyState_(#VK_SPACE) And GetActiveGadget() = #ClipList : System\ForceDrop = #True : SetDroppedState(System\Informator\InBox, #True) 
ElseIf System\ForceDrop : SetDroppedState(System\Informator\InBox, #False) : System\ForceDrop = #False : EndIf ; Прекращаем насилие, вот.
EndProcedure

Procedure DoTiming() ; Bindback.
Select EventTimer()
Case #tBackupTimer   : DoBackup()
Case #tRunTimer      : Define Win = EventWindow(), *VP.ViewPort = ExtractWP(WindowID(Win)) : ShiftInfo(*VP\Informator)
Case #tMainInfoTimer : ShiftInfo(System\Informator)
Case #tSnapTimer     : SnapVerse()
Case #tCollector     : GCollection()
Case #tSearchTimer   : SearchPulse() 
Case #tNoiseTimer    : NoiseGarden()
EndSelect
EndProcedure

Procedure MoveWatcher() ; Bindback.
SetSnapped(EventWindow(), #NoSnap, ElapsedMilliseconds())
SnapVerse()
EndProcedure

Procedure UseTray() ; Bindback.
DisableDebugger
Select EventType() ; Выбираем по событию...
Case #PB_EventType_LeftClick, #PB_EventType_LeftDoubleClick   : ReturnWindow()
Case #PB_EventType_RightClick, #PB_EventType_RightDoubleClick : ShowTrayMenu()
EndSelect
EnableDebugger
EndProcedure

Procedure HideMe() ; Bindback.
SendNotifyMessage_(#HWND_BROADCAST, System\HideMsg, 0, 0)
ForEach System\ViewPorts() ; Прикрываем все окна данных.
If System\ViewPorts()\WindowID : HideWindow(System\ViewPorts()\WindowID, #True)
Else : DeleteElement(System\ViewPorts()) ; Удаляем нод (да, тут).
EndIf
Next : HideWindow(#MainWindow, #True) : AddTray()
EndProcedure
;}
;} {End/Procedures}

;{ ==Preparations==
Repeat: System\DupMutex = CreateMutex_(0, 0, #SinglerMx) : If System\DupMutex : Break : Else : Delay(100) : EndIf : ForEver
If GetLastError_() : ErrorBox("program is already running !" + #CR$ + "Press 'OK' to exit.") : End : EndIf
CompilerIf Not #PB_Compiler_Debugger : OnErrorCall(@ErrorHandler()) ; Выставляем обработчик ошибок, коли нет своего.
CompilerEndIf : OpenIniFile() ; Открываем файл настроек...
; -Win preparations-
System\XPLegacy = Bool(OSVersion() = #PB_OS_Windows_XP Or OSVersion() = #PB_OS_Windows_8) * #WS_EX_COMPOSITED ; Ддя порядка.
If System\XPLegacy : System\SizeMsg = #WM_SIZE : Else : System\SizeMsg = #WM_WINDOWPOSCHANGED : EndIf 
If OSVersion() > #PB_OS_Windows_XP : System\Bullet = "•" : Else : System\Bullet = "" : EndIf ; Контекстно-зависимый круг.
If OSVersion() => #PB_OS_Windows_10 : System\SnapShift = 5 : EndIf     ; Корректор для стыкующихся окон.
System\LastSrc    = GetForegroundWindow_()                             ; Сохраняем окно, которое могло бы быть источником.
System\MainWindow = OpenWindow(#MainWindow, 0, 0, #MinWidth, #MinHeight, #Title, #WinFlags|#PB_Window_ScreenCentered)
System\OwnHandle  = GetWindowThreadProcessId_(System\MainWindow, 0)    ; Сохраняем указатель.
StickyWindow(#MainWindow, #True) : TextGadget(#VoidGadget, 0, 0, 0, 0, "[Drag[|]Drop]") ; Just to extract font.
SetContain(System\FlickBuf) : System\ListID = ListViewGadget(#ClipList, #ClipLOffset, #ClipUOffset, 0, 0)
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
RestoreList() : LoadFont(#fInFont, "Sylfaen", 9)                           ; Шрифт информационной строки.
EditorGadget(#RTFParser, 0, 0, 0, 0, #PB_Editor_ReadOnly)                  ; RTF->STR parser.
HideGadget(#RTFParser, #True) : System\RTFParser = GadgetID(#RTFParser)    ; Getting ptr.
RichEdit_SetInterface(System\RTFParser)                                    ; Полная синхронизация.
MakeTextPlain(System\RTFParser) : WebGadget(#HTMLParser, 0, 0, 0, 0, "")   ; HTML->STR parser.
Define WBrowser.IWebBrowser2 = GetWindowLongPtr_(GadgetID(#HTMLParser), #GWL_USERDATA) : WBrowser\put_Silent(#True) 
SetGadgetAttribute(#HTMLParser, #PB_Web_BlockPopups, #True)                ; На всякий случай.
System\ComBrowser = COMate_WrapCOMObject(WBrowser.IWebBrowser2)            ; Создаем враппер.
HideGadget(#HTMLParser, #True) : System\WebObject = GetWindowLongPtr_(GadgetID(#HTMLParser), #GWL_USERDATA) ; Getting ptr too.
; -Button bar GUI-
LoadFont(#fButtonFont, "Palatino Linotype", 10, #PB_Font_Bold)  ; Загружаеv шрифт кнопочной панели.
LoadFont(#fOCularFont, "Webdings", 20,  #PB_Font_HighQuality)   ; Загружаем спец. шрифт кнопки паноптикума.
DefBarButton(#Button_Ocular   , Chr(78), #fOCularFont, #PB_Button_Toggle)
DefBarButton(#Button_Clear    , "Clear")   : DefBarButton(#Button_Switch, "") 
DefBarButton(#Button_Options  , "Options") : DefBarButton(#Button_Terminate, "Terminate")
ContainerGadget(#EyeFrame, 0, 0, 0, #ButtonHeight + #EFSize, #PB_Container_Flat) : CloseGadgetList()
;;;;;;;;;;;;
; -Additional GUI-
StringGadget(#SearchBar, #SearchLOffset, #SearchUOffset, 10, #SearchHeight, "", #PB_String_BorderLess|#WS_BORDER|#PB_Text_Center)
ButtonGadget(#LurkBack , 0, #SearchUOffset, #SBWidth, #SBHeight, Chr(51), #BS_FLAT)
ButtonGadget(#LurkForth, 0, #SearchUOffset, #SBWidth, #SBHeight, Chr(52), #BS_FLAT)
LoadFont(#fSBarFont, "Verdana", 8, #PB_Font_Italic) : LoadFont(#fSButtonFont, "Marlett", 12)
LoadFont(#fBoxedFont, "Tahoma", 8) ; Совместимость для Win8 и тому подобного недоGUI.
SetGadgetFont(#SearchBar, FontID(#fSBarFont)) : SetGadgetFont(#LurkBack, FontID(#fSButtonFont))
SetGadgetFont(#LurkForth, FontID(#fSButtonFont)) : SetGadgetColor(#SearchBar, #PB_Gadget_BackColor, $EFEFEF)
InitInformer(System\Informator) : AddWindowTimer(#MainWindow, #tMainInfoTimer, 100) ;Строка информации.
EnableGadgetDrop(#SearchBar, #PB_Drop_Text, #DragOff) ; Включаем возможность бросать туда текст.
SetGadgetAttribute(#SearchBar, #PB_String_MaximumLength, #SearchLimit) ; Лимит, дабы не наглели.
System\SearchBar = GadgetID(#SearchBar) : SetSearcherDummy() : ChangeCB(System\SearchBar, SearchBarCallback())
; -Extra win preparations-
SetupBuffers(System\FlickBuf)
WindowBounds(#MainWindow, #MinWidth, #MinHeight, #PB_Ignore, #PB_Ignore) : SetWindowCallback(@WinCallback(), #MainWindow) 
SwitchStyle(System\MainWindow, #WS_EX_LAYERED|System\XPLegacy)                             ; Выставляем необходимы стили окна.
OpenWindow_SettingsWindow()     : StickyWindow(#SettingsWindow, #True) : SetOpacity()      ; Opaque right now.
ChangeCB(GadgetID(#Button_Front), ContainerCallback()) : ChangeCB(GadgetID(#Button_Back) , ContainerCallback())
ChangeCB(System\ListID          , DataListCallback()) : ;SetWindowColor(#MainWindow, #WinShade) 
; -Panopticum-
With System\Panopticum
InitVPWindow(\WindowID, \Stabilizer) : Define *Window = \WindowID, *VPort.ViewPort = System\Panopticum ; Основная аллокация.
Editorial(\PlainArea)   : MakeTextPlain(GadgetID(\PlainArea)) : Editorial(\RTFArea) : MetaPort(\MetaArea)
BMPort(\BitmapArea)     : WebPort(\HTMLArea) : \NoiseArea = ImageGadget(VASizings(), #Null, #PB_Image_Border)
\ViewArea = \NoiseArea  : ConfigureVP(System\Panopticum) ; Последние приготовления, да-да.
EndWith
; -Additional stuff from preferences-
RestoreVolatile() : UpdateSwitch() : AddWindowTimer(#MainWindow, #tBackupTimer, 15 * #Minute) ; Таймер резервного копирования.
; -Shortcuts-
ShortControl(Return, #cReturn, 0) ; Восстановление элемента под курсором. Ну, одно из. Альтернативное.
ShortControl(F3, #cFindNext, 0) : ShortControl(F3, #cFindPrev,  #PB_Shortcut_Shift)
ShortControl(F, #cGoSearch)   : ShortControl(P, #cPanopticum)                   ;<^ Управление поиском с клавиатуры.
ShortControl(S,     #cSaveAs) : ShortControl(V,      #cCtrlV) ; Сохранение (SaveAs)                / Показ (Show Data).
ShortControl(R, #cSetComment) : ShortControl(C,       #cCopy) ; Комментарий (Comment)              / Восстановление текущего элемента.
ShortControl(Insert,  #cCopy) : ShortControl(Up,    #cMoveUP) ; Восстановление текущего элемента   / Переставить вверх (Move Up).
ShortControl(Down,#cMoveDown) : ShortControl(O,    #cOptions) ; Переставить вверх (Move Down)      / Вызов панели опций.
ShortControl(Q, #cQWERTYSwap) : ShortControl(Delete,#cDelete, #Null)  ; Быстрая коррекция раскладки / Удаление (Remove).
ShortControl(PageUp, #cThrowUp) : ShortControl(PAgeDown, #cThrowDown) ; Броски элемента по списку вверх / вниз.
ShortControl(A,      #cCtrlA) ; Выделить все содержимое строки ввода.
; -Hotkeys-
For I = 1 To #HotKeys ; Инициализируем управление хоткеями.
RegisterHotKey_(System\MainWindow, I, #MOD_CONTROL, '1' - 1 + I) 
RegisterHotKey_(System\MainWindow, #Hype+I,  #CtrlShift, '1' - 1 + I) 
; ...Ну и новое слово - хоткеи сверхбыстрого доступа:
Next I : RegHypeKey(#BringKey) : RegHypeKey(#SwapKey) : RegHypeKey(#SaveKey)
; -Misc settings-
PreferenceGroup("Misc") ; Загружаем группу.
ReadPreferenceIntegerEX(System\Options\ListMax, "MaxEntries", 0, GetGadgetAttribute(#LimitSpin, #PB_Spin_Maximum))
ReadPreferenceIntegerEX(System\Options\SizeLimit, "SizeLimit", 0, GetGadgetAttribute(#SizeSpin, #PB_Spin_Maximum))
If System\Options\ListMax And ReadPreferenceInteger("PreserveHotkeys", #True):System\Options\HPreserve=#True:EndIf
RegBoolOption("HiddenStart"  , HideStart , HideStart)
RegBoolOption("GlassWindow"  , GlassWin  , GlassWin  , #True)
RegBoolOption("AlterPasting" , AltPasting, AltPasting, #True)
RegBoolOption("ListMimicry"  , MimicList , Mimic     , #True)
RegBoolOption("PaintFixer"   , FixPaint  , Fix       , #True)
RegBoolOption("ReRaise"      , AutoReboot, ReRaise   , #True)
RegBoolOption("UseDump"      , UseDump   , UseDump   , #True)
RegBoolOption("WipeAwareness", WipeWarn  , WipeWarn  , #True)
RegBoolOption("DumpImages"   , SaveImages, DumpImages)
RegBoolOption("DumpListings" , SaveLists , DumpLists)
RegBoolOption("SnapWindows"  , SnapWin   , Snap      , #True)
RegBoolOption("Hide2Tray"    , UseTray   , Tray      , #True)
RegBoolOption("FixTrayIcon"  , FixedTray , Fixed)
; -Prepare formats-
System\ClipRTF  = RegisterClipboardFormat_(#CF_RTF)
System\ClipHTML = RegisterClipboardFormat_("HTML Format")
RegFormat(5, #CF_TEXT       , "STR" , #PB_Drop_Text)    : RegFormat(4, #CF_HDROP     , "DIR", #PB_Drop_Files) 
RegFormat(3, System\ClipHTML, "HTML", System\ClipHTML)  : RegFormat(2, #CF_BITMAP    , "BMP", #PB_Drop_Image)
RegFormat(1, #CF_ENHMETAFILE, "META", #CF_METAFILEPICT) : RegFormat(0, System\ClipRTF, "RTF", System\ClipRTF)
Pref(#CF_HTML) = Pref(System\ClipHTML)                  : Pref(#CF_RichText) = Pref(System\ClipRTF)
SetDropCallback(@DropCB()) : EnableGadgetDrop(#ClipList, #PB_Drop_Private, #DragOff) ; Временно, до полной имплементации.
; -Последние приготовления-
System\SetupWindow = WindowID(#SettingsWindow) ; Получаем ID.
InitTable()                          ; Подготавливаем таблицу.
ClosePreferences() : WriteAllPrefs() ; Закрываем файл настроек.
RestoreDump() : CheckMaximum() : UpdateTitle() ; Восстанавливаем все.
System\CloseMsg   = RegisterWindowMessage_("Slippery_Close")   ; Регистрируем сообщение (close).
System\UpdateMsg  = RegisterWindowMessage_("Slippery_Adapt")   ; Регистрируем сообщение (update).
System\HideMsg    = RegisterWindowMessage_("Slippery_Hide")    ; Регистрируем сообщение (hide).
System\RenumMsg   = RegisterWindowMessage_("Slippery_Renum")   ; Регистрируем сообщение (renum).
System\RestoreMsg = RegisterWindowMessage_("Slippery_Restore") ; Регистрируем сообщение (restore).
ExtractIconEx_(ProgramFilename(), 0, 0, @System\AppIcon, 1) ; Готовим трей к использованию.
AllowSetForegroundWindow_(GetCurrentProcess_())             ; Странное оно, но да ладно.
SetForegroundWindow_(System\LastSrc)                        ; Выставляем на передний план.
System\NextWindow = SetClipboardViewer_(System\MainWindow)  ; Ставим перехватчик.
If System\Options\HideStart : AddTray() : Else  ; Сразу стартуем с иконкой:
If System\Options\FixedTray : AddTray() : EndIf : HideWindow(#MainWindow, #False)
SetForegroundWindow_(System\MainWindow) : SAG() : EndIf
HLLine(UsedNodeIdx()) : HadleFirstOpen() ; Ставим курсор на выделение, вот да.
System\LastAnalyzed = -1 : AddWindowTimer(#MainWindow, #tSnapTimer, #SnapTime) ; Всеобщее тяготение.
System\ActualReq = ""    ; Исправляем очень странную, но все же ошибку.
; -Extra timing & binding before loop-
AddWindowTimer(#MainWindow, #tCollector, #Minute) ; Сборщик мусора.
AddWindowTimer(#MainWindow, #tSearchTimer, 50)    ; Напоминание о необходимости поиска.
AddWindowTimer(#MainWindow, #tNoiseTimer, #Second / 30) ; Шумогенератор.
BindEvent(#PB_Event_Timer, @DoTiming() , #PB_All) : BindEvent(#PB_Event_MoveWindow , @MoveWatcher(), #PB_All)
BindEvent(#PB_Event_Menu,  @UseMenu()  , #PB_All) : BindEvent(#PB_Event_GadgetDrop , @GadgetDrop() , #MainWindow)
BindEvent(#PB_Event_SysTray, @UseTray(), #PB_All) : BindEvent(#PB_Event_CloseWindow, @HideMe()     , #MainWindow)
BindGadgetEvent(#SearchBar, @SearchDaemon(), #PB_EventType_Change)
BindGadgetEvent(#SearchBar, @SearchDaemon(), #PB_EventType_Focus)
BindGadgetEvent(#SearchBar, @SearchDaemon(), #PB_EventType_LostFocus)
; -Search flagi definition (targets)-
DefSearchFlag("text" , #sfText   , Target    , #sPlainText)
DefSearchFlag("src"  , #sfSrc    , Target    , #sWSource)
DefSearchFlag("rem"  , #sfRem    , Target    , #sRemark)
DefSearchFlag("id"   , #sfID     , Target    , #sDictID)
DefSearchFlag("list" , #sfList   , Target    , #sListLine)
; -Search flagi definition (exclusion)-
DefSearchFlag("~str" , #sfNoSTR   , DenySTR  , #CF_TEXT)
DefSearchFlag("~bmp" , #sfNoBMP   , DenyBMP  , #CF_BITMAP)
DefSearchFlag("~dir" , #sfNoDIR   , DenyDIR  , #CF_HDROP)
DefSearchFlag("~rtf" , #sfNoRTF   , DenyRTF  , #CF_RichText)
DefSearchFlag("~html", #sfNoHTML  , DenyHTML , #CF_HTML)
DefSearchFlag("~meta", #sfNoMETA  , DenyMETA , #CF_ENHMETAFILE)
DefSearchFlag("inv"  , #sfUnDeny  , DenyNot  , #True)
; -Search flagi definition (refining)-
DefSearchFlag("case"  , #sfCase   , CaseFlag  , #PB_String_CaseSensitive)
DefSearchFlag("~case" , #SfNocase , CaseFlag  , #PB_String_NoCase)
DefSearchFlag("reg"   , #sfRegular, RegFlag   , #True)
DefSearchFlag("whole" , #sfWhole  , WholeFlag , #True)
DefSearchFlag("open"  , #sfOpen   , OpenFlag  , #True)
DefSearchFlag("not"   , #sfNot    , DenialFlag, #True)
DefSearchFlag("sel"   , #sfSel    , HLFlag    , #True)
DefSearchFlag("~sel"  , #sfNoSel  , HLFlag    , #False)
DefSearchFlag("bound" , #sfHKBound  , BindingFlag, #HotBound)
DefSearchFlag("~bound", #sfHKUnbound, BindingFlag, #HotUnBound)
;} {End/Preparations}
;{ ==Main loop==
HideWindow(#MainWindow, #False)
With System\GUIEvent
Repeat : SetAnalyzis() : ReceiveEvent(System\GUIEvent)
If \Window = #MainWindow Or \Type = #PB_Event_Menu
Select \Type
Case #PB_Event_Gadget
; -----------
Select \Gadget
Case #ClipList : Select \SubType : Case #PB_EventType_LeftDoubleClick : RestoreData(GetGadgetState(#ClipList))
                                   Case #PB_EventType_DragStart       : DragNode(GetGadgetState(#ClipList))
EndSelect
Case #Button_Clear     : WipeData() : SAG()
Case #Button_Switch    : System\AcceptNew ! 1 : UpdateSwitch()
Case #Button_Ocular    : If PanVisibility() : BanePanopticum() : Else : BringPanopticum() : EndIf
Case #Button_Options   : BringOptions()
Case #Button_Terminate : Break
Case #LurkForth        : InitSearch()
Case #LurkBack         : InitSearch(#True)
EndSelect
; -----------
Case #WM_HOTKEY        : AnalyzeHotkey(EventwParam())
Case #WM_CHAR          : SelectKey()
EndSelect
ElseIf \Window = #SettingsWindow : OptionsLoop()
ElseIf \Type   = #WM_RBUTTONUP   ; Контекстное меню.
ViewPortMenu(ExtractWP(\Window))
EndIf
ForEver
EndWith
;} {End/Loop}
;{ ==AfterMath==
! AfterMath: ; Спасибо Фреду за наше счастливое детство !
SetWindowCallback(0, #MainWindow) : HideWindow(#SettingsWindow, #True) : HideWindow(#MainWindow, #True)
ChangeClipboardChain_(System\MainWindow, System\NextWindow)     ; Убираемся из цепи.
DisableDebugger : RemoveSysTrayIcon(#TrayIcon) : EnableDebugger ; Иконка в трее.
SendNotifyMessage_(#HWND_BROADCAST, System\CloseMsg, 0, 0) : DoBackUp() ; На прощание - сохраняем данные.
;} {End/AfterMath}
; IDE Options = PureBasic 5.40 LTS (Windows - x86)
; Folding = C6-v4--88-4-----+-4-f-8---+---8----+84-P+
; EnableUnicode
; EnableUser
; UseIcon = ClipBoard.ico
; Executable = ..\SlipperyClip.exe
; CurrentDirectory = ..\
; IncludeVersionInfo
; VersionField0 = 1,21,0,0
; VersionField1 = 1,21,0,0
; VersionField2 = Guevara-chan [~R.i.P]
; VersionField3 = Slippery Clip
; VersionField4 = 1.21
; VersionField5 = 1.21
; VersionField6 = Slippery Clip|board manager
; VersionField7 = Slippery Clip
; VersionField8 = SlipperyClip.exe
; VersionField9 = Copyleft (ɔ) 2010, Guevara-chan
; VersionField13 = Guevara-chan@Mail.ru
; VersionField14 = http://vk.com/guevara_chan