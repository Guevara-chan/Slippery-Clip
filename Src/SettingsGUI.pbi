;{- Enumerations / DataSections
;{ Windows
Enumeration
  #SettingsWindow
EndEnumeration
;}
;{ Gadgets
Enumeration
  #Frame3D_0
  #Text_1
  #Text_4
  #FontField
  #Button_Font
  #CheckMaxSize
  #SizeSpin
  #CheckLimit
  #LimitSpin
  #CheckPreserve
  #Frame3D_8
  #CheckUseDump
  #CheckWipeWarn
  #CheckDumpImages
  #Button_Front
  #Button_Back
  #CheckDumpLists
  #CheckGlassWin
  #CheckTray
  #CheckFixed
  #CheckHideStart
  #CheckFix
  #CheckReRaise
  #CheckAltPasting
  #CheckMimic
  #CheckSnap
  #Button_Cancel
  #Button_Reset
  #Button_Accept
EndEnumeration
;}
;{ Fonts
Enumeration
  #Font_Frame3D_0
  #Font_Text_1
  #Font_Text_4
  #Font_FontField
  #Font_Button_Font
  #Font_CheckMaxSize
  #Font_SizeSpin
  #Font_CheckLimit
  #Font_LimitSpin
  #Font_CheckPreserve
  #Font_Frame3D_8
  #Font_CheckUseDump
  #Font_CheckWipeWarn
  #Font_CheckDumpImages
  #Font_CheckDumpLists
  #Font_CheckGlassWin
  #Font_CheckTray
  #Font_CheckFixed
  #Font_CheckHideStart
  #Font_CheckFix
  #Font_CheckReRaise
  #Font_CheckAltPasting
  #Font_CheckMimic
  #Font_CheckSnap
  #Font_Button_Cancel
  #Font_Button_Reset
  #Font_Button_Accept
EndEnumeration
;}
;}
Procedure OpenWindow_SettingsWindow()
  If OpenWindow(#SettingsWindow, 509, 138, 290, 510, "=Settings=", #PB_Window_WindowCentered|#PB_Window_SystemMenu|#PB_Window_Tool|#PB_Window_TitleBar|#PB_Window_Invisible)
    FrameGadget(#Frame3D_0, 5, 0, 280, 165, "Data list:")
    TextGadget(#Text_1, 15, 20, 90, 20, "Front color:")
    TextGadget(#Text_4, 160, 20, 85, 20, "Back color:")
    CanvasGadget(#FontField, 15, 50, 230, 25)
    ButtonGadget(#Button_Font, 245, 50, 30, 25, "...", #BS_FLAT)
    CheckBoxGadget(#CheckMaxSize, 15, 80, 180, 25, "Data rejection bound (Kb):")
    GadgetToolTip(#CheckMaxSize, "Filter incoming traffic by size limitation. Current list stay intact.")
    SpinGadget(#SizeSpin, 200, 80, 75, 25, 1, 100500, #PB_Spin_Numeric)
    CheckBoxGadget(#CheckLimit, 15, 110, 185, 25, "Limit number of entries:")
    GadgetToolTip(#CheckLimit, "List would be truncated to new size.")
    SpinGadget(#LimitSpin, 200, 110, 75, 25, 1, 100500, #PB_Spin_Numeric)
    CheckBoxGadget(#CheckPreserve, 15, 135, 260, 25, "Force preserving hotkey-bound entries")
    GadgetToolTip(#CheckPreserve, "Not tested fully. Approach with caution.")
    FrameGadget(#Frame3D_8, 5, 170, 280, 301, "Miscellaneous:")
    CheckBoxGadget(#CheckUseDump, 15, 190, 225, 25, "Store data from list into dump file")
    CheckBoxGadget(#CheckWipeWarn, 15, 210, 235, 25, "Ask user before clearing data list")
    CheckBoxGadget(#CheckDumpImages, 15, 232, 258, 25, "Store images into dump file (expensive)")
    CheckBoxGadget(#CheckDumpLists, 15, 253, 210, 25, "Store file listings into dump file")
    CheckBoxGadget(#CheckGlassWin, 15, 274, 215, 25, "Become transparent if lost focus")
    GadgetToolTip(#CheckGlassWin, "50% alpha for unfocused window.")
    CheckBoxGadget(#CheckTray, 15, 295, 195, 25, "Show tray icon while hidden")
    CheckBoxGadget(#CheckFixed, 15, 315, 230, 25, "...And while not - as well (4Switch)")
    GadgetToolTip(#CheckFixed, "Miranda-style")
    CheckBoxGadget(#CheckHideStart, 15, 337, 170, 25, "Hide window on startup")
    CheckBoxGadget(#CheckFix, 15, 358, 260, 25, "Auto-fixing for Paint's selection")
    GadgetToolTip(#CheckFix, "Should be BMP, not META.")
    CheckBoxGadget(#CheckReRaise, 15, 379, 260, 25, "Auto-restart in case of irreparable error")
    CheckBoxGadget(#CheckAltPasting, 15, 400, 245, 25, "Bruteforce pasting emulation (Ctrl+C)")
    GadgetToolTip(#CheckAltPasting, "Use keyboard messages instead of WM_PASTE. More compatible.")
    CheckBoxGadget(#CheckMimic, 15, 421, 260, 25, "Use data list layout for plain text render")
    CheckBoxGadget(#CheckSnap, 15, 442, 260, 25, "Try aligning windows to each other")
    GadgetToolTip(#CheckSnap, "Experimental. May complicate viewport management.")
    ButtonGadget(#Button_Cancel, 5, 480, 80, 25, "Cancel", #BS_FLAT)
    ButtonGadget(#Button_Reset, 105, 480, 80, 25, "Reset", #BS_FLAT)
    ButtonGadget(#Button_Accept, 205, 480, 80, 25, "Accept", #BS_FLAT)
    ContainerGadget(#Button_Front, 105, 20, 30, 25, #PB_Container_Double)
    CloseGadgetList()
    ContainerGadget(#Button_Back, 245, 20, 30, 25, #PB_Container_Double)
    CloseGadgetList()
    ; Gadget Fonts
    SetGadgetFont(#Frame3D_0, LoadFont(#Font_Frame3D_0, "Palatino Linotype", 10, #PB_Font_HighQuality))
    SetGadgetFont(#Text_1, LoadFont(#Font_Text_1, "Palatino Linotype", 12, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#Text_4, LoadFont(#Font_Text_4, "Palatino Linotype", 12, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#FontField, LoadFont(#Font_FontField, "Palatino Linotype", 12, #PB_Font_Bold|#PB_Font_Italic|#PB_Font_HighQuality))
    SetGadgetFont(#Button_Font, LoadFont(#Font_Button_Font, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckMaxSize, LoadFont(#Font_CheckMaxSize, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#SizeSpin, LoadFont(#Font_SizeSpin, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckLimit, LoadFont(#Font_CheckLimit, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#LimitSpin, LoadFont(#Font_LimitSpin, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckPreserve, LoadFont(#Font_CheckPreserve, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#Frame3D_8, LoadFont(#Font_Frame3D_8, "Palatino Linotype", 10, #PB_Font_HighQuality))
    SetGadgetFont(#CheckUseDump, LoadFont(#Font_CheckUseDump, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckWipeWarn, LoadFont(#Font_CheckWipeWarn, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckDumpImages, LoadFont(#Font_CheckDumpImages, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckDumpLists, LoadFont(#Font_CheckDumpLists, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckGlassWin, LoadFont(#Font_CheckGlassWin, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckTray, LoadFont(#Font_CheckTray, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckFixed, LoadFont(#Font_CheckFixed, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckHideStart, LoadFont(#Font_CheckHideStart, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckFix, LoadFont(#Font_CheckFix, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckReRaise, LoadFont(#Font_CheckReRaise, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckAltPasting, LoadFont(#Font_CheckAltPasting, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckMimic, LoadFont(#Font_CheckMimic, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#CheckSnap, LoadFont(#Font_CheckSnap, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#Button_Cancel, LoadFont(#Font_Button_Cancel, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#Button_Reset, LoadFont(#Font_Button_Reset, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    SetGadgetFont(#Button_Accept, LoadFont(#Font_Button_Accept, "Palatino Linotype", 10, #PB_Font_Bold|#PB_Font_HighQuality))
    ;
    SetClassLongPtr_(GadgetID(#Button_Front), #GCL_STYLE, #CS_DBLCLKS)
    SetClassLongPtr_(GadgetID(#Button_Back), #GCL_STYLE, #CS_DBLCLKS)
    SetProp_(GadgetID(#Button_Front), "Gadget", #Button_Front)
    SetProp_(GadgetID(#Button_Back) , "Gadget", #Button_Back)
    SetProp_(GadgetID(#Button_Front), "ChooserTitle", @"Choose text coloring:")
    SetProp_(GadgetID(#Button_Back) , "ChooserTitle", @"Choose background coloring:")
    #bCancel = 1<<10 : AddKeyboardShortcut(#SettingsWindow, #PB_Shortcut_Escape, #bCancel)
    #bAccept = 2<<10 : AddKeyboardShortcut(#SettingsWindow, #PB_Shortcut_Return, #bAccept)
  EndIf
EndProcedure

; IDE Options = PureBasic 5.30 (Windows - x86)
; Folding = -
; EnableXP