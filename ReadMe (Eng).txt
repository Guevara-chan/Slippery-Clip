*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-*
Title: Slippery Clip
Version: v1.21 (Release)
Distribution: FreeWare OpenSource
Dev. environment: PureBASIC v5.30
Dependencies: COMate+ (bundled)
*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-*

[*] Slippery Clip is a compact and easy-to-use tool for expanding clipboard's functional.
[*] Aside from optional tray icon, program's window could be returned to screen with Ctrl+Shif+T combination.
[*] Double click to list item puts it in the clipboard. Such element would be marked by arrow (->).
[*] Deleting the active element unloads it from the clipboard, leaving it empty.
[*] Right-clicking on list brings context menu for selected item.
[*] To bind hotkey for selected element, use Ctrl+[1-9] combination. Associated button would appear before it's ID.
[*] Using a Ctrl+[1-9] combination outside of program's window would place associated element into clipboard to be auto-pasted into the active application.
[*] Pressing 1-9 keys in program's window would place associated element into clipboard.
[*] Aside from mentioned above, Slippery Clip also supports the following shortcuts: Enter/Ctrl+C/Ctrl+Ins to paste selected element into clipboard, Ctrl+O to show options, Ctrl+F to set focus on search bar, and F3/Shift+F3 to find requested data starting after/before current selection. Arrow could also be used together with Control, moving element up and down respectively.
[*] Keyboard shortcuts Ctrl+P и Ctrl+V allows swift access to data of highlighted note, displaying nodal or shared overview windows correspondingly.
[*] Among other things, also possible to access certain features without ever touching window. There, Shift+Ctrl+[1-9] assigns hotkey to current content of clipboard, Shift+Ctrl+Q inverts its layout, while Shift+Ctrl+S offer saving dialogue. On a side node, Space could be used for bringing info breakdown about highlighted node and Shift+ESC to terminate searching prematurely.
[*] For data viewer, keys '+' and '-' can be used to resize window, while Ctrl+R - instead of menu item "Raw Copy".
[*] This application selects data from clipboard with following priority sequence: RTF>META>BMP>HTML>STR.
[*] Size of metafile calculated with ACDSee's method, ie by bounds difference (rclBounds).
[*] By default, search interface assumes [text/~case] as header to look in text data without considering cases and highlight results in viewports.
[*] To ease handling of characters outside the stand sets, Slippery Clip’ search engine offers basic reparsing mechanism. Replaceable sequences always begins with ` and offers following combinations:
<*> `` = ` (sequencer self-isolation).
<*> `~ = Alt+010 (line feed symbol).
<*> `| = Alt+013 (carriage return symbol).
<*>`#%hex%%hex% - character insertion by it’s ASCII-code (as in, `#09 equivalent to tabulation char).
<*>`$%hex%%hex%%hex%%hex% character insertion by it’s Unicode (as in, `#$0046 do become F letter).
<*> Any sequences beyond given above, including attempts of zero-char insertion (`#00/`$0000), is interpreted by their primordial layout.
[*] As for now, search engine autoreplace following characters with reparsing codes: Alt+009, Alt+010 and Alt+013.

P.S.  Once you got nothing better to do, check http://www.pcre.org/pcre.txt - it's fun enough.
