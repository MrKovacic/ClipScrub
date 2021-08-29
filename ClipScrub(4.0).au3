#NoTrayIcon
    #include <GUIConstants.au3>
	#include <GuiRichEdit.au3>
	#include <GuiEdit.au3>
	#include <ScrollBarsConstants.au3>
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=paste.ico
#AutoIt3Wrapper_Outfile=ClipScrup.exe
#AutoIt3Wrapper_Res_Description=Cleans extraneous formatting from text in the clipboard and more!
#AutoIt3Wrapper_Res_Fileversion=4.1
#AutoIt3Wrapper_Res_LegalCopyright=2015 EHR and Mike Kovacic
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
Global $Version = 4.1
#include <GUIConstantsEx.au3>
#include <TrayConstants.au3>
#include <Misc.au3>
Opt("TrayMenuMode", 3)
Opt("GUIOnEventMode", 0)
Global $hDLL = DllOpen("user32.dll")
Global $option1 = 0
Global $option2 = 0
Global $option3 = 0
Global $option4 = 0
Global $option5 = 1
Global $option6 = 0
Global $option7 = 0
Global $option8 = 0
Global $Manual = 0
If _Singleton(@ScriptName, 0) = 0 Then
	MsgBox(16, "Oops", "Only one instance of ClipScrub may run at a time, and ClipScrub is already running.")
	Exit
EndIf
Global $iSettings = TrayCreateMenu("Extra Settings")
Global $iOP4 = TrayCreateItem("Make all text UPPERCASE", $iSettings)
Global $iOP1 = TrayCreateItem("Make all text lowercase", $iSettings)
Global $iOP2 = TrayCreateItem("Remove all Duplicate Spaces", $iSettings)
Global $iOP3 = TrayCreateItem("Remove blank lines", $iSettings)
Global $iOP5 = TrayCreateItem("Show Tray Tips", $iSettings)
Global $iHotKey = TrayCreateMenu("HotKey")
Global $iCS = TrayCreateItem("Ctrl + Space", $iHotKey) ;11 20
Global $iSS = TrayCreateItem("Shift + Space", $iHotKey) ;10 20
Global $CMW = TrayCreateItem("Ctrl + Mouse Wheel", $iHotKey) ;11 04
Global $AMS = TrayCreateItem("Alt + Mouse Wheel", $iHotKey) ;12 04
Global $iSMW = TrayCreateItem("Shift + Mouse Wheel", $iHotKey) ;10 04
TrayCreateItem("")
Global $ON = TrayCreateItem("ClipScrub ON Auto")
Global $ONMan = TrayCreateItem("ClipScrub ON Manual")
Global $OFF = TrayCreateItem("ClipScrub OFF")
TrayCreateItem("")
Local $iHelp = TrayCreateItem("Help")
Local $idAbout = TrayCreateItem("About")
TrayCreateItem("")
Local $idExit = TrayCreateItem("Exit")
TraySetState($TRAY_ICONSTATE_SHOW)
TraySetToolTip("ClipScrub is ON.")
TrayTip("ClipScrub", "ClipScrub is on automatic.", 5, 0)
;TraySetIcon("shell32.dll", 138)
AdlibRegister("clean", 100)
TrayItemSetState($ON, $TRAY_CHECKED)
TrayItemSetState($iCS, $TRAY_CHECKED)
TrayItemSetState($iOP5, $TRAY_CHECKED)
While 1
keys()
	Switch TrayGetMsg()
		Case $iOP1
		TRAYITEMSETSTATE($IOP4, $TRAY_UNCHECKED)
		$OPTION4 = 0
			If TrayItemGetState($iOP1) = ($TRAY_ENABLE + $TRAY_CHECKED) Then
				$option1 = 0
				TrayItemSetState($iOP1, $TRAY_UNCHECKED)
			ElseIf TrayItemGetState($iOP1) = ($TRAY_ENABLE + $TRAY_UNCHECKED) Then
				$option1 = 1
				TrayItemSetState($iOP1, $TRAY_CHECKED)
			EndIf
			
			
		Case $iOP2
			If TrayItemGetState($iOP2) = ($TRAY_ENABLE + $TRAY_CHECKED) Then
				$option2 = 0
				TrayItemSetState($iOP2, $TRAY_UNCHECKED)
			ElseIf TrayItemGetState($iOP2) = ($TRAY_ENABLE + $TRAY_UNCHECKED) Then
				$option2 = 1
				TrayItemSetState($iOP2, $TRAY_CHECKED)
			EndIf
		Case $iOP5
			If TrayItemGetState($iOP5) = ($TRAY_ENABLE + $TRAY_CHECKED) Then
				$option5 = 0
				TrayItemSetState($iOP5, $TRAY_UNCHECKED)
			ElseIf TrayItemGetState($iOP5) = ($TRAY_ENABLE + $TRAY_UNCHECKED) Then
				$option5 = 1
				TrayItemSetState($iOP5, $TRAY_CHECKED)
			EndIf
		Case $iOP3
			If TrayItemGetState($iOP3) = ($TRAY_ENABLE + $TRAY_CHECKED) Then
				$option3 = 0
				TrayItemSetState($iOP3, $TRAY_UNCHECKED)
			ElseIf TrayItemGetState($iOP3) = ($TRAY_ENABLE + $TRAY_UNCHECKED) Then
				$option3 = 1
				TrayItemSetState($iOP3, $TRAY_CHECKED)
			EndIf
		Case $iOP4
				TRAYITEMSETSTATE($IOP1, $TRAY_UNCHECKED)
		$OPTION1 = 0
			If TrayItemGetState($iOP4) = ($TRAY_ENABLE + $TRAY_CHECKED) Then
				$option4 = 0
				TrayItemSetState($iOP4, $TRAY_UNCHECKED)
			ElseIf TrayItemGetState($iOP4) = ($TRAY_ENABLE + $TRAY_UNCHECKED) Then
				$option4 = 1
				TrayItemSetState($iOP4, $TRAY_CHECKED)
			EndIf
		Case $iCS
			uncheckall()
			TrayItemSetState($iCS, $TRAY_CHECKED)
		Case $iSS
			uncheckall()
			TrayItemSetState($iSS, $TRAY_CHECKED)
		Case $CMW
			uncheckall()
			TrayItemSetState($CMW, $TRAY_CHECKED)
		Case $AMS
			uncheckall()
			TrayItemSetState($AMS, $TRAY_CHECKED)
		Case $iSMW
			uncheckall()
			TrayItemSetState($iSMW, $TRAY_CHECKED)
		Case $idAbout ;
			MsgBox(0, "About", "ClipScrub" & @CRLF & @CRLF & "Version: " & $Version & @CRLF & "Description: Cleans extraneous formatting from text in the clipboard and more!" & @CRLF & @CRLF & "Support / Opinions / complaints / Questions: MikeKovacic@gmail.com")		
		Case $iHelp ;
				help()
		Case $idExit ; Exit the loop.
			ExitLoop
		Case $ONMan
		uncheckonoff()
		TrayItemSetState($ONMan, $TRAY_CHECKED)
		AdlibUnRegister("clean")
		$manual = 1
		TraySetToolTip("ClipScrub is ON Manual.")
		if $option5 = 1 then TrayTip("ClipScrub", "ClipScrub is ON Manual.", 5, 0)
		Case $ON
		$manual = 0
		uncheckonoff()
		TrayItemSetState($ON, $TRAY_CHECKED)
		AdlibUnRegister("clean")
		AdlibRegister("clean", 100)
		TraySetToolTip("ClipScrub is ON Automatic.")
		if $option5 = 1 then TrayTip("ClipScrub", "ClipScrub is ON Automatic.", 5, 0)
		Case $OFF
		$manual = 0
		uncheckonoff()
		TrayItemSetState($OFF, $TRAY_CHECKED)
		AdlibUnRegister("clean")
		TraySetToolTip("ClipScrub is OFF.")
		if $option5 = 1 then TrayTip("ClipScrub", "ClipScrub is OFF.", 5, 0)
	EndSwitch
WEnd
Func clean()
	Local $sData = ClipGet()
	If $option1 = 1 Then $sData = StringLower($sData)
	If $option2 = 1 Then $sData = StringReplace($sData, "  ", " ")
	If $option3 = 1 Then $sData = _StringReplaceBlank($sData, 0)
	If $option4 = 1 Then $sData = Stringupper($sData)
	ClipPut($sData)
	if $manual = 1 AND $option5 = 1 then TrayTip("ClipScrub", "Clipboard has been scrubbed.", 2, 0)
EndFunc   ;==>clean
Func uncheckall()
	TrayItemSetState($iCS, $TRAY_UNCHECKED)
	TrayItemSetState($iSS, $TRAY_UNCHECKED)
	TrayItemSetState($CMW, $TRAY_UNCHECKED)
	TrayItemSetState($AMS, $TRAY_UNCHECKED)
	TrayItemSetState($iSMW, $TRAY_UNCHECKED)
EndFunc   ;==>uncheckall
Func uncheckonoff()
	TrayItemSetState($ON, $TRAY_UNCHECKED)
	TrayItemSetState($ONMan, $TRAY_UNCHECKED)
	TrayItemSetState($OFF, $TRAY_UNCHECKED)
EndFunc
Func _StringReplaceBlank($sString, $sSpaces = "")
	If $sSpaces Then
		$sSpaces = "\s*"
	EndIf
	$sString = StringRegExpReplace($sString, "(?s)\r\n" & $sSpaces & "\r\n", @CRLF)
	If @extended Then
		Return SetError(0, @extended, $sString)
	EndIf
	Return SetError(1, 0, $sString)
EndFunc   ;==>_StringReplaceBlank
Func Keys()
	If $Manual = 1 and TrayItemGetState($iCS)  = ($TRAY_ENABLE + $TRAY_CHECKED) And _IsPressed(11, $hDLL) And _IsPressed(20, $hDLL) Then clean()
	If $Manual = 1 and TrayItemGetState($iSS)  = ($TRAY_ENABLE + $TRAY_CHECKED) And _IsPressed(10, $hDLL) And _IsPressed(20, $hDLL) Then clean()
	If $Manual = 1 and TrayItemGetState($CMW)  = ($TRAY_ENABLE + $TRAY_CHECKED) And _IsPressed(11, $hDLL) And _IsPressed(04, $hDLL) Then clean()
	If $Manual = 1 and TrayItemGetState($AMS)  = ($TRAY_ENABLE + $TRAY_CHECKED) And _IsPressed(12, $hDLL) And _IsPressed(04, $hDLL) Then clean()
	If $Manual = 1 and TrayItemGetState($iSMW) = ($TRAY_ENABLE + $TRAY_CHECKED) And _IsPressed(10, $hDLL) And _IsPressed(04, $hDLL) Then clean()
EndFunc   ;==>Keys


func Help()
$hGui = GUICreate("Help",600,600) ; will create a dialog box that when displayed is centered
    Local $idMyedit = _GUICtrlRichEdit_Create($hGui,"", 0, 0, 600, 600, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL))
    GUISetState(@SW_SHOW)
	dim $Data
_GUICtrlRichEdit_WriteLine($idMyedit, "ClipScrub " & $Version & @CRLF, +12, "+bo", 0x000000)
_GUICtrlRichEdit_WriteLine($idMyedit, "ClipScrub is a simple application designed with the simple purpose of scrubbing and / or processing the text you copy. When you copy text from a word document, or a web page, there may be hidden formatting with that text, even though you can’t see it. Sometimes this causes document nightmares like irregular formatting, extraneous spaces and many other problems." & @CRLF, -10, "-bo", 0x000000)
_GUICtrlRichEdit_WriteLine($idMyedit, "ClipScrub " & $Version & " has 3 modes, Auto, Manual and Off. In Auto mode, ClipScrub is constantly scrubbing extraneous ‘invisible’ formatting from any text you copy from any source. This will ensure any document where you paste your text will not have any formatting from the previous document.  In Manual mode, ClipScrub is still running, however it is not keeping a constant eye on your clipboard, but will respond to whatever hotkey you choose ( Ctrl + Space by default ). By right clicking on the little paint brush icon (ClipScrub) on your task bar, you can choose what hotkey will tell ClipScrub to scrub your clipboard.  This is handy when using Office programs like Word, which are sensitive to automated changes in the clipboard. " & @CRLF, 0, "", 0x000000)
_GUICtrlRichEdit_WriteLine($idMyedit, "While in auto mode, Microsoft Word may not show the right click menu while in a document. Should this happen, switch to Manual mode and use a hot key to scrub." & @CRLF, 0, "", 0xff0000)
_GUICtrlRichEdit_WriteLine($idMyedit, "While ClipScrub is on, whether you use a hotkey in manual mode, or simply use the auto mode, ClipScrub will clean all hidden formatting from copied text. There are extra settings bundled with ClipScrub that can process the clipboard further, which can be accessed by right clicking on the ClipScrub icon in your taskbar, and choosing Extra Settings." & @CRLF, 0, "", 0x000000)
_GUICtrlRichEdit_WriteLine($idMyedit, "Features include:" & @CRLF, 0, "+bo", 0x000000)
_GUICtrlRichEdit_WriteLine($idMyedit, "Make all text UPPERCASE - Makes all text copied, paste as uppercase." & @CRLF, 0, "-bo", 0x000000)
_GUICtrlRichEdit_WriteLine($idMyedit, "Make all text lowercase - Makes all text copied, paste as lowercase." & @CRLF, 0, "", 0x000000)
_GUICtrlRichEdit_WriteLine($idMyedit, "Remove all Duplicate Spaces - Removes all duplicate space characters, while still leaving single spaces." & @CRLF, 0, "", 0x000000)
_GUICtrlRichEdit_WriteLine($idMyedit, "Remove blank lines - Removes all extraneous empty line breaks and leaving only one in its place." & @CRLF, 0, "", 0x000000)
_GUICtrlRichEdit_WriteLine($idMyedit, "Should you have any problems, email MikeKovacic@gmail.com" & @CRLF, 0, "", 0x000000)

_GUICtrlEdit_Scroll($idMyedit, $SB_PAGEUP)
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                ExitLoop

        EndSwitch
    WEnd
    GUIDelete()
Endfunc
Func _GUICtrlRichEdit_WriteLine($hWnd, $sText, $iIncrement = 0, $sAttrib = "", $iColor = -1)

    ; Count the @CRLFs
    StringReplace(_GUICtrlRichEdit_GetText($hWnd, True), @CRLF, "")
    Local $iLines = @extended
    ; Adjust the text char count to account for the @CRLFs
    Local $iEndPoint = _GUICtrlRichEdit_GetTextLength($hWnd, True, True) - $iLines
    ; Add new text
    _GUICtrlRichEdit_AppendText($hWnd, $sText & @CRLF)
    ; Select text between old and new end points
    _GuiCtrlRichEdit_SetSel($hWnd, $iEndPoint, -1)
    ; Convert colour from RGB to BGR
    $iColor = Hex($iColor, 6)
    $iColor = '0x' & StringMid($iColor, 5, 2) & StringMid($iColor, 3, 2) & StringMid($iColor, 1, 2)
    ; Set colour
    If $iColor <> -1 Then _GuiCtrlRichEdit_SetCharColor($hWnd, $iColor)
    ; Set size
    If $iIncrement <> 0 Then _GUICtrlRichEdit_ChangeFontSize($hWnd, $iIncrement)
    ; Set weight
    If $sAttrib <> "" Then _GUICtrlRichEdit_SetCharAttributes($hWnd, $sAttrib)
    ; Clear selection
    _GUICtrlRichEdit_Deselect($hWnd)

EndFunc