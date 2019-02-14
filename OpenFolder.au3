#include <GuiConstants.au3>
#include <File.au3>
#include <Array.au3>
#include <GuiListView.au3>
#include <IE.au3> 
Opt("GUICoordMode", 2)  
;Disable tray menu so script cannot be accidentally paused by clicking tray icon
Opt("TrayMenuMode",1)

; Init double click vars
$dbl_oldx=0
$dbl_oldy=0
$dbl_StartTime=0

$windowtitle="OpenFolder 1.6"

; If no or more than two parameters are supplied return an error (two is the max and one is optional)
If $CmdLine[0]=0 OR $CmdLine[0]>2 then
	MsgBox(4096 + 16,$windowtitle,"Usage: OpenFolder.exe <searchfolder> <optional search query>" & @LF & "(use quotes around the optional search query)")
	Exit
EndIf

;If more than one parameter is supplied check if the first is an existing directory
;Take first command line parameter as a $searchfolder
$searchfolder = $CmdLine[1]

;Strip trailing backslashes from working directory
While StringRight($searchfolder,1)="\"
	$searchfolder=StringTrimRight ($searchfolder,1)
Wend
;Add ONE trailing slash to the searchfolder
$searchfolder = $searchfolder & "\"

;Get attributes of searchfolder (used to check for existance and for verifying it's a folder and not a file)
$attrib = FileGetAttrib($searchfolder)
If @error Then
	MsgBox(4096 + 16,$windowtitle,"Supplied folder (" & $searchfolder & ") cannot be located." & @LF & "Usage: OpenFolder.exe <searchfolder> <optional search query>")
	Exit
Else
	If NOT StringInStr($attrib, "D") Then
		MsgBox(4096 + 16,$windowtitle,"Supplied folder (" & $searchfolder & ") seems to be a file instead of a folder." & @LF & "Usage: OpenFolder.exe <searchfolder> <optional search query>")
		Exit
		EndIf
EndIf

;Change the working directory to the searchfolder
FileChangeDir($searchfolder)

; If a second parameter is supplied, set this as the first search query to execute immediatly
If $CmdLine[0]=2 then
	$CommandLineSearchQuery = $CmdLine[2]
EndIf

$windowtitle=$windowtitle & " (" & $searchfolder & ")"

; Create the GUI window and controls
$Gui = GuiCreate($windowtitle,400,500)
$Label_1 = GuiCtrlCreateLabel("Type a part of the foldername:", 2,2, 150, 20, 0x1000)
$Input = GUICtrlCreateInput ("", 0, -1, 200, 20, 0x1000)
$Button1 = GuiCtrlCreateButton("Locate", 0, -1, 50, 20)

$size = WinGetClientSize(AutoItWinGetTitle())

Dim $listview, $Btn_FindExact, $Btn_FindPartial, $Btn_FindByID, $Btn_Exit, $msg, $ret, $input_find, $index, $item[5]
$listview = GUICtrlCreateListView("Foldername", -400, 0, 400, 480, BitOR($LVS_SINGLESEL, $LVS_SHOWSELALWAYS, $LVS_NOSORTHEADER))
GUICtrlSendMsg($listview, $LVM_SETEXTENDEDLISTVIEWSTYLE, $LVS_EX_GRIDLINES, $LVS_EX_GRIDLINES)
GUICtrlSendMsg($listview, $LVM_SETEXTENDEDLISTVIEWSTYLE, $LVS_EX_FULLROWSELECT, $LVS_EX_FULLROWSELECT)
_GUICtrlListView_SetColumnWidth ($listview, 0, $LVSCW_AUTOSIZE_USEHEADER)

; Contextmenu for clipboard
$ContextMenu = GUICtrlCreateContextMenu($listview)
$ContextOpenFolder = GUICtrlCreateMenuitem ("Open folder",$ContextMenu)
; next one creates a menu separator (line)
GuiCtrlCreateMenuitem ("",$ContextMenu)
$ContextCopyToClipboard = GUICtrlCreateMenuitem ("Copy nr. to clipboard",$ContextMenu)
; Read ini file for extra context menu options
$ContextMenus = IniReadSectionNames(@ScriptDir & "\OpenFolder.ini")
If NOT @error Then
	For $i = 1 To $ContextMenus[0]
        $ContextMenus[$i] = GUICtrlCreateMenuitem ($ContextMenus[$i],$ContextMenu)
    Next
EndIf

$SmoothScroll = RegRead("HKCU\Control Panel\Desktop","SmoothScroll")
RegWrite("HKCU\Control Panel\Desktop","SmoothScroll","REG_DWORD",0)
EnvUpdate()

; Execute a command line search query, if any
If IsDeclared("CommandLineSearchQuery") Then
	GUICtrlSetData($Input,$CommandLineSearchQuery)
	$Data = GUICtrlRead($Input)
	RefreshList()
	GUICtrlSetState($Input,$GUI_FOCUS)
EndIf

; Run the GUI until it is closed
GuiSetState()
While 1
    $msg = GuiGetMsg()
    Switch $msg
	Case $GUI_EVENT_CLOSE
        ExitLoop
    ;When button is pressed, label text is changed
    ;to combobox value
	Case $Button1
		$Data = GUICtrlRead($Input)
		RefreshList()
		GUICtrlSetState($Input,$GUI_FOCUS)
	Case $ContextOpenFolder
		OpenFolder(_GUICtrlListView_GetItemTextString($listview, _GUICtrlListView_GetNextItem ( $listview )))
	Case $ContextCopyToClipboard
		$Number = StringRegExp(_GUICtrlListView_GetItemTextString($listview, _GUICtrlListView_GetNextItem ( $listview )),'([0-9]{6})',1)
		; If a number was retrieved then put in on the clipboard
		If IsArray($Number) Then ClipPut($Number[0])
	Case $GUI_EVENT_PRIMARYDOWN
		mouseClicks()
	Case $Input
		$Data = GUICtrlRead($Input)
		RefreshList()
		GUICtrlSetState($Input,$GUI_FOCUS)
	Case $listview
		OpenFolder(_GUICtrlListView_GetItemTextString($listview, _GUICtrlListView_GetNextItem ( $listview )))
	Case Else
		; Check if a clicked item is a custom context menu item
		If IsArray($ContextMenus) Then
			For $element In $ContextMenus
				If $element = $msg Then
					; Try to find a number in the selected item
					$Number = StringRegExp(_GUICtrlListView_GetItemTextString($listview, _GUICtrlListView_GetNextItem ( $listview )),'([0-9]{6})',1)
					; If a number was retrieved then open website
					If IsArray($Number) Then
						; Read the website from the ini file using the text of the selected menu item
						$url=IniRead(@ScriptDir & "\OpenFolder.ini", GUICtrlRead ($element,1), "URL", "") & $Number[0]
						; If the ini read deliverd a value then open up a browser window
						If NOT ($url = "" & $Number[0]) Then ShellExecute($url)
					EndIf	
				EndIf
			Next
		EndIf
    EndSwitch
WEnd

GUIDelete()

Exit

Func RefreshList()
; Clear the visible listview
_GUICtrlListView_DeleteAllItems ($listview)

; Read all the folders
$FileList=_FileListToArray($searchfolder,"*" & $Data & "*",2)

; Remove first element from array (contains the number of folders, a counter)
_ArrayDelete($FileList,0)
; Check if _ArrayDelete returned an error, if so there were probably no results
; Then loop through the array elements and add them to a visible list
If Not (@error) Then
	For $element In $FileList
		GUICtrlCreateListViewItem($element, $listview)
	Next
EndIf
EndFunc

Func OpenFolder($FolderToOpen)
	If $FolderToOpen="" Then
			MsgBox(4096 + 16,$windowtitle,"No folder selected")
			Return
	EndIf
	Run(@ComSpec & " /c start """" " & """" & $searchfolder & $FolderToOpen & """", "", @SW_HIDE)
EndFunc

; Check whether listview is doubleclicked
Func _DoubleClicked($winhandle,$ControlHandle=0)
        Local $info= GUIGetCursorInfo ($winhandle)
        Local $diff = TimerDiff($dbl_startTime)    
        Local $mousespeed = RegRead("HKCU\Control Panel\Mouse","DoubleClickSpeed")
        If $mousespeed = "" Then $mousespeed = 500
        
        IF $diff < $mousespeed and $dbl_oldx=$info[0] and $dbl_oldy=$info[1] Then
           $dbl_oldx=0
           $dbl_oldy=0
           $dbl_StartTime=0
           If $ControlHandle=0 Then
               return $info[4]
           Else
             if $ControlHandle=$info[4] Then
                 return True
             Else
                 return False
             Endif
           Endif
       Else
            $dbl_oldx=$info[0]
            $dbl_oldy=$info[1]
            $dbl_StartTime=TimerInit()
        Endif
EndFunc    

Func mouseClicks()    
    Select
    Case _DoubleClicked($Gui,$ListView)
        OpenFolder(_GUICtrlListView_GetItemText($listview, _GUICtrlListView_GetNextItem ( $listview )))
    EndSelect
Endfunc