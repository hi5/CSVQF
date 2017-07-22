/*
Name          : CSV Quick filter
Purpose       : Load a CSV and use search criteria to filter the list
Version       : 0.49q
Updated       : 20122410
AHK Forum     : https://autohotkey.com/boards/viewtopic.php?f=6&t=34867
Requirements  : CSV Library [lib] https://github.com/hi5/CSV
                Also uses code by jsherk http://www.autohotkey.com/forum/viewtopic.php?t=73246
                and Attach by majkinetor http://www.autohotkey.com/forum/topic53317.html&highlight=forms
                <warning>Code needs some serious cleanup as it is a bit of (h|n)asty hack job</warning>
*/

; ini/setup

#NoEnv
#SingleInstance, Force
SetBatchLines,-1
Version=0.49q

Menu, tray, icon, shell32.dll, 23 ; search icon

MyTextStart:="Change column order for presentation in listview"
;SearchIndicator=/-\|
;StringSplit, si, SearchIndicator
;IndicatorC=1
SearchMethod=0
AppWindow=CSV Quick filter - %version%
GroupAdd, AppTitle, %AppWindow%
MaxRes=100
Menu, Tray, Tip, %AppWindow%
Mute=1
HelpForum=
(
Visit the forum at www.autohotkey.com/forum/ to report any bugs,
feature requests or the latest updates of %AppWindow%
You can post as guest, no registration required.
[search for CSVQuickfilter or CSVQF on the forum]

This program is written in AutoHotkey, a free, open-
source (scripting) utility for Windows, you can learn
more at www.autohotkey.com
)

; check for cmdline parameters
; file=param1 delimiter=param2 header=param3 colums=param4
CmdlineOpt=
(
Command line options:
CSVQF file ["delimiter"] ["header"] ["Columns to use in CSV"]
Example, opening a | delimited file:
CSVQF data.csv "|"

Use the first row as header for listview: (header = 1,Y,Yes,T,True)
CSVQF data.csv "," "1"

Use \t for a tab delimited file
CSVQF tabdata.txt "\t"

Only use specific columns (1 & 5):
CSVQF data.csv "," "0" "1,5"
)

ParamCount=%0%
Param1=%1%
Param2=%2%
Param3=%3%
Param4=%4%

If (ParamCount = 0)
	{
	 CmdLine=1
	 Gosub, ReadHistory
	 Gui, Destroy
	 Gui, Add, Radio,    x16  y72  w120 h20 vR1 Checked, &Comma
	 Gui, Add, Radio,    x16  y102 w120 h20 vR2, &Tab
	 Gui, Add, Radio,    x16  y132 w120 h20 vR3, Se&micolon
	 Gui, Add, Radio,    x16  y162 w120 h20 vR4, &Space
	 Gui, Add, Radio,    x16  y192 w20  h20 vR5,
	 Gui, Add, Edit,     x36  y190 w100 h25 vOther gOther, Other
	 Gui, Add, Edit,     x156 y35  w260 h20 hwndhMyText vMyText gOnChangeMyText, %MyTextStart%
	 Gui, Add, Text,     x156 y60  w260 h20, Recent files ((Dbl-)Click=Select. DEL=Remove)
	 Gui, Add, Listbox,  x156 y78  w260 h120 gHistorySelect vHistorySelect, %HistoryFiles%
	 Gui, Add, GroupBox, x6   y52  w140 h172, Delimiter
	 Gui, Add, Edit,     x6   y6   w300 h25 vFile, Select file
	 Gui, Add, Button,   x316 y6   w100 h25 gBrowse, Br&owse
	 Gui, Add, Checkbox, x6   y35  w150 h20 vFirstRow, Use first row as heade&r
	 Gui, Add, Button,   x156 y198 w40  h25 gHelpStart, ?
	 Gui, Add, Button,   x206 y198 w100 h25 gExit, E&xit
	 Gui, Add, Button,   x316 y198 w100 h25 gContinue, O&K
	 Gui, Show,          x171 y135 w430 h234, %AppWindow%
	 Sleep 100
	 ControlClick, ListBox1, %AppWindow%
	 ControlFocus, Button11, %AppWindow%	
	 Return
	}
Else If (ParamCount = 1)
	{
	 File:=Param1
	 Delimiter:=","
	}
Else If (ParamCount = 2)
	{
	 File:=Param1
	 Delimiter:=Param2
	 If (Param2 = "\t") or (Param2 = "/t")
	 	Delimiter:=A_Tab
	 If (Param2 = "\s") or (Param2 = "/s") or (Param2 = "")
	 	Delimiter:=A_Space
	}
Else If (ParamCount = 3)
	{
	 File:=Param1
	 Delimiter:=Param2
	 If (Param2 = "\t") or (Param2 = "/t")
		Delimiter:=A_Tab
	 If (Param2 = "\s") or (Param2 = "/s") or (Param2 = "")
		Delimiter:=A_Space
	 If Param3 in 1,Y,Yes,T,True,y,yes,t,true,YES,TRUE
		FirstRow:=1
	}
Else If (ParamCount = 3)
	{
	 File:=Param1
	 Delimiter:=Param2
	 If (Param2 = "\t") or (Param2 = "/t")
		Delimiter:=A_Tab
	 If (Param2 = "\s") or (Param2 = "/s") or (Param2 = "")
		Delimiter:=A_Space
	 If Param3 in 1,Y,Yes,T,True
		FirstRow:=1
	}
Else If (ParamCount = 4)
	{
	 File:=Param1
	 Delimiter:=Param2
	 If (Param2 = "\t") or (Param2 = "/t")
		Delimiter:=A_Tab
	 If (Param2 = "\s") or (Param2 = "/s") or (Param2 = "")
		Delimiter:=A_Space
	 If Param3 in 1,Y,Yes,T,True
		FirstRow:=1
	 ColOrder:=Param4
	} 

; load file + build gui

GuiStart:

Header=
FileRead, FullData, %file%

DataIdentifier:=A_Now ; make sure we have a unique CSV Identifier (see CSV lib for details)
Gui, Destroy          ; not pretty but works ;-)
Gui, +Resize
AppWindow=[%file%] - CSV Quick filter - %version%
GroupAdd, AppTitle, %AppWindow%           ; for context sensitive hotkeys (up/down etc)
TrayTip, Loading, Loading %file%..., 3, 1 ; loading can be slow so alert user we are doing something

CSV_Load(file, DataIdentifier, Delimiter)

If (ColOrder <> "")
	{
	 If (FirstRow <> 1)
		{
		 Loop, parse, ColOrder, CSV
		 	 Header .= "c" A_LoopField "|"
		}
	 Loop, % CSV_TotalRows(DataIdentifier) ; prepare search rows for better performance
		{
		 Row:=A_Index
		 Loop, parse, ColOrder, CSV
			{
			 RowDataIdentifier%Row% .= CSV_ReadCell(DataIdentifier, Row, A_LoopField) ","
			}
		}
	} Else If (ColOrder = "") {
	 Loop, % CSV_TotalCols(DataIdentifier) ; prepare a header
		{
		 If (FirstRow <> 1)
			Header .= "c" A_Index "|"
		 ColOrder .= A_Index "," ; will save some repetitive code later on
		}
	 StringTrimRight, ColOrder, ColOrder, 1
	 Loop, % CSV_TotalRows(DataIdentifier) ; prepare search rows for better performance
		RowDataIdentifier%A_Index% := CSV_ReadRow(DataIdentifier, A_Index)
	}

If (FirstRow = 1)
	{
	 Loop, parse, ColOrder, CSV
		Header .= CSV_ReadCell(DataIdentifier, 1, A_LoopField) " (" A_LoopField ")|"
	}

StringTrimRight, Header, Header, 1

StringReplace,SearchColumn,ColOrder,`,,|,All
StringReplace,SearchColumn,SearchColumn,%A_Space%,,All
SearchColumn:= "All||Last|" SearchColumn  

If (SearchColumn = "All||Last|")
	Loop, parse, Header, |
		SearchColumn .= A_Index "|"

; MsgBox % FirstRow ":" ColOrder ":" Header ; for debug only

Gui, Add, Text,x5 gFilter, &Filter:
; Gui, Add, Text,xp+30 yp w10 vIndicator, ;/ ; no longer use this as it slooooooooooows it down way too much
Gui, Add, Edit,xp+30 yp-3 w200 h20 vCurrText gGetText,
Gui, Add, Text,xp+205 yp+3, in col
Gui, Add, DropDownList, 0x8000 xp+30 yp-3 vCol w100 r10, %SearchColumn%
Gui, Add, Checkbox, 0x8000 gSetSearchMethod vRegEx xp+120 yp+3 w130, &RegEx (case sensitive)
Gui, Add, Checkbox, 0x8000 vMute gMute checked%mute% xp+140 yp w50, &Mute
Gui, Add, Button, 0x8000 xp+50 yp-3 w55 h20, &Open
Gui, Add, Button, 0x8000 xp+60 yp  w55 h20, &Export
Gui, Add, Button, 0x8000 xp+60 yp  w55 h20, &Close
Gui, Add, Button, 0x8000 xp+60 yp  gHelp w20 h20, ?
Width:=A_ScreenWidth-20
Height:=A_ScreenHeight-120
Gui, Add, Listview, HWNDhe1 x5 y30 w%Width% h%Height% grid,%Header%
FillFirstTime=1
Gosub, FillListView
VisibleRows:=Ceil(Height/20)        ; for pagedown/pageup
Gui, Add, StatusBar,,               ; statusbar
SB1:=Round(.4*Width)
SB3:=Round(.2*Width)
SB_SetParts(SB1,SB1,SB3)
SB_SetText(file,1)
If (ColOrder <> "")	                ; maybefixlater
	SB_SetText("Rows: " CSV_TotalRows(DataIdentifier) - FirstRow ", Cols: " CSV_TotalCols(DataIdentifier) ", Show: " ColOrder,2)
Else If (ColOrder = "")
	SB_SetText("Rows: " CSV_TotalRows(DataIdentifier) - FirstRow ", Cols: " CSV_TotalCols(DataIdentifier),2)
SB_SetText("Hits: NA",3)
TrayTip                             ; file is now loaded so we can turn it off in case it was a fast loading file
Gosub SetAttach                     ; comment this line to disable Attach
Gui, Show, autosize center Maximize, %AppWindow%
Gosub, GetText
Return

SetAttach:
	Attach(he1, "w1 h1")
Return

Filter:
ControlFocus, Edit1, %AppWindow%
Return

GetText:
StartTime := A_TickCount
ControlGetText, CurrText, Edit1, %AppWindow%
If (CurrText = LastText)
	Return
CurrLen:=StrLen(CurrText)
If (CurrLen = 0)
	{
	 Gosub, FillListView
	 SB_SetText("Hits: NA",3)
	 LastText=
	 Return
	}
Gui, 1:Default
LastText := CurrText
ControlGetText, Col, ComboBox1, %AppWindow%
GuiControl, 1:-Redraw, MyListView
Gosub,ClearList
counter=
match=
HitList=

Loop, % CSV_TotalRows(DataIdentifier)
	{
	 If (A_TickCount - StartTime > 250)
		ControlGetText, CurrText, Edit1, %AppWindow%
	 If (CurrText <> LastText)
		 Goto GetText

	 If ((FirstRow = 1) and (A_Index = 1))
		{
		 Continue ; skip first row, header
		}

;   If (SearchFile() = 0) and (InStr(CurrText,"|") <> 0)
;    {
;     LV_Delete()
;     SB_SetText("Hits: 0", 3)
;   	 Break
;   	} 

	 Row:=A_Index

	 If (Col = "All")
		SearchIn:=RowDataIdentifier%Row%
	 Else If (Col = "Last")
		{
		 Loop, % CSV_TotalCols(DataIdentifier)
			{
			 LastCell:=CSV_TotalCols(DataIdentifier)+1-A_Index
			 SearchIn:=CSV_ReadCell(DataIdentifier, Row, LastCell)
			 If (SearchIn <> "") ; we found the last non empty cell
				Break
			}
		}
	 Else
		SearchIn:=CSV_ReadCell(DataIdentifier, Row, Col)
	 ; MsgBox % SearchIn	; debug only

	 If (InStr(CurrText,"|") <> 0 and (SearchMethod <>1)) ; TODO check for RE
		{
		 SearchMethod=2
		}	
	 Else	
		SearchMethod=0

	 If (SearchMethod=0)
		{
		 Found=0
		 StringSplit, SplitWords, CurrText, %A_Space%
		 Loop, %SplitWords0%
			{
			 Search:=SplitWords%A_Index%
			 If (InStr(SearchIn,Search) > 0)
				Found++
			}
		 If (Found = SplitWords0)
			{
			 Gosub, AddRow
			 Found=0
			}
		}
	 Else if (SearchMethod=1) ; use regex
		{
		 If (RegExMatch(SearchIn, CurrText) > 0)
			{
			 Gosub, AddRow
			}
		}
	 Else if (SearchMethod=2) ; use column searches
		{
		 StringSplit, Query, CurrText, %A_Space%
		 Found=0
		 Loop, %Query0%
			{
			 StringSplit, QueryWord, Query%A_Index%, |
			 SearchIn:=CSV_ReadCell(DataIdentifier, Row, QueryWord1)
			 Search:=QueryWord2
			 If (InStr(SearchIn,Search) <> 0)
				Found++
			}
		 If (Found = Query0)
			{
			 Gosub, AddRow
			 Found=0
			}
		}

;	 If (IndicatorC >= 4)
;	 	IndicatorC=0
;	 IndicatorC++
;	 GuiControl, , Indicator, % si%IndicatorC%
	 GuiControl, 1:+Redraw, MyListView
	}
	GuiControl, , Indicator, =
	SB_SetText("Hits: " . Counter . " / done",3)
If (Mute = 0)
	Soundplay, c:\WINDOWS\Media\chimes.wav
Return

AddRow:
Counter++
LV_Add("","")
Loop, Parse, ColOrder, CSV
	{
	 If (A_LoopField = "")
		Continue
	 LV_Modify(Counter, "Col" . A_Index , CSV_ReadCell(DataIdentifier, Row, A_LoopField))
	} 
SB_SetText("Hits: " Counter, 3)
Return

ClearList:
LV_Delete()
Return

FillListView:
Gosub, ClearList
;MsgBox % ColOrder ; for debug only
Loop, % CSV_TotalRows(DataIdentifier)
	{
	 Row:=A_Index
	 LV_Add("","")
;	 MsgBox % RowDataIdentifier%Row% ; for debug only
	 Loop, Parse, ColOrder, CSV
		{
		 If (A_LoopField = "")
			Continue
		 LV_Modify(Row, "Col" . A_Index , CSV_ReadCell(DataIdentifier, Row, A_LoopField))
		}
	}
If (FirstRow = 1)
	LV_Delete(1)
If FillFirstTime=1                ; so you can set colwidths manually later on
	LV_ModifyCol()
FillFirstTime=0
CurrText=
Return


#IfWinActive, ahk_group AppTitle  ; Hotkeys only work in the just created GUI

Up::
PreviousPos:=LV_GetNext()
If (PreviousPos = 0)              ; exception, focus is not on listview this will allow you to jump to last item via UP key
	{
	 ControlSend, SysListview321, {End}, %AppWindow%
	 Return
	}
ControlSend, SysListview321, {Up}, %AppWindow%
ItemsInList:=LV_GetCount()
ChoicePos:=PreviousPos-1
If (ChoicePos <= 1)
	ChoicePos = 1
If (ChoicePos = PreviousPos)
	ControlSend, SysListview321, {End}, %AppWindow%
ControlFocus, Edit1, %AppWindow%
Return

Down::
PreviousPos:=LV_GetNext()
ControlSend, SysListview321, {Down}, %AppWindow%
ItemsInList:=LV_GetCount()
ChoicePos:=PreviousPos+1
If (ChoicePos > ItemsInList)
	ChoicePos := ItemsInList
If (ChoicePos = PreviousPos)
	ControlSend, SysListview321, {Home}, %AppWindow%
ControlFocus, Edit1, %AppWindow%
Return

Pgdn::
ControlGetFocus, CurrCtrl, %AppWindow%
IfEqual, CurrCtrl, Edit1
	ControlSend, SysListView321, {Down %VisibleRows%}, %AppWindow%
Return

Pgup::
ControlGetFocus, CurrCtrl, %AppWindow%
IfEqual, CurrCtrl, Edit1
	ControlSend, SysListView321, {Up %VisibleRows%}, %AppWindow%
Return

^c::
Copy=1
Enter::
FullRow=
ControlGetFocus, CurrCtrl, %AppWindow%
IfEqual, CurrCtrl, Edit1
	{
	 Gui, Submit, NoHide
	 SelItem := LV_GetNext()
	 If (SelItem = 0)
		SelItem = 1
	 Loop, Parse, ColOrder, CSV
		{
		 LV_GetText(CellData, SelItem, A_Index)
		 If (Copy = 0)
			FullRow .= "Column " A_LoopField ":" A_Tab Format4CSV(CellData) "`n"
		 Else If (Copy = 1)
			FullRow .= A_Tab Format4CSV(CellData) Delimiter
		}
	 If (Copy = 1)
		{
		 Clipboard:=FullRow
		 MsgBox,64,Row copied, Row copied to clipboard
		} Else If (Copy = 0) {
			MsgBox,64,Row view, % FullRow
		}
	}
Copy=0
Return

#IfWinActive

Mute:
Gui, Submit, Nohide
ControlFocus, Edit1, %AppWindow%
Return

SetSearchMethod:
Gui, Submit, Nohide
If (RegEx = 0) or (RegEx = "")
	SearchMethod=0
Else
	SearchMethod=1 ; use RegEx
ControlFocus, Edit1, %AppWindow%
Lasttext=lajflasjflasjfljalfalsfx123
Gosub, GetText
Return

ButtonOpen:
Gui, destroy
Reload
Return


Browse:
FileSelectFile, File, 1, %A_ScriptDir%, Select CSV File, *.csv
if ErrorLevel
	Return
Else
	GuiControl, , Edit3, %File%
Return

HelpStart:
helptitle:=AppWindow
HelpText=
(join`r`n
Select a file and set the delimiter and press OK

If you want to change the columns that will be visible in
the listview and/or the order, enter a comma separated list:
1,4,8 will only show columns 1 4 and 8 in your listview
3,2,1 will show the columns in that order
The export function will use this order as well.
If you use columns that don't exist they show up empty.

%helpforum%
)
MsgBox, 32, %helptitle%, %HelpText%
Return

Other:
GuiControl, , Button5, 1
Return

Continue:
CmdLine=0
Gui, 1:Submit, NoHide

If (R1 = 1)
	Delimiter:=","
If (R2 = 1)
	Delimiter:=A_Tab
If (R3 = 1)
	Delimiter:=";"
If (R4 = 1)
	Delimiter:=A_Space
If (R5 = 1)
	Delimiter:=Other


; MsgBox % "|" Delimiter "|" File ; for debug only

IfNotExist, %File%
	{
	 MsgBox,48, %AppWindow%, Please select a file
	 Return
	}

If (MyText = MyTextStart)
	MyText=
ColOrder:=MyText
Gui, 1: Destroy
;MsgBox % File ":" Delimiter ":" ColOrder ; for debug only
InHistory=0
n:=Chr(5)
AddToHistory:=File n Delimiter n ColOrder n FirstRow
Files.Insert(AddToHistory)
For v in Files
	{
	 If (InHistory=1)
		Continue
	 If (AddToHistory = Files[A_Index]) and (InHistory=0)
		{
		 InHistory=1
		 SaveHistory .= "`n" Files[A_Index]
		 Continue
		}
	 SaveHistory .= "`n" Files[A_Index]
	}

FileDelete,	%A_ScriptDir%\CSVQF.history
FileAppend, %SaveHistory%, %A_ScriptDir%\CSVQF.history
Gosub, GuiStart
Return

ButtonExport:
; CSV_LVSave(FileName, CSV_Identifier, Delimiter, OverWrite?, Gui) ; syntax
FileSelectFile, ExportFile, , %A_ScriptDir%, Save file as, *.csv
if ErrorLevel
	{
	 Return
	}
IfNotInString, ExportFile, .csv
	ExportFile .= ".csv"
CSV_LVSave(ExportFile, DataIdentifier, Delimiter)
MsgBox, Data exported as %ExportFile%
Return

Help:
helptitle:=SubStr(AppWindow,InStr(AppWindow,"] -")+3)
HelpText=
(join`r`n
This program allows you to load a CSV file (any delimited file)
and use various search criteria to filter the listview.
You can export the results to a new file.
The regular expression search is case sensitive and should be a
Perl-compatible regular expression (PCRE, www.pcre.org)
Note: an entire row of the CSV is searched at once and not
on a cell by cell basis to provide faster search results.

%CmdlineOpt%

Press enter: Show row data in a message box
Press ctrl-c: Copy row data to clipboard

%helpforum%
)
MsgBox, 32, %helptitle%, %HelpText%
Return

;Esc::
Exit:
ButtonClose:
GuiClose:
;GuiEscape:

ExitApp
Return

; code by jsherk http://www.autohotkey.com/forum/viewtopic.php?t=73246
OnChangeMyText:
    Gui, Submit, NoHide ; Get the info entered in the GUI
    NewText := RegExReplace(MyText, "[^0-9,]", "") ; Allow digits and comma only
    If NewText != %MyText% ; Check if any invalid characters were removed
    {
        ControlGet, cursorPos, CurrentCol,, %MyText%, A ; Get current cursor position
        GuiControl, Text, MyText, %NewText% ; Write text back to Edit control
        cursorPos := cursorPos - 2
        SendMessage, 0xB1, cursorPos, cursorPos,, ahk_id %hMyText% ; EM_SETSEL ; Add hwndh to Edit control for this to work
    }
    GuiControl, Text, TypedText, %NewText% ; Write text back to the "You Typed:" Text control
Return

HistorySelect:
If A_GuiEvent not in DoubleClick,Normal
	Return
Gui, 1: Submit, NoHide
Loop, parse, History, `n, `r
	{
	 If (InStr(A_Loopfield, HistorySelect) > 0)
		{
		 Gui, Default
		 StringSplit, UpdateField, A_Loopfield, % Chr(5)
		 GuiControl, 1:, Edit3, %UpdateField1%   ; set file
		 GuiControl, 1:, Button5, 1              ; radio of other type...
		 GuiControl, 1:, Edit1, %UpdateField2%   ; other type of separator
		 GuiControl, 1:, Edit2, %UpdateField3%   ; column order
		 GuiControl, 1:, Button8, %UpdateField4% ; user first row as header

		 If (UpdateField2 = ",")
			{
			 ;GuiControl, 1:, Button1, 1
			 GuiControl, 1:, Edit1, Other ; don't ask
			 ControlSend, Button1, !c, %AppWindow%
		 	}
		 Else If (UpdateField2 = A_Tab)
			{
			 ;GuiControl, 1:, Button2, 1
			 ControlSend, Button2, !t, %AppWindow%
			 GuiControl, 1:, Edit1, Other
			}
		 Else If (UpdateField2 = ";")
			{
			 ;GuiControl, 1:, Button3, 1
			 ControlSend, Button3, !m, %AppWindow%
			 GuiControl, 1:, Edit1, Other
			}
		 Else If (UpdateField2 = A_Space)
			{
			 ;GuiControl, 1:, Button4, 1
			 ControlSend, Button4, !s, %AppWindow%
			 GuiControl, 1:, Edit1, Other
			}
		}
	}
Return

#IfWinActive ahk_group AppTitle
$Del::
ControlGetFocus, ListboxControl, %AppWindow%
If (CmdLine = 0)
	{
	 Send {Del}
	 Return
	}
If (ListboxControl <> "ListBox1")
	{
	 Send {Del}
	 Return
	}
SendMessage, 0x188, 0, 0, ListBox1, %GUITitle%  ; 0x188 is LB_GETCURSEL (for a ListBox).
Pos:=ErrorLevel+1
HistoryFiles=
Files.Remove(Pos)
For v in Files
	HistoryFiles .= SubStr(Files[A_Index],1,InStr(Files[A_Index],Chr(5))-1) "|"
GuiControl, , ListBox1, |%HistoryFiles%
; MsgBox % HistoryFiles
Return
#IfWinActive

ReadHistory:
Files := Array()
FileRead, History, %A_ScriptDir%\CSVQF.history
;File n Delimiter n ColOrder n FirstRow
Loop, parse, History, `n, `r
	{
	 If (A_LoopField = "")
		Continue ; skip empty lines
	 HistoryFiles .= SubStr(A_LoopField,1,InStr(A_LoopField,Chr(5))-1) "|"
	 Files.Insert(A_LoopField)
	}
Return

SearchFile()
	{
	 Global
	 StringReplace, QuickCurrText, CurrText, %A_Space%, .*, All
	 Result:=RegExMatch(FullData, "ismU)" . QuickCurrText)
	 Return Result
	}
