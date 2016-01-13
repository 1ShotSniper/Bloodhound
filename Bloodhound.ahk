#NoEnv
; #Warn
#singleinstance force
;#WinActivateForce
#include <CSVLib>
SendMode Input
SetWorkingDir %A_ScriptDir%
DetectHiddenWindows off
coordmode mouse,client
coordmode pixel,client
FileReadLine Regions,Parameters.dat,1
FileReadLine Registrar,Parameters.dat,2
SetCapslockState AlwaysOff
loop
{
	FileSelectFile filename,3,,,Comma Separated Values (*.csv)
	if filename is not space
		break
}
WinActivate ahk_exe chrome.exe
WinMove ahk_exe chrome.exe,,0,0,1024,768
CSV_LOAD(filename,"sheet","`,")
if CSV_TotalCols("sheet") <2
	msgbox Error!
loop
{
	Row:=A_Index
	Status:= CSV_ReadCell("sheet",Row,5)
	if Status is not space
		continue
	CSV_Save(filename,"sheet","1")
	full:=
	pages:=0
	clipboard:=
	Y:=
	currentcell:=5
	if A_Index > CSV_TotalRows("sheet")
	{
		SetCapslockState Off
		ExitApp 0
	}
	LRegion:=CSV_ReadCell("sheet",Row,4)
	surname:=CSV_ReadCell("sheet",Row,3)
	name:=CSV_ReadCell("sheet",Row,2)
	if LRegion is space
		Lregion:=Regions
	send {ESC 3}{TAB 11}{HOME}{RIGHT 3}
	sleep 500
	send {TAB 2}
	sleep 500
	send {up 2}{down 2}
	sleep 500
	clipboard:=surname
	send {TAB}^v
	sleep 50
	send {TAB}
	StringReplace,Z,name,%A_SPACE%,`.`,
	send %Z%
	; Gonna update this soon
	send {TAB 2}{ENTER}
	loop
	{
		if pages>=20
		{
			CSV_ModifyCell("sheet","SKIP",Row,5)
			break
		}
		sleep 5000
		loop
		{
			clipboard:=
			send ^a^c^+{HOME}
			clipwait
			if not CountSubstring(clipboard,surname)>1 and not instr(clipboard,"No records found")
			{				
					sleep 1000
					continue
			}
			else
				break
		}
		if(InStr(clipboard,"No records found"))
			break
		full:=full . SubStr(clipboard,Instr(clipboard,"First Name")+12,Instr(clipboard,"First Page")-Instr(clipboard,"First Name")-15)
		pages:=pages+1
		sleep 50
		if A_Index = 1
		{
			send {TAB}f
			sleep 100
			send gf
			send f
			sleep 100
			send gd
		}
		else
			SEND ]]
	}
	if pages>=20
	{
		continue
	}
	
	;sort algorithm start

	idList0:=0
	StringSplit details,full,`r,`n
	if details0=1
	{
		StringSplit gridd,details1,%A_Tab%
		CSV_ModifyCell("sheet",gridd5,Row,5)
		continue
	}
	StringSplit sequence,name,%A_Space%
	m:=1
	loop %details0%
	{
		StringSplit gridd,details%A_Index%,%A_Tab%
		StringSplit bureauName,gridd5,%A_Space%
		match:=1
		loop %sequence0%
		{
			if CountSubstring(sequence%A_Index%,bureauName%A_Index%)!=1
			{
				match:=0
				break
			}
		}
		if match=1
		{
			idList%m%:=gridd1
			idList0:=idList0+1
			m:=m+1
		}
	}

	;algorithm end
	
	if idList0!=0
		send {ESC}{TAB 11}{HOME}
	sleep 500
	numtraced:=0
	if idList0=1
		CSV_ModifyCell("sheet",idList1,Row,5)
	else
	{
		loop %idList0%
		{
			if numtraced>=10
				break
			clipboard:=idList%A_Index%
			idnum:=idList%A_Index%
			previous:=A_Index-1
			if not idList%A_Index% != idList%previous%
					continue
			send {ESC}gg
			sleep 100
			send {g down}{i down}
			sleep 100
			send {backspace 13}^v
			sleep 100
			send {i up}{g up}
			sleep 5000
			loop
			{
				if mod(%A_Index%,3) = 0
				{
					clipboard:=idnum
					send {g down}{i down}
					sleep 100
					send {backspace 13}^v
					sleep 100
					send {i up}{g up}
				}
				clipboard:=	
				send ^a^c
				clipwait
				send ^+{HOME}
				if not CountSubstring(clipboard,idnum)>1
				{
					sleep 1000
					continue
				}
				else
					break
			}
			if clipboard contains %LRegion%
			{
				CSV_ModifyCell("sheet",idnum,row,currentcell)
				currentcell:=currentcell+1
				numtraced:=numtraced+1
			}
			else
			{	
				sleep 50
				A:=CountSubstring(SubStr(clipboard,InStr(clipboard,"PROPERTIES"),InStr(clipboard,"TELEPHONE NUMBERS")-InStr(clipboard,"PROPERTIES")-2),"`r`n")-3
				sleep 50
				if  not InStr(SubStr(clipboard,InStr(clipboard,"PROPERTIES"),InStr(clipboard,"TELEPHONE NUMBERS")-InStr(clipboard,"PROPERTIES")-2),"No records found") and InStr(SubStr(clipboard,InStr(clipboard,"PROPERTIES"),InStr(clipboard,"TELEPHONE NUMBERS")-InStr(clipboard,"PROPERTIES")-2),Registrar)
					loop %A%
					{
						sleep 200
						ImageSearch,, Y, 0, 0,1024,768, *50 *trans0x3F474E %A_scriptdir%\Search.bmp
						Sleep 200
						mousemove 650,Y+23+(A_Index * 12),0
						click
						sleep 5000
						loop
						{
							clipboard:=
							send ^a^c^+{HOME}
							clipwait
							if not InStr(Clipboard,idnum)
							{
								sleep 1000
								continue
							}
							else
								break
						}
						sleep 50
						if clipboard contains %LRegion%
						{
							CSV_ModifyCell("sheet",idnum,row,currentcell)
							currentcell:=currentcell+1
							numtraced:=numtraced+1
							break
						}
						else if A_Index!=A
						{
							clipboard:=idnum
							send {g down}{i down}
							sleep 100
							send {BackSpace 13}{i up}{g up}
							send ^v
							sleep 5000
						}
					}
			}	
			send gg
		}
		if numtraced=0
			CSV_ModifyCell("sheet","No Trace",Row,5)
	}
}
CountSubstring(fullstring, substring){
  StringReplace, junk, fullstring, %substring%, , UseErrorLevel
	return errorlevel
}
F1::
pause
return
