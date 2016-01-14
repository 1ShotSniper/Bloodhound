#NoEnv
; #Warn
#singleinstance force
;#WinActivateForce
#include <CSV_Lib>
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
{	
	msgbox Error!
	Exitapp 1
}
loop
{
	Row:=A_Index
	Status:= CSV_ReadCell("sheet",Row,5)
	if Status is not space
		continue
	CSV_Save(filename,"sheet","1")
	full:=
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
	StringReplace,Z,name,%A_SPACE%,`.`,,1
	loop
	{
		if A_Index>=20
		{
			CSV_ModifyCell("sheet","SKIP",Row,5)
			continue 2
		}		
		send {F6}
		sleep 50
		clipboard :="javascript"
		clipwait
		send ^v
		sleep 50
		clipboard:=":surnameandnamelookup('" . Surname . "', '" . Z . ".', '0', '" . A_Index-1 . "')"
		clipwait
		send ^v
		sleep 50
		send {ENTER}
		sleep 5000
		loop
		{
			clipboard:=
			send ^a^c^+{HOME}
			clipwait
			if not CountSubstring(clipboard,surname)>0 and not instr(clipboard,"No records found")
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
	}
	FilterNames(full,row,name)
	numtraced:=0
	
	if idList0=1
	{
		CSV_ModifyCell("sheet",idList1,Row,5)
	}
	else
	{
		loop %idList0%
		{
			if numtraced>=10
				break
			idnum:=idList%A_Index%			
			previous:=A_Index-1
			if idnum = idList%previous%
					continue
			send {F6}
			sleep 50
			clipboard :="javascript"
			clipwait
			send ^v
			sleep 50
			clipboard:=":idlookup('" . idnum . "')"
			clipwait
			send ^v
			sleep 50
			send {ENTER}{ESC}
			sleep 5000
			loop
			{
				if mod(%A_Index%,10) = 0
				{
					send {F6}
					sleep 50
					clipboard :="javascript"
					clipwait
					send ^v
					sleep 50
					clipboard:=":idlookup('" . idnum . "')"
					clipwait
					send ^v
					sleep 50
					send {ENTER}{ESC}
				}
				clipboard:=	
				send ^a^c
				clipwait
				send ^+{HOME}
				if not CountSubstring(clipboard,idnum)>0
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
				A:=CountSubstring(SubStr(clipboard,InStr(clipboard,"PROPERTIES"),InStr(clipboard,"TELEPHONE NUMBERS")-InStr(clipboard,"PROPERTIES")-2),"`r`n")-3
				if  not InStr(SubStr(clipboard,InStr(clipboard,"PROPERTIES"),InStr(clipboard,"TELEPHONE NUMBERS")-InStr(clipboard,"PROPERTIES")-2),"No records found") and InStr(SubStr(clipboard,InStr(clipboard,"PROPERTIES"),InStr(clipboard,"TELEPHONE NUMBERS")-InStr(clipboard,"PROPERTIES")-2),Registrar) and Registrar is not space
					loop %A%
					{
						send {PGDN}
						sleep 50
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
						if clipboard contains %LRegion%
						{
							CSV_ModifyCell("sheet",idnum,row,currentcell)
							currentcell:=currentcell+1
							numtraced:=numtraced+1
						}
						else if A_Index!=A
						{
							send {F6}
							sleep 50
							clipboard :="javascript"
							clipwait
							send ^v
							sleep 50
							clipboard:=":idlookup('" . idnum . "')"
							clipwait
							send ^v
							sleep 50
							send {ENTER}{ESC}
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
FilterNames(full,row,name){
	global
	idList0:=0
	StringSplit details,full,`r,`n
	if details0=1
	{
		StringSplit gridd,details1,%A_Tab%
		CSV_ModifyCell("sheet",gridd5,Row,5)
		return
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
			tempseq:=sequence%A_Index%
			if bureauName%A_Index% not contains %tempseq%
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
}
F1::
pause
return
