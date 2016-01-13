; includes bugfixes from:
; - http://www.autohotkey.com/forum/viewtopic.php?p=400669#400669
; - http://www.autohotkey.com/forum/viewtopic.php?p=453352#453352
; AutoHotkey_L Version: 1.1.08.01
; Author: Kdoske, trueski, SoLong&Thx4AllTheFish, segalion
; http://www.autohotkey.com/forum/viewtopic.php?p=329126#329126
;##################################################    CSV    FUNCTIONS     

;####################################################################################################################
;CSV Functions
;####################################################################################################################
CSV_Load(FileName, CSV_Identifier="", Delimiter="`,")
{
  Local Row
  Local Col
  temp :=  %CSV_Identifier%CSVFile
  FileRead, temp, %FileName%
  String = a,b,c`r`nd,e,f,,"g`r`n",h`r`nB,`n"C`nC",D
  e:= """" ; the encapsulation character (tipical ")
  RegExNeedle:= "\n(?=[^" e "]*" e "([^" e "]*" e "[^" e "]*" e ")*([^" e "]*)\z)"
  String := RegExReplace(String, RegExNeedle , "" )
  StringReplace,String, String,`r,@,All ;only for see msgbox 
  Loop, Parse, temp, `n, `r
  {
    Col := ReturnDSVArray(A_LoopField, CSV_Identifier . "CSV_Row" . A_Index . "_Col", Delimiter)
    Row := A_Index
  }
  %CSV_Identifier%CSV_TotalRows := Row
  %CSV_Identifier%CSV_TotalCols := Col
  %CSV_Identifier%CSV_Delimiter := Delimiter
  SplitPath, FileName, %CSV_Identifier%CSV_FileName, %CSV_Identifier%CSV_Path
  IfNotInString, FileName, `\
  {
    %CSV_Identifier%CSV_FileName := FileName
    %CSV_Identifier%CSV_Path := A_ScriptDir
  }
  %CSV_Identifier%CSV_FileNamePath := %CSV_Identifier%CSV_Path . "\" . %CSV_Identifier%CSV_FileName
}
;####################################################################################################################
CSV_Save(FileName, CSV_Identifier, OverWrite="1")
{
Local Row
Local Col

If OverWrite = 0
 IfExist, %FileName%
   Return
 
FileDelete, %FileName%

EntireFile =
CurrentCSV_TotalRows := %CSV_Identifier%CSV_TotalRows
CurrentCSV_TotalCols := %CSV_Identifier%CSV_TotalCols
Loop, %currentcsv_totalrows%
{
    Row := A_Index
   Loop, %currentCSV_TotalCols%
   {
      Col := A_Index
      EntireFile .= Format4CSV(%CSV_Identifier%CSV_Row%Row%_Col%Col%)
      If (Col <> %CSV_Identifier%CSV_TotalCols)
         EntireFile .= %CSV_Identifier%CSV_Delimiter
   }
   If (Row < %CSV_Identifier%CSV_TotalRows)
      EntireFile .= "`n"
} 
   StringReplace, temp, temp, `r`n`r`n, `r`n, all   ;Remove all blank lines from the CSV file
   loop,
   {
      stringright, test, EntireFile, EntireFile, 1
      if (test = "`n") or (test = "`r")
         stringtrimright, EntireFile, EntireFile, 1
      Else
         break
   }
    FileAppend, %EntireFile%, %FileName%
}
;####################################################################################################################
CSV_TotalRows(CSV_Identifier)
{
  global   
  CurrentCSV_TotalRows := %CSV_Identifier%CSV_TotalRows
  Return %CurrentCSV_TotalRows%
}
;####################################################################################################################
CSV_TotalCols(CSV_Identifier)
{
  global   
  CurrentCSV_TotalCols := %CSV_Identifier%CSV_TotalCols
  Return %CurrentCSV_TotalCols%
}
;####################################################################################################################
CSV_Delimiter(CSV_Identifier)
{
  global
  CurrentCSV_Delimiter := %CSV_Identifier%CSV_Delimiter
  Return %CurrentCSV_Delimiter%
}
;####################################################################################################################
CSV_FileName(CSV_Identifier)
{
  global   
  CurrentCSV_FileName := %CSV_Identifier%CSV_FileName
  Return %CurrentCSV_FileName%
}
;####################################################################################################################
CSV_Path(CSV_Identifier)
{
  global
  CurrentCSV_Path := %CSV_Identifier%CSV_Path
  Return %CurrentCSV_Path%
}
;####################################################################################################################
CSV_FileNamePath(CSV_Identifier)
{
  global
  CurrentCSV_FileNamePath := %CSV_Identifier%CSV_FileNamePath
  Return %CurrentCSV_FileNamePath%
}
;####################################################################################################################
CSV_ModifyCell(CSV_Identifier, Value, Row, Col)
  {
   global
    %CSV_Identifier%CSV_Row%Row%_Col%Col% := Value
  }
;####################################################################################################################
CSV_ReadCell(CSV_Identifier, Row, Col)
{
  Local CellData
  CellData := %CSV_Identifier%CSV_Row%Row%_Col%Col%
  Return %CellData%
}
;####################################################################################################################
Format4CSV(F4C_String)
{
   Reformat:=False ;Assume String is OK
   IfInString, F4C_String,`n ;Check for linefeeds
      Reformat:=True ;String must be bracketed by double quotes
   IfInString, F4C_String,`r ;Check for linefeeds
      Reformat:=True
   IfInString, F4C_String,`, ;Check for commas
      Reformat:=True
   IfInString, F4C_String, `" ;Check for double quotes
   {   Reformat:=True
      StringReplace, F4C_String, F4C_String, `",`"`", All ;The original double quotes need to be double double quotes
   }
   If (Reformat)
      F4C_String=`"%F4C_String%`" ;If needed, bracket the string in double quotes
   Return, F4C_String
}
;#################################################################################################################### 
