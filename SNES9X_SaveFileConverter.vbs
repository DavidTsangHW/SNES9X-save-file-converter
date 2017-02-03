'1 FEB 2017
'SNES9X SAVE FILE CONVERTER
'CONVERT SNES9X SAVE STATES BETWEEN ANDROID AND WINDOWS

'Copyright (C) 2017  DAVID TSANG
'exactlytheapp@gmail.com
'https://play.google.com/store/apps/developer?id=CodeMonkeyA&hl=en

'***********************IMPORTANT******************************

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

'USAGE
'	RULE OF THUMB ALWAYS BACKUP SAVE FILES 
'
'1. 	COPY ALL WINDOWS AND ANDROID SAVE FILES TO A FOLDER. EXECUTE THIS PROGRAM IN THE FOLDER WITH SAVE FILES.
'
'2. 	IT WILL CONVERT SAVE FILES INTO ANDROID/WINDOWS FORMAT BASED ON FILE TIMESTAMP 
'	IF CORRESPONDING ANDROID/WINDOWS SAVE FILE DOES NOT EXISTS, THE PROGRAM WILL CREATE IT.
'	IF BOTH ANDOIRD AND WINDOWS SAVE FILES EXIST, THE PROGRAM WILL COMPARE THEIR TIMESTAMPS AND OVERWRITES THE OLDER ONE.
'	A LOG FILE IS CREATED EACH TIME AFTER EXECUTION.
'
'3.	COPY THE SAVE FILES BACK TO SNES9X SAVE FOLDER.

'CONVERSION TABLE
'--------------------------
'ANDROID	WINDOWS
'.0A.frz 	.000
'.0B.frz 	.001
'.0C.frz	.002
'.0D.frz	.003
'.0E.frz	.004
'.0F.frz	.005
'.0G.frz	.006
'.0H.frz	.007
'.0I.frz	.008
'.0J.frz	.009

Dim CurrentDirectory
Dim fs
Dim line
Dim quickIndex
Dim count

redim PCExt(9)
redim AndExt(9)

for i = 0 to 9
	PCExt(i) = "00" & i
	AndExt(i) = "0" & ucase(chr(65+i)) & ".frz"
	quickIndex = quickIndex & PCExt(i) & " 0" & ucase(chr(65+i))
next

set fs=CreateObject("Scripting.FileSystemObject")
CurrentDirectory = fs.GetAbsolutePathName(".")

Set MyFiles = fs.GetFolder(CurrentDirectory)  

count = 0

For Each MyFile In MyFiles.Files

	GetAnExtension = fs.GetExtensionName(MyFile.name)
	GetAnExtension = GetANExtension
	Fname = MyFile.name
	
	if GetAnExtension = "frz" then
		Fname = replace(Fname, "." & GetAnExtension, "", 1, -1, 1)
	end if

	s = Split(ucase(Fname), ".")
	ext = s(Ubound(s))

	'Validate file extension
	if instr(quickIndex, ext) > 0 then

		index = right(ext,1)
		filename = replace(Fname, ext, "", 1, -1, 1)

		if not isnumeric(index) and len(ext) = 2 then  '0A, 0B, 0C...
			index = asc(index) - 65	
			target = PCExt(index)		
		else					       '001, 002, 003...
			target = AndExt(index) 
		end if	

		target = filename & target

		call syncFile(MyFile, target)
		
	end if

Next

call writeFile("log.txt", line)

	msgbox count & " file(s) converted." & vbcrlf & vbcrlf & "Read log.txt for details", 64,"SNES9X - Android/Windows save files converter"

sub syncFile(MyFile, target)

	dim flag

	flag = -1 

	if fs.FileExists(target) then
		
		set t = fs.GetFile(CurrentDirectory & "\" & target)

		if MyFile.DateLastModified > t.DateLastModified then
		
			flag = 1		
		end if

	else
		flag = 1

	end if

	if flag = 1 then

		call fs.copyfile(MyFile.name,target)
		count = count + 1
		line = line & "Copying " & MyFile.name & vbtab & MyFile.datelastModified & vbtab & target & vbcrlf

	end if


end sub

sub writeFile(Filename,Lines)
	
	set outputFile=fs.CreateTextFile(Filename,1)
	outputFile.writeLine(Lines)	
	outputFile.Close

end sub