$data = Import-Csv -Path .\test_info.csv

foreach ($line in $data) {
	$myWD = $pwd.ToString()
	$webhook = $line.Webhook
	$room = $line.Room
	$name = $line.Name
    $message = "$name in room $room"
	$helpTitle = $line.Help_Title
	$medTitle = $line.Medical_Title
	$threatTitle = $line.Threat_Title
	mkdir $room

	$line1 = "`$myTeamsWebHook = `"$webhook`""
	$line2_Help = "Invoke-RestMethod -Method post -ContentType `'Application/Json`' -Body `'{`"title`": `"$helpTitle`", `"text`":`"$message`"}`' -Uri `$myTeamsWebHook"
	$line2_Medical = "Invoke-RestMethod -Method post -ContentType `'Application/Json`' -Body `'{`"title`": `"$medTitle`",`"text`":`"$message`"}`' -Uri `$myTeamsWebHook"
	$line2_Threat = "Invoke-RestMethod -Method post -ContentType `'Application/Json`' -Body `'{`"title`": `"$threatTitle`",`"text`":`"$message`"}`' -Uri `$myTeamsWebHook"

	#Create Help Needed .ps1 file
	$fileToMake = "$myWD\$room\test1.ps1"
	$line1 | Out-File -FilePath $fileToMake -Append
	$line2_Help | Out-File -FilePath $fileToMake -Append

	#Create Medical .ps1 file
	$fileToMake = "$myWD\$room\test2.ps1"
	$line1 | Out-File -FilePath $fileToMake -Append
	$line2_Medical | Out-File -FilePath $fileToMake -Append

	#Create Threat .ps1 file
	$fileToMake = "$myWD\$room\test3.ps1"
	$line1 | Out-File -FilePath $fileToMake -Append
	$line2_Threat | Out-File -FilePath $fileToMake -Append

	#Create "Alert - Help Needed" shortcut
	$target = "-File `"$myWD\$room\test1.ps1`""
	$WshShell_Help = New-Object -comObject WScript.Shell
	$Shortcut_Help = $WshShell_Help.CreateShortcut($myWD + "\$room\Alert - Help Needed.lnk")
	$Shortcut_Help.TargetPath = """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"""
	$Shortcut_Help.Arguments = $target
	#$Shortcut.IconLocation = "$myWD\Icons\Alert - Blue.ico"
	$Shortcut_Help.WorkingDirectory = $myWD
	$Shortcut_Help.Save()

	#Create "Alert - Medical" shortcut
	$target = "-File `"$myWD\$room\test2.ps1`""
	$WshShell_Medical = New-Object -comObject WScript.Shell
	$Shortcut_Medical = $WshShell_Medical.CreateShortcut($myWD + "\$room\Alert - Medical.lnk")
	$Shortcut_Medical.TargetPath = """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"""
	$Shortcut_Medical.Arguments = $target
	#$Shortcut.IconLocation = "$myWD\Icons\Alert - Yellow.ico"
	$Shortcut_Medical.WorkingDirectory = $myWD
	$Shortcut_Medical.Save()

	#Create "Alert - Threat" shortcut
	$target = "-File `"$myWD\$room\test1.ps3`""
	$WshShell_Threat = New-Object -comObject WScript.Shell
	$Shortcut_Threat = $WshShell_Threat.CreateShortcut($myWD + "\$room\Alert - Threat.lnk")
	$Shortcut_Threat.TargetPath = """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"""
	$Shortcut_Threat.Arguments = $target
	#$Shortcut.IconLocation = "$myWD\Icons\Alert - Red.ico"
	$Shortcut_Threat.WorkingDirectory = $myWD
	$Shortcut_Threat.Save()
}

