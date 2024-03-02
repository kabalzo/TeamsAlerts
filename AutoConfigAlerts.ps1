$data = Import-Csv -Path .\AlertTemplate.csv
$dept = ""

while ($true) {
	$_dept= Read-Host "Enter the Department"
	Write-Host "Department entered was: $_dept"
	$selection = Read-Host "Confirm [y] [n]?"
	if ($selection -eq "y") {
		$_dept = $_dept.ToUpper()
		$dept = $_dept
		try {
			mkdir $_dept -ErrorAction Stop
		}
		catch {
			Remove-Item $_dept -Recurse -Force
			mkdir $_dept
		}
		break
	}
	elseif ($selection -eq "n") {
		continue 
	}
	else { 
		Write-Host "Invalid selection. Try again." -ForegroundColor Red
	}
}

foreach ($line in $data) {
	$myWD = $pwd.ToString()
	$webhook = $line.Webhook
	$room = $line.Room
	$name = $line.Name
    $message = "$name in room $room"
	$helpTitle = $line.Help_Title
	$medTitle = $line.Medical_Title
	$threatTitle = $line.Threat_Title
	mkdir "$dept\$room"
	$startInPath = "$myWD\$dept\$room"

	$line1 = "`$myTeamsWebHook = `"$webhook`""
	$line2_Help = "Invoke-RestMethod -Method post -ContentType `'Application/Json`' -Body `'{`"title`": `"$helpTitle`", `"text`":`"$message`"}`' -Uri `$myTeamsWebHook"
	$line2_Medical = "Invoke-RestMethod -Method post -ContentType `'Application/Json`' -Body `'{`"title`": `"$medTitle`",`"text`":`"$message`"}`' -Uri `$myTeamsWebHook"
	$line2_Threat = "Invoke-RestMethod -Method post -ContentType `'Application/Json`' -Body `'{`"title`": `"$threatTitle`",`"text`":`"$message`"}`' -Uri `$myTeamsWebHook"

	#Create Help Needed .ps1 file
	$fileToMake = "$startInPath\test1.ps1"
	$line1 | Out-File -FilePath $fileToMake -Append
	$line2_Help | Out-File -FilePath $fileToMake -Append

	#Create Medical .ps1 file
	$fileToMake = "$startInPath\test2.ps1"
	$line1 | Out-File -FilePath $fileToMake -Append
	$line2_Medical | Out-File -FilePath $fileToMake -Append

	#Create Threat .ps1 file
	$fileToMake = "$startInPath\test3.ps1"
	$line1 | Out-File -FilePath $fileToMake -Append
	$line2_Threat | Out-File -FilePath $fileToMake -Append

	#Create "Alert - Help Needed" shortcut
	$target = "-File `"$startInPath\test1.ps1`""
	$WshShell_Help = New-Object -comObject WScript.Shell
	$Shortcut_Help = $WshShell_Help.CreateShortcut("$startInPath\Alert - Help Needed.lnk")
	$Shortcut_Help.TargetPath = """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"""
	$Shortcut_Help.Arguments = $target
	try {
		$Shortcut.IconLocation = "$myWD\Icons\Alert - Yellow.ico"
	}
	catch {
		Write-Host "Unable to set custom icon" -ForegroundColor Red
	}
	$Shortcut_Help.WorkingDirectory = $myWD
	$Shortcut_Help.Save()

	#Create "Alert - Medical" shortcut
	$target = "-File `"$startInPath\test2.ps1`""
	$WshShell_Medical = New-Object -comObject WScript.Shell
	$Shortcut_Medical = $WshShell_Medical.CreateShortcut("$startInPath\Alert - Medical.lnk")
	$Shortcut_Medical.TargetPath = """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"""
	$Shortcut_Medical.Arguments = $target
	try {
		$Shortcut.IconLocation = "$myWD\Icons\Alert - Blue.ico"
	}
	catch {
		Write-Host "Unable to set custom icon" -ForegroundColor Red
	}
	$Shortcut_Medical.WorkingDirectory = $myWD
	$Shortcut_Medical.Save()

	#Create "Alert - Threat" shortcut
	$target = "-File `"$startInPath\test3.ps1`""
	$WshShell_Threat = New-Object -comObject WScript.Shell
	$Shortcut_Threat = $WshShell_Threat.CreateShortcut("$startInPath\Alert - Threat.lnk")
	$Shortcut_Threat.TargetPath = """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"""
	$Shortcut_Threat.Arguments = $target
	try {
		$Shortcut.IconLocation = "$myWD\Icons\Alert - Red.ico"
	}
	catch {
		Write-Host "Unable to set custom icon" -ForegroundColor Red
	}
	$Shortcut_Threat.WorkingDirectory = $myWD
	$Shortcut_Threat.Save()
}

