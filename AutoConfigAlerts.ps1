$data = Import-Csv -Path .\test_info.csv
$department = ""

while ($true) {
	$department = Read-Host "Enter the Department"
	Write-Host "Department entered was: $department"
	$selection = Read-Host "Confirm [y] [n]?"

	if ($selection -eq "y") {
		$department = $department.ToUpper()
		try {
			mkdir $department -ErrorAction Stop
		}
		catch {
			Remove-Item $department -Recurse -Force
			mkdir $department
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

function CreateShortcuts ([string]$StartPath, [string]$ShortcutName, [string]$IconFilePath, [string]$Department, [string]$AlertType) {
	$target = "-File `"$AlertType`""
	$WshShell = New-Object -comObject WScript.Shell
	$Shortcut = $WshShell.CreateShortcut("$StartPath\$ShortcutName")
	$Shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
	$Shortcut.Arguments = $target
	$Shortcut.IconLocation = $IconFilePath
	$Shortcut.WorkingDirectory = $StartPath
	$Shortcut.Save()
}

function CreatePS1Files ([string]$MakeFile, [string]$WriteLine1, [string]$WriteLine2) {
	$WriteLine1 | Out-File -FilePath $MakeFile -Append
	$WriteLine2 | Out-File -FilePath $MakeFile -Append
}

foreach ($entry in $data) {
	$myWD = $pwd.ToString()
	$webhook = $entry.Webhook
	$room = $entry.Room
	$name = $entry.Name
    $message = "$name in room $room"
	$helpTitle = $entry.Help_Title
	$medTitle = $entry.Medical_Title
	$threatTitle = $entry.Threat_Title
	$startInPath = "$myWD\$department\$room"
	$lowDept = $department.ToLower()

	mkdir "$department\$room"

	$PS1FileLine1 = "`$myTeamsWebHook = `"$webhook`""
	$PS1FileLine2_Help = "Invoke-RestMethod -Method post -ContentType `'Application/Json`' -Body `'{`"title`": `"$helpTitle`", `"text`":`"$message`"}`' -Uri `$myTeamsWebHook"
	$PS1FileLine2_Medical = "Invoke-RestMethod -Method post -ContentType `'Application/Json`' -Body `'{`"title`": `"$medTitle`",`"text`":`"$message`"}`' -Uri `$myTeamsWebHook"
	$PS1FileLine2_Threat = "Invoke-RestMethod -Method post -ContentType `'Application/Json`' -Body `'{`"title`": `"$threatTitle`",`"text`":`"$message`"}`' -Uri `$myTeamsWebHook"

	$helpPath = "alertingpost_$lowDept`_helpneeded.ps1"	
	$medPath = "alertingpost_$lowDept`_urgent_medical.ps1"
	$threatPath = "alertingpost_$lowDept`_urgent_threat.ps1"

	#Create .ps1 files
	CreatePS1Files -MakeFile "$startInPath\$helpPath" -WriteLine1 $PS1FileLine1 -WriteLine2 $PS1FileLine2_Help
	CreatePS1Files -MakeFile "$startInPath\$medPath" -WriteLine1 $PS1FileLine1 -WriteLine2 $PS1FileLine2_Medical
	CreatePS1Files -MakeFile "$startInPath\$threatPath" -WriteLine1 $PS1FileLine1 -WriteLine2 $PS1FileLine2_Threat

	#Create .lnk shortcuts
	CreateShortcuts -StartPath $startInPath -ShortcutName "Alert - Help Needed.lnk" -IconFilePath "$myWD\Icons\Alert - Yellow.ico" -WorkingDir $myWD -Department $lowDept -AlertType $helpPath
	CreateShortcuts -StartPath $startInPath -ShortcutName "Alert - Medical.lnk" -IconFilePath "$myWD\Icons\Alert - Blue.ico" -WorkingDir $myWD -Department $lowDept -AlertType $medPath
	CreateShortcuts -StartPath $startInPath -ShortcutName "Alert - Threat.lnk" -IconFilePath "$myWD\Icons\Alert - Red.ico" -WorkingDir $myWD -Department $lowDept -AlertType $threatPath
}

