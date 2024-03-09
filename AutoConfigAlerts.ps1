$dataPath = ".\AlertTemplate.cvs"
#$dataPath = ".\test_info.csv"
$data = Import-Csv -Path $dataPath
$department = ""
$PS1FileLine2Part1 = "Invoke-RestMethod -Method post -ContentType `'Application/Json`' -Body `'{`"title`":"
$PS1FileLine2Part2 = "}`' -Uri `$myTeamsWebHook"
$myWD = $pwd.ToString()
$iconFiles = @("Yellow.ico", "Blue.ico", "Red.ico")
$shortcutNames = @("Help Needed", "Medical", "Threat")

#Menu
while ($true) {
	$department = Read-Host "Enter the Department"
	Write-Host "Department entered was: $department"
	$selection = Read-Host "Confirm [y] [n]?"

	if ($selection -eq "y") {
		$department = $department.ToUpper()
  		$lowDept = $department.ToLower()
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

function CreateShortcuts ([string]$StartPath, [string]$ShortcutName, [string]$IconFile, [string]$Department, [string]$FileName) {
	$arguments = "-File `"$FileName`""
	$WshShell = New-Object -comObject WScript.Shell
	$Shortcut = $WshShell.CreateShortcut("$StartPath\Alert - $ShortcutName.lnk")
	$Shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
	$Shortcut.Arguments = $arguments
	$Shortcut.IconLocation = "$myWD\Icons\Alert - $IconFile"
	$Shortcut.WorkingDirectory = $StartPath
	$Shortcut.Save()
}

function CreatePS1Files ([string]$MakeFile, [string]$WriteLine1, [string]$WriteLine2) {
	$WriteLine1 | Out-File -FilePath $MakeFile -Append
	$WriteLine2 | Out-File -FilePath $MakeFile -Append
}

$ps1FileNames = @("alertingpost_$lowDept`_helpneeded.ps1", "alertingpost_$lowDept`_urgent_medical.ps1", "alertingpost_$lowDept`_urgent_threat.ps1")
foreach ($entry in $data) {
	$webhook = $entry.Webhook
	$PS1FileLine1 = "`$myTeamsWebHook = `"$webhook`""
	$room = $entry.Room
	$name = $entry.Name
    $message = "$name in room $room"
	$helpTitle = $entry.Help_Title
	$medTitle = $entry.Medical_Title
	$threatTitle = $entry.Threat_Title
	$startInPath = "$myWD\$department\$room"

	mkdir $startInPath

	$PS1FileLine2 = @("$PS1FileLine2Part1 `"$helpTitle`", `"text`": `"$message`"$PS1FileLine2Part2", "$PS1FileLine2Part1 `"$medTitle`", `"text`": `"$message`"$PS1FileLine2Part2","$PS1FileLine2Part1 `"$threatTitle`", `"text`": `"$message`"$PS1FileLine2Part2")
	#Create .ps1 files
	for ($i=0; $i -lt 3; $i++) {
		$ps1File = $ps1FileNames[$i]
		CreatePS1Files -MakeFile "$startInPath\$ps1File" -WriteLine1 $PS1FileLine1 -WriteLine2 $PS1FileLine2[$i]
		CreateShortcuts -StartPath $startInPath -ShortcutName $shortcutNames[$i] -IconFile $iconFiles[$i] -Department $lowDept -FileName $ps1FileNames[$i]
	}
}
