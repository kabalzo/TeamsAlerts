$dataPath = ".\AlertTemplate.csv"
$data = Import-Csv -Path $dataPath
$PS1FileLine2Part1 = "Invoke-RestMethod -Method post -ContentType `'Application/Json`' -Body `'{`"title`":"
$PS1FileLine2Part2 = "}`' -Uri `$myTeamsWebHook"
$iconFiles = @("Yellow.ico", "Blue.ico", "Red.ico")
$shortcutNames = @("Help Needed", "Medical", "Threat")

$department = Split-Path -Path $pwd -Leaf
$lowDept = $department.ToLower()
	
try {
	mkdir "$pwd\..\..\..\$department" -ErrorAction Stop
}
catch {
	while ($true) {
		$choice = Read-Host "Department: $department already exists. Do you want to overwrite it [y] [n]?"
		if ($choice -eq "y") {
			Remove-Item "$pwd\..\..\..\$department" -Recurse -Force
			mkdir "$pwd\..\..\..\$department"
			break
		}
		elseif ($choice -eq "n") {
			Write-Host "Goodbye"
			timeout /t 5
			exit
		}
		else {
			Write-Host "Invalid selection" -ForegroundColor Red
		}
	}
}
function CreateShortcuts ([string]$StartPath, [string]$ShortcutName, [string]$IconFile, [string]$Department, [string]$FileName) {
	$arguments = "-File `"$FileName`""
	$WshShell = New-Object -comObject WScript.Shell
	$Shortcut = $WshShell.CreateShortcut("$StartPath\Alert - $ShortcutName.lnk")
	$Shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
	$Shortcut.Arguments = $arguments
	$Shortcut.IconLocation = "$pwd\..\..\Icons\Alert - $IconFile"
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
	$startInPath = "$pwd\..\..\..\$department\$room"

	mkdir $startInPath

	$PS1FileLine2 = @("$PS1FileLine2Part1 `"$helpTitle`", `"text`": `"$message`"$PS1FileLine2Part2", "$PS1FileLine2Part1 `"$medTitle`", `"text`": `"$message`"$PS1FileLine2Part2","$PS1FileLine2Part1 `"$threatTitle`", `"text`": `"$message`"$PS1FileLine2Part2")
	#Create .ps1 files
	for ($i=0; $i -lt 3; $i++) {
		$ps1File = $ps1FileNames[$i]
		CreatePS1Files -MakeFile "$startInPath\$ps1File" -WriteLine1 $PS1FileLine1 -WriteLine2 $PS1FileLine2[$i]
		CreateShortcuts -StartPath $startInPath -ShortcutName $shortcutNames[$i] -IconFile $iconFiles[$i] -Department $lowDept -FileName $ps1FileNames[$i]
	}
}
Write-Host "Setup complete"
timeout /t -1
