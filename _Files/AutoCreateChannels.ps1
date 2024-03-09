$data = Import-Csv -Path .\xyz.csv
#$groupID = Read-Host "Enter the groupID for Teams channel you have permisson for"
$groupID = ""
Write-Host "Checking for MicrosoftTeams module..."

function CreatChannels {
	Write-Host "Proceeding with configuration`n"
	try {
		Write-Host "Please follow the authentication prompt..."
		$auth = Connect-MicrosoftTeams -ErrorAction Stop
		Write-Host "Successfully authenticated`n" -ForegroundColor Green
	}
	catch {
		Write-Host "Failed to authenticate" -ForegroundColor Red
		Read-Host "Press any key to exit"
		exit
	}

	Write-Host "Attempting to create teams channels..."
	foreach ($line in $data) {
		try {
			$name = $line.Name
			$room = $line.Room
			$channelName = "$room - $name"
			$newTeam = New-TeamChannel -GroupId $groupID -DisplayName $channelName -MembershipType Standard -ErrorAction Stop
			Write-Host "Successfully created channel: $channelName" -ForegroundColor Green
		}
		catch {
			Write-Host "Failed to create channel: $channelName" -ForegroundColor Red
		}
	}
	Write-Host "Setup complete`n"
}

$teamsModuleInstalled = Get-InstalledModule -Name "MicrosoftTeams" -ErrorAction SilentlyContinue
if ($teamsModuleInstalled -eq $null) {
	try {
		Write-Host "MicrosoftTeams module not found" -ForegroundColor Red
		Write-Host "Attempting to install MicrosoftTeams module..."
		Install-Module -Name MicrosoftTeams -Force -AllowClobber -ErrorAction Stop
		$teamsModuleInstalled = Get-InstalledModule -Name "MicrosoftTeams" -ErrorAction SilentlyContinue
		if ($teamsModuleInstalled -ne $null) {
			Write-Host "Successfully installed MicrosoftTeams module`n" -ForegroundColor Green
			CreatChannels
		}
		else {
			Write-Host "Failed to install the MicrosoftTeams Powershell module. Configure channels manually." -ForegroundColor Red
			Read-Host "Press any key to exit"
			exit
		}
	}
	catch {
		Write-Host "Failed to install the MicrosoftTeams Powershell module. Configure channels manually." -ForegroundColor Red
		Read-Host "Press any key to exit"
		exit
	}
}
else {
	Write-Host "MicrosoftTeams module already installed" -ForegroundColor Green
	CreatChannels
}
timeout /t -1

