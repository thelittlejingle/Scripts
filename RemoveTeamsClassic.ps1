##############################################
#                                            #
#   Script Name: RemoveTeamsClassic.ps1      #
#   Description: Brief description here      #
#   Version: 1.0.0                           #
#   Created by: thelittlejingle              #
#   Date: 2024-08-10                         #
#                                            #
##############################################
# This script automates the removal of the old Microsoft Teams (classic) from user profiles on a Windows machine.
# It deletes the Teams folder within each user's AppData directory, uninstalls the Teams Machine-Wide Installer,
# and removes any remaining installer files from the Program Files directory.


# Define the folder name to remove
$folderToRemove = "Teams"

# Get all user profiles
$userProfiles = Get-WmiObject Win32_UserProfile | Where-Object { $_.Special -eq $false }

foreach ($profile in $userProfiles) {
    $appDataPath = Join-Path -Path $profile.LocalPath -ChildPath "AppData\Local\Microsoft\$folderToRemove"

    # Check if the folder exists and remove it
    if (Test-Path $appDataPath) {
        try {
            Remove-Item -Path $appDataPath -Recurse -Force
            Write-Host "Removed: $appDataPath"
        }
        catch {
            # Output the error message using formatted string
            $errorMsg = "Failed to remove {0}: {1}" -f $appDataPath, $_.Exception.Message
            Write-Error $errorMsg
        }
    }
    else {
        Write-Host "Folder does not exist: $appDataPath"
    }
}

# Function to find the GUID for Teams Machine-Wide Installer
function Get-TeamsInstallerGUID {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $teamsGUID = $null

    # Search through registry subkeys
    Get-ChildItem -Path $regPath | ForEach-Object {
        $key = $_.PSPath
        $displayName = (Get-ItemProperty -Path $key).DisplayName
        if ($displayName -match "Teams Machine-Wide Installer") {
            $teamsGUID = (Get-ItemProperty -Path $key).PSChildName
            return
        }
    }

    return $teamsGUID
}

# Get the GUID of Teams Machine-Wide Installer
$teamsInstallerGUID = Get-TeamsInstallerGUID

if ($teamsInstallerGUID) {
    try {
        $uninstallCommand = "msiexec.exe /x $teamsInstallerGUID /qn"
        Write-Host "Running command: $uninstallCommand"
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $teamsInstallerGUID /qn" -Wait -NoNewWindow
        Write-Host "Uninstall command executed."
    }
    catch {
        Write-Error ("Failed to run uninstall command: {0}" -f $_.Exception.Message)
    }
}
else {
    Write-Host "Teams Machine-Wide Installer not found. Skipping uninstall."
}

# Remove Teams Installer folder if it exists
$installerFolderPath = "C:\Program Files (x86)\Teams Installer"

if (Test-Path $installerFolderPath) {
    try {
        Remove-Item -Path $installerFolderPath -Recurse -Force
        Write-Host "Removed: $installerFolderPath"
    }
    catch {
        # Output the error message using formatted string
        $errorMsg = "Failed to remove {0}: {1}" -f $installerFolderPath, $_.Exception.Message
        Write-Error $errorMsg
    }
}
else {
    Write-Host "Teams Installer folder does not exist: $installerFolderPath"
}

Write-Host "Script completed."
