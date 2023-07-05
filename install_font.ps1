#check if PowerShell is installed
function IsPowerShellInstalled {
    return $PSVersionTable.PSVersion.Major -ge 7
}

# Attempt to install PowerShell using winget
$wingetInstallCommand = "winget install --id Microsoft.Powershell --source winget"
Invoke-Expression -Command $wingetInstallCommand

if (IsPowerShellInstalled) {
    Write-Host "PowerShell installed successfully."
}
else {
    Write-Host "PowerShell installation using winget failed. Falling back to manual installation."

    $downloadUrl = "https://github.com/PowerShell/PowerShell/releases/download/v7.3.4/PowerShell-7.3.4-win-x64.msi"
    $destinationPath = "$env:TEMP\PowerShell-7.3.4-win-x64.msi"

    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $destinationPath

        Start-Process -FilePath msiexec.exe -ArgumentList "/i `"$destinationPath`" /qn /norestart" -Wait -Verb RunAs

        Remove-Item -Path $destinationPath -Force

        Write-Host "PowerShell installed successfully."
    }
    catch {
        Write-Host "Error occurred during PowerShell installation: $_"
    }
}
