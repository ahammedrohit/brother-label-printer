$programName = "b-pac"

# Check if any program named or containing "b-pac" is installed
$programInstalled = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE '%$programName%'"

if ($programInstalled) {
    # Program is already installed
    Write-Output "Already Installed"
}
else {
    # Program is not installed, install it
    $msiFile = Join-Path $PSScriptRoot "lib\bcciw34007_64.msi"

    if (Test-Path $msiFile) {
        Write-Output "Installing b-pac..."
        $startProcessParams = @{
            FilePath     = "msiexec.exe"
            ArgumentList = "/i `"$msiFile`" /quiet"
            Verb         = "RunAs"
            Wait         = $true
            PassThru     = $false
        }
        
        try {
            Start-Process @startProcessParams
            $programInstalled = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE '%$programName%'"
            
            if ($programInstalled) {
                Write-Output "b-pac installed successfully."
            }
            else {
                Write-Output "Failed to install b-pac."
            }
        }
        catch {
            Write-Output "Failed to start the installation process."
        }
    }
    else {
        Write-Output "Unable to find the bcciw34007_64.msi file in the lib folder."
    }
}
