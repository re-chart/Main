# Initial script created by B de Zeeuw
# Script modified by R. Visser
# Check and Update office if available.

# Functie Sleep-Progress Bar
# Use Sleep-Progress <seconds> 

Function Sleep-Progress($TotalSeconds) {
    $Counter = 0;
    for ($i = 0; $i -lt $TotalSeconds; $i++) {
        $Progress = [math]::Round(100 - (($TotalSeconds - $Counter) / $TotalSeconds * 100));
        Write-Progress -Activity "Waiting..." -Status "$Progress% Complete:" -SecondsRemaining ($TotalSeconds - $Counter) -PercentComplete $Progress;
        Start-Sleep 1
        $Counter++;
    }   
}
 
# Start Office Version Check Script

Write-Host "Checking script directory..." -ForegroundColor Green
if (!(Test-Path "C:\Scripts")) {
    Write-Host "Creating the C:\Scripts directory..." -ForegroundColor Green
    New-Item -ItemType Directory -Path "C:\Scripts"
}
Write-Host "Script directory exists" -ForegroundColor Green

Write-Host "Creating log file..." -ForegroundColor Green
Get-Date | Out-File -FilePath "C:\Scripts\Office365Update.log"
Add-Content -Path "C:\Scripts\Office365Update.log" -Value "Office 365 update process started" 
$logfile = "C:\Scripts\Office365Update$((Get-Date).ToString('dd-MM-yyyy')).log"
Write-Host "Log file created: $logfile" -ForegroundColor Green


Write-Host "Starting the Office 365 update process..." -ForegroundColor DarkMagenta
Add-Content -Path $logfile -Value "Starting the Office 365 update process..."

#Check if the OfficeC2RClient.exe exists
if (Test-Path "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe") {
    Write-Host "OfficeC2RClient.exe exists"
    Write-Host "Let's continue with the update process..." -ForegroundColor Green
    Add-Content -Path $logfile -Value "OfficeC2RClient.exe exists"
} else {
    Write-Host "OfficeC2RClient.exe does not exist"
    Add-Content -Path $logfile -Value "OfficeC2RClient.exe does not exist"
    Write-Host "Exiting the update process..." -ForegroundColor Red
    Add-Content -Path $logfile -Value "Exiting the update process..."

    exit 1
}


Add-Content -Path $logfile -Value "Checking the Office version..."

Sleep-Progress 3

$officeVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration").VersionToReport
Write-Host "Office version: $officeVersion" -ForegroundColor Cyan

# Checks if the Office version is already up to date

Sleep-Progress 3

######### Enter newest version to deploy if else then deploy
if ($officeVersion -eq "16.0.17231.20236") {
    Write-Host "Office is already up to date" -ForegroundColor Green
    Add-Content -Path $logfile -Value "Office is already up to date"
    Write-Host "Exiting the update process..." -ForegroundColor Cyan

#Message for Monitoring
    Write-Host $officeVersion -ForegroundColor Green 
    Add-Content -Path $logfile -Value "Exiting the update process..."
    Start-Sleep 2
    exit 0
}
#Start Update
cmd.exe "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user displaylevel=false forceappshutdown=false

Write-Host "Waiting for the update process to finish..." -ForegroundColor Yellow
if ($LASTEXITCODE -eq 0) {
    Write-Host "Office 365 update was successful" -ForegroundColor Green
    Add-Content -Path $logfile -Value "Office 365 update was successful"
} else {
    Write-Host "Office 365 update failed" -ForegroundColor Red
    Add-Content -Path $logfile -Value "Office 365 update failed"
}
