# Initial script created by Bert de Zeeuw
# Script modified by R. Visser

# Functie Sleep-Progress Bar
#Use Sleep-Progress <seconds> 

Function Sleep-Progress($TotalSeconds) {
    $Counter = 0;
    for ($i = 0; $i -lt $TotalSeconds; $i++) {
        $Progress = [math]::Round(100 - (($TotalSeconds - $Counter) / $TotalSeconds * 100));
        Write-Progress -Activity "Waiting..." -Status "$Progress% Complete:" -SecondsRemaining ($TotalSeconds - $Counter) -PercentComplete $Progress;
        Start-Sleep 1
        $Counter++;
    }   
}
 
#Sleep-Progress 5


#Start Office Version Check Script

Write-Host "Checking script directory..." -ForegroundColor Green
if (!(Test-Path "C:\T4")) {
    Write-Host "Creating the C:\T4 directory..." -ForegroundColor Green
    New-Item -ItemType Directory -Path "C:\T4"
}
Write-Host "Script directory exists" -ForegroundColor Green

Write-Host "Creating log file..." -ForegroundColor Green
Get-Date | Out-File -FilePath "C:\T4\Office365Update.log"
Add-Content -Path "C:\T4\Office365Update.log" -Value "Office 365 update process started" 
$logfile = "C:\T4\Office365Update$((Get-Date).ToString('dd-MM-yyyy')).log"
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

#Checks if the Office version is already up to date

Sleep-Progress 3

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