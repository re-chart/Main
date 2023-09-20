#$Appname = "Googlechrome"
# Check Choco is installed
if (-Not (Get-Command -Name Choco -ErrorAction SilentlyContinue)) {
    
    # Choco is niet geinstalleerd
          $progressPreference = 'silentlyContinue'
Write-Information "Install Choco"

Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

    # Install succesvol geweest?
    if ($?) {
        Write-Host "Choco succesvol geinstalleerd"
    } else {
        Write-Host "Choco failed"
    }
} else {
    Write-Host "Choco is al geinstalleerd"
}
#choco install $Appname