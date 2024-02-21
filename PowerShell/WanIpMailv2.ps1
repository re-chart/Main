# Run parameter: Powershell.exe -file .\wanip.ps1 -UpdateIP <current wan ip> if else script will send mail with new IP.
# I created a scheduled job on my host to check every day for IP because ISP is changing it sometimes.
param (
    [Parameter(Mandatory=$true)]
    [string]$UpdateIP
)


# Check dir if excists if not create
$EmailTo = "email@address.com"


$WAN = (Invoke-WebRequest "myexternalip.com/raw").Content

if ([string]::IsNullOrEmpty($_)) {


try {
    if ($WAN -notmatch $UpdateIP) {
        Send-MailMessage -To $EmailTo -From $EmailTo -Subject "WAN IP has changed" -Body "The new WAN IP is $WAN - also change Scheduled JOB" -SmtpServer smtpserver
    }  
    else {
        Write-Host "WAN IP has not changed" -ForegroundColor Green
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}} 
