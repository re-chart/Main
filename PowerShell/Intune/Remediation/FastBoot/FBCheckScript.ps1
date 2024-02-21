#Created by R. Visser -> www.richard-visser.nl
#FastbootCheck
if
((Get-ItemPropertyValue -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -Name 'HiberbootEnabled' ) -eq 0)
{ Write-Host "FastBoot Disabled" 
Exit 0 } 
else 
{ Write-Host "FastBoot Enabled" 
Exit 1 };
