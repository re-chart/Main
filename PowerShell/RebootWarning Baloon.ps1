# Scripted by R. Visser -> www.richard-visser.nl 
# Reboot warning - System Online for longer time. Warning via baloon tip

Set-ExecutionPolicy Unrestricted -Force
$boot = (Get-CimInstance -ClassName win32_operatingsystem).lastbootuptime


#Aantal dagen aanpassen (-5)) voor 5 dagen
if ($boot -le (Get-Date).AddDays(-5)) #
	{


Add-Type -AssemblyName System.Windows.Forms
$global:balmsg = New-Object System.Windows.Forms.NotifyIcon
$balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
$balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
$balmsg.BalloonTipText = ‘Uw Computer is al langere tijd niet opnieuw opgestart. Ons advies is dit wel te doen voor een juiste werking en beveiliging'
$balmsg.BalloonTipTitle = "Attention - Bericht van uw systeembeheerder"
$balmsg.Visible = $true
$balmsg.ShowBalloonTip(30000) }