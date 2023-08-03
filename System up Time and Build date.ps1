PowerShell
"Build Date- " + [System.Management.ManagementDateTimeconverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem).installdate) 
 
$sut = (New-TimeSpan ([System.Management.ManagementDateTimeconverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem).LastBootUpTime)) (Get-Date)) 
 
Switch ($sut) 
{ 
   {$_.days -eq 1} {$Days="1 Day"} 
   {$_.days -gt 1} {$Days=[string]$_.days + " Days"} 
   {$_.hours -eq 1} {$Hrs="1 Hour"} 
   {$_.hours -gt 1} {$Hrs=[string]$_.hours + " Hours"} 
   {$_.Minutes -eq 1} {$Mins="1 Minute "} 
   {$_.Minutes -gt 1} {$Mins=[string]$_.minutes + " Minutes"} 
   {$_.Seconds -eq 1} {$Secs="1 Second "} 
   {$_.Seconds -gt 1} {$Secs=[string]$_.seconds + " Seconds"} 
} 
 
"System Uptime- $Days $Hrs $Mins $Secs"