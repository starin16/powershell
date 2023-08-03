<#
Script ,that gets all laptops from Active Directory Organization Unit, and gets their serial numbers with export to csv file
#>


#---------------Job------------------
$GetSerialBlock = {
	Param($Laptopname)
	$LapInfo = New-Object -TypeName PSObject
	if(Test-Connection $Laptopname -Quiet)
	{
		$BiosWmiObj = Get-WmiObject win32_bios -ComputerName $Laptopname
		Add-Member -InputObject $LapInfo -MemberType NoteProperty -Name LaptopName -Value $Laptopname
		Add-Member -InputObject $LapInfo -MemberType NoteProperty -Name Manufacturer -Value $BiosWmiObj.Manufacturer
		Add-Member -InputObject $LapInfo -MemberType NoteProperty -Name SerialNumber -Value $BiosWmiObj.SerialNumber
	}
	else 
	{
	    $NoConnection = "Can't connect"
	    Add-Member -InputObject $LapInfo -MemberType NoteProperty -Name LaptopName -Value $Laptopname
	    Add-Member -InputObject $LapInfo -MemberType NoteProperty -Name Manufacturer -Value "$NoConnection"
		Add-Member -InputObject $LapInfo -MemberType NoteProperty -Name SerialNumber -Value "$NoConnection"
	}
	return $LapInfo
}


#----------------Variables--------------
$CSVpath = "____Path______"                                                                    #path to file like C:\TMP\info.csv
$Laptops = Get-ADComputer -SearchBase '____AD_OU_____' -Filter 'ObjectClass -eq "Computer"'    #Path to Organisation Unit With Your Laptops like OU=Laptops,DC=Domain,DC=com
$InfoHash = @()
$MaxThreads = 20                                                                               #Number of threads you want to run simultaneously


#---------------Job Execution------------
foreach($LaptopObj in $Laptops)
{
	$Laptop = $LaptopObj.name.ToString()
	Write-Host "Working on $Laptop"
	Start-Job -ScriptBlock $GetSerialBlock -ArgumentList $Laptop
        While (@(Get-Job | Where { $_.State -eq "Running" }).Count -ge $MaxThreads) {    #|  Thread number limitation  
        Write-Verbose "Waiting for open thread...($MaxThreads Maximum)"                  #|
        Start-Sleep -Seconds 3                                                           #|
        }                                                                                #|
}

#-------------Getting Results-------------
Get-Job | Wait-Job
$InfoHash += Get-Job | Receive-Job
Get-Job | Remove-Job
$InfoHash |select LaptopName, Manufacturer, SerialNumber| Export-Csv $CSVpath

