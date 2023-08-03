<#
	Author  : Matthew Starin
	Date	: 8-1-23
	Script	: Software Inventory 
	Purpose	: List of all software installed on a laptop.		
#>

#variables 
	$vUserName = (Get-Item env:\username).Value
	$vComputerName = (Get-Item env:\Computername).Value
	$filepath = (Get-ChildItem env:\userprofile).value
	$name = (Get-Item env:\Computername).Value 
## Html Style
	$a = "<style>"
	$a = $a + "BODY{background-color:Lavender ;}"
	$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
	$a = $a + "TH{border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:thistle}"
	$a = $a + "TD{border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:PaleGoldenrod}"
	$a = $a + "</style>"
# removing old HTML Report if exists
	if (test-Path $filepath\$name.html) { remove-Item $filepath\$name.html;
	Write-Host -ForegroundColor white -BackgroundColor Red    "Old file removed"
	}
# Running Command 
	ConvertTo-Html -Title "Software Information for $name" -Body "<h1> Computer Name : $name </h1>" >  "$filepath\$name.html"
	Get-WmiObject win32_Product -ComputerName $name | Select Name,Version,PackageName,Installdate,Vendor | Sort Installdate -Descending `
	                                         | ConvertTo-html  -Head $a -Body "<H2> Software Installed</H2>" >> "$filepath\$name.html"							 
	$Report = "The Report is generated On  $(get-date) by $((Get-Item env:\username).Value) on computer $((Get-Item env:\Computername).Value)"
	$Report  >> "$filepath\$name.html" 
## Opening file and the file 
	write-Host "file is saved in $filepath and the name of file is $name.html" -ForegroundColor Cyan
	invoke-Expression "$filepath\$name.html" 
## END of the SCRIPT ## 





