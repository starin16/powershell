﻿<# ==============================================================================================
 
    Name : Server Inventory (Get-ServerInventory.ps1)
    Description : Get informations from remote servers with WMI and ouput in an Excel File                 
 
    Author : Matthew Starin
    
    * Select list of servers from a CSV file with an OpenFileDialog
    * Get remotely Servers informations with WMI and Powershell :
    * General (Domain, role in the domain, hardware manufacturer, type and model, cpu number, memory capacity, operating system and sp level)
    * System (BIOS name, BIOS version, hardware serial number, time zone, WMI version, virtual memory file location, virtual memory current usage, virtual memory peak usage and virtual memory allocated)
    * Processor (Processor(s), processor type, family, speed in Mhz, cache size in GB and socket number)
    * Memory (Bank number, label, capacity in GB, form and type)
    * Disk (Disk type, letter, capacity in GB, free space in GB + display a chart Excel)
    * Network (Network card, DHCP enable or not, Ip address, subnet mask, default gateway, Dns servers, Dns registered or not, primary and secondary wins and wins lookup or not) 
    * Installed Programs (Display name, version, install location and publisher) 
    * Share swith NTFS rights (Share name, user account, rights, ace flags and ace type) 
    * Services (Display name, name, start by, start mode and path name)
    * Scheduled Tasks (Name, last run time, next run time and run as)
    * Printers (Locationm, name, printer state and status, share name and system name)
    * Process (Name, Path and sessionID)
    * Local Users (Groups, users)
    * ODBC Configured (dsn, Server, Port, DatabaseFile, DatabaseName, UID, PWD, Start, LastUser, Database, DefaultLibraries, DefaultPackage, DefaultPkgLibrary, System, Driver, Description)
    * ODBC Drivers Installed (Driver, DriverODBCVer, FileExtns, Setup)    
    * MB to GB conversion
    * Display of the progress of the script

# ============================================================================================== #>

# Load the Excel Assembly, Locally or from GAC
try {
    Add-Type -ASSEMBLY "Microsoft.Office.Interop.Excel"  | out-null
}
catch {
    [Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Excel") | out-null
}
$xlConditionValues=[Microsoft.Office.Interop.Excel.XLConditionValueTypes]
$xlTheme=[Microsoft.Office.Interop.Excel.XLThemeColor]
$xlChart=[Microsoft.Office.Interop.Excel.XLChartType]
$xlIconSet=[Microsoft.Office.Interop.Excel.XLIconSet]
$xlDirection=[Microsoft.Office.Interop.Excel.XLDirection]

# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'Read-OpenFileDialog' - Open an open File Dialog box
# ________________________________________________________________________
Function Read-OpenFileDialog([string]$InitialDirectory, [switch]$AllowMultiSelect) {      
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog        
    $openFileDialog.ShowHelp = $True    # http://www.sapien.com/blog/2009/02/26/primalforms-file-dialog-hangs-on-windows-vista-sp1-with-net-30-35/
    $openFileDialog.initialDirectory = $initialDirectory
    $openFileDialog.filter = "csv files (*.csv)|*.csv|All files (*.*)| *.*"
    $openFileDialog.FilterIndex = 1
    $openFileDialog.ShowDialog() | Out-Null
    return $openFileDialog.filename
}
# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'Translate-AccessMask' - Translate integer value in string
# ________________________________________________________________________
Function Translate-AccessMask($val) {
    Switch ($val)
    {
        2032127 {"FullControl"; break}
        1179785 {"Read"; break}
        1180063 {"Read, Write"; break}
        1179817 {"ReadAndExecute"; break}
        -1610612736 {"ReadAndExecuteExtended"; break}
        1245631 {"ReadAndExecute, Modify, Write"; break}
        1180095 {"ReadAndExecute, Write"; break}
        268435456 {"FullControl (Sub Only)"; break}
        default {$AccessMask = $val; break}
    }
}
# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'Translate-AceType' - Translate integer value in string
# ________________________________________________________________________
Function Translate-AceType($val) {
    Switch ($val)
    {
        0 {"Allow"; break}
        1 {"Deny"; break}
        2 {"Audit"; break}
    }
}
# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'Translate-AceFlagse' - Translate integer value in string
# ________________________________________________________________________
<#  OBJECT_INHERIT_ACE
    1 (0x1)
    Noncontainer child objects inherit the ACE as an effective ACE.
    For child objects that are containers, the ACE is inherited as an inherit-only ACE unless the NO_PROPAGATE_INHERIT_ACE bit flag is also set.
    CONTAINER_INHERIT_ACE
    2 (0x2)
    Child objects that are containers, such as directories, inherit the ACE as an effective ACE. The inherited ACE is inheritable unless the NO_PROPAGATE_INHERIT_ACE bit flag is also set.
    NO_PROPAGATE_INHERIT_ACE
    4 (0x4)
    If the ACE is inherited by a child object, the system clears the OBJECT_INHERIT_ACE and CONTAINER_INHERIT_ACE flags in the inherited ACE. This prevents the ACE from being inherited by subsequent generations of objects.
    INHERIT_ONLY_ACE
    8 (0x8)
    Indicates an inherit-only ACE which does not control access to the object to which it is attached. If this flag is not set, the ACE is an effective ACE which controls access to the object to which it is attached.
    Both effective and inherit-only ACEs can be inherited depending on the state of the other inheritance flags.
    INHERITED_ACE
    16 (0x10)
    The system sets this bit when it propagates an inherited ACE to a child object.
    Access these the same way. You can break them out using the bitwise AND operator or just test for the totals #>
Function Translate-AceFlags($val) {
    Switch ($val)
    {
        0 {"0"}
        1 {"Noncontainer child objects inherit"; break}
        2 {"Containers will inherit and pass on"; break}
        3 {"Containers AND Non-containers will inherit and pass on"; break}       
    }
}
# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'Get-NtfsRights' - Enumerates NTFS rights of a folder
# ________________________________________________________________________
Function Get-NtfsRights($name,$path,$comp) {
	$path = [regex]::Escape($path)
	$share = "\\$comp\\$name"
	$wmi = gwmi Win32_LogicalFileSecuritySetting -filter "path='$path'" -ComputerName $comp
	$wmi.GetSecurityDescriptor().Descriptor.DACL | where {$_.AccessMask -as [Security.AccessControl.FileSystemRights]} |select `
                @{name="ShareName";Expression={$share}},
				@{name="Principal";Expression={"{0}\{1}" -f $_.Trustee.Domain,$_.Trustee.name}},
				@{name="Rights";Expression={Translate-AccessMask $_.AccessMask }},
				@{name="AceFlags";Expression={Translate-AceFlags $_.AceFlags }},
				@{name="AceType";Expression={Translate-AceType $_.AceType }}
				
}
# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'listProgramsInstalled' - get info in registry 
# ________________________________________________________________________
Function listProgramsInstalled ($uninstallKey) {
    $array = @()

    $computername = $strComputer           
    $remoteBaseKeyObject = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computername)     
    $remoteBaseKey = $remoteBaseKeyObject.OpenSubKey($uninstallKey)             
    $subKeys = $remoteBaseKey.GetSubKeyNames()            
    foreach($key in $subKeys){            
        $thisKey=$UninstallKey+"\\"+$key          
        $thisSubKey=$remoteBaseKeyObject.OpenSubKey($thisKey) 
        $psObject = New-Object PSObject        
        $psObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
        $psObject | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
        $psObject | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))
        $psObject | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
        $array += $psObject
    }           
    $array
}

# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'getTasks' - get scheduled tasks on remote server 
# ________________________________________________________________________
Function getTasks($path) {
    $out = @()
    # Get root tasks
    $schedule.GetFolder($path).GetTasks(0) | % {
        $xml = [xml]$_.xml
        $out += New-Object psobject -Property @{
            "Name" = $_.Name
            "Path" = $_.Path
            "LastRunTime" = $_.LastRunTime
            "NextRunTime" = $_.NextRunTime
            "Actions" = ($xml.Task.Actions.Exec | % { "$($_.Command) $($_.Arguments)" }) -join "`n"
            "RunAs" = ($xml.Task.Principals.principal.userID)
        }
    }
    # Get tasks from subfolders
    $schedule.GetFolder($path).GetFolders(0) | % {
        $out += getTasks($_.Path)
    }    
    $out
}
# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'getLocalUsersInGroup' - get local users in groups 
# ________________________________________________________________________
Function getLocalUsersInGroup {
    if($domainRole -le 3) {
        $serverADSIObject = [ADSI]"WinNT://$strComputer,computer"
        $localUserinGroups=@()
        $serverADSIObject.psbase.children | Where { $_.psbase.schemaClassName -eq 'group' } |`
            foreach {
                $group =[ADSI]$_.psbase.Path
                $group.psbase.Invoke("Members") | `
                foreach {$localUserinGroups += New-Object psobject -property @{Group = $group.Name;User=(($_.GetType().InvokeMember("Adspath", 'GetProperty', $null, $_, $null)) -replace "WinNT://","")}}
            }
    }
    else {
        $localUserinGroups = @()
    }
    $localUserinGroups
}
# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'listODBCConfigured' - get ODBC connections configured 
# ________________________________________________________________________
Function listODBCConfigured ($odbcConfigured) {
    $computername = $strComputer 
    $arrayConfigured = @()           
    $remoteBaseKeyObject = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computername)     
    $remoteBaseKey = $remoteBaseKeyObject.OpenSubKey($odbcConfigured)             
    $subKeys = $remoteBaseKey.GetSubKeyNames()            
    foreach($key in $subKeys){            
        $thisKey=$odbcConfigured+"\\"+$key          
        $thisSubKey=$remoteBaseKeyObject.OpenSubKey($thisKey)         
        $psObjectConfigured = New-Object PSObject
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $computername
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "DSN" -Value $($thisSubKey.GetValue("dsn"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "Server" -Value $($thisSubKey.GetValue("Server"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "Port" -Value $($thisSubKey.GetValue("Port"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "DatabaseFile" -Value $($thisSubKey.GetValue("DatabaseFile"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "DatabaseName" -Value $($thisSubKey.GetValue("DatabaseName"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "UID" -Value $($thisSubKey.GetValue("UID"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "PWD" -Value $($thisSubKey.GetValue("PWD"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "Start" -Value $($thisSubKey.GetValue("Start"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "LastUser" -Value $($thisSubKey.GetValue("LastUser"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "Database" -Value $($thisSubKey.GetValue("Database"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "DefaultLibraries" -Value $($thisSubKey.GetValue("DefaultLibraries"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "DefaultPackage" -Value $($thisSubKey.GetValue("DefaultPackage"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "DefaultPkgLibrary" -Value $($thisSubKey.GetValue("DefaultPkgLibrary"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "System" -Value $($thisSubKey.GetValue("System"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "Driver" -Value $($thisSubKey.GetValue("Driver"))
        $psObjectConfigured | Add-Member -MemberType NoteProperty -Name "Description" -Value $($thisSubKey.GetValue("Description"))
        $arrayConfigured += $psObjectConfigured
    }           
    $arrayConfigured    
}
# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'listODBCInstalled' - get ODBC connections installed 
# ________________________________________________________________________
Function listODBCInstalled ($odbcDriversInstalled) {
    $computername = $strComputer 
    $arrayInstalled = @()       
    $remoteBaseKeyObject = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computername)     
    $remoteBaseKey = $remoteBaseKeyObject.OpenSubKey($odbcDriversInstalled)             
    $subKeys = $remoteBaseKey.GetSubKeyNames()            
    foreach($key in $subKeys){            
        $thisKey=$odbcDriversInstalled+"\\"+$key          
        $thisSubKey=$remoteBaseKeyObject.OpenSubKey($thisKey)         
        $psObjectInstalled = New-Object PSObject
        $psObjectInstalled | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $computername
        $psObjectInstalled | Add-Member -MemberType NoteProperty -Name "Driver" -Value $($thisSubKey.GetValue("Driver"))
        $psObjectInstalled | Add-Member -MemberType NoteProperty -Name "DriverODBCVer" -Value $($thisSubKey.GetValue("DriverODBCVer"))
        $psObjectInstalled | Add-Member -MemberType NoteProperty -Name "FileExtns" -Value $($thisSubKey.GetValue("FileExtns"))
        $psObjectInstalled | Add-Member -MemberType NoteProperty -Name "Setup" -Value $($thisSubKey.GetValue("Setup"))
        $arrayInstalled += $psObjectInstalled
    }           
    $arrayInstalled    
}

# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'ListFile' - get server based on a CSV file
# ________________________________________________________________________
Function ListFile {	
    $fileOpen = Read-OpenFileDialog 
    if($fileOpen -ne '') {	
		$colComputers = Import-Csv $fileOpen
    }
    $colComputers
}

# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Function Name 'FormatOutput' - format output in an worksheet
# ________________________________________________________________________
Function FormatOutput ($items, $worksheet, $rowNumber, $columnNumber, $strComputer) {	                                    
    $worksheet.Cells.Item($rowNumber, 1) = $strComputer   
    $items.psobject.properties |
    %{          	                                                                    
        if($_.Name -ne 'DomainRole' -and $_.Name -ne 'TotalPhysicalMemory' -and $_.Name -ne 'Capacity' -and $_.Name -ne 'DriveType' -and $_.Name -ne 'Size' `
            -and $_.Name -ne 'FreeSpace' -and $_.Name -ne 'Group') {         
            $worksheet.Cells.Item($rowNumber, $columnNumber) = $items.($_.Name) 
        }
        else {
            if($_.Name -eq 'DomainRole') {
                Switch($items.($_.Name)) {
                    0{$worksheet.Cells.Item($rowNumber, $columnNumber) = "Stand Alone Workstation"}
        			1{$worksheet.Cells.Item($rowNumber, $columnNumber) = "Member Workstation"}
        			2{$worksheet.Cells.Item($rowNumber, $columnNumber) = "Stand Alone Server"}
        			3{$worksheet.Cells.Item($rowNumber, $columnNumber) = "Member Server"}
        			4{$worksheet.Cells.Item($rowNumber, $columnNumber) = "Back-up Domain Controller"}
        			5{$worksheet.Cells.Item($rowNumber, $columnNumber) = "Primary Domain Controller"}
        			default{"Undetermined"}
                }
                $domainRole = $items.($_.Name)
            }
            else {
                if($_.Name -eq 'TotalPhysicalMemory') {$worksheet.Cells.Item($rowNumber, $columnNumber) = [math]::round($items.($_.Name)/1024/1024/1024, 0)}  
                else {                                   
                    if($_.Name -eq 'Capacity') {$worksheet.Cells.Item($rowNumber, $columnNumber) = [math]::round($item.($_.Name)/1024/1024/1024, 0)}
                    else {                            
                        if($_.Name -eq 'DriveType') {
                            Switch($item.($_.Name)) {
                                2{$worksheet.Cells.Item($rowNumber, $columnNumber) = "Floppy"}
                        		3{$worksheet.Cells.Item($rowNumber, $columnNumber) = "Fixed Disk"}
                        		5{$worksheet.Cells.Item($rowNumber, $columnNumber) = "Removable Media"}
                        		default{"Undetermined"}
                    		}
                        }
                        else {                    
                            if($_.Name -eq 'Size') {
                                $worksheet.Cells.Item($rowNumber, $columnNumber) = [math]::round($item.Size/1024/1024/1024, 0)
                            } 
                            else {
                                if($_.Name -eq 'FreeSpace') {
                                    $worksheet.Cells.Item($rowNumber, $columnNumber) = [math]::round($item.FreeSpace/1024/1024/1024, 0)
                                } 
                                else {
                                    if($_.Name -eq 'Group') {                                                                 
                                        $worksheet.Cells.Item($rowNumber, $columnNumber) = $($item.($_.Name))
                                    }
                                }
                            } 
                        }                            
                    }                    
                }
            }
        }                         
        $columnNumber = $columnNumber + 1 
    }      
    return $columnNumber
}

# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Let's the script begins !
# ________________________________________________________________________

$colComputers = ListFile	
		
Write-Progress -Activity "Getting Inventory" -status "Running..." -id 1

# Launch Excel
<#
Set-Culture en-US # https://social.technet.microsoft.com/Forums/windowsserver/en-US/6bfeeb89-8fa8-4e6b-8588-cabe7f7291ea/excel-with-powershell-and-this-big-problem-of-locale
#>
$currentThread = [System.Threading.Thread]::CurrentThread
$culture = [System.Globalization.CultureInfo]::InvariantCulture
$currentThread.CurrentCulture = "en-US"
$currentThread.CurrentUICulture = "en-US"
$excelDocument = New-Object -Com Excel.Application

$excelDocument.visible = $False

# Create worksheets
$workbook = $excelDocument.Workbooks.Add()

for($i=1; $i -le 14; $i++){$worksheet = $workbook.Worksheets.Add();}

$worksheet1 = $workbook.Worksheets.Item(1);$worksheet2 = $workbook.WorkSheets.Item(2);$worksheet3 = $workbook.WorkSheets.Item(3);$worksheet4 = $workbook.WorkSheets.Item(4)
$worksheet5 = $workbook.WorkSheets.Item(5);$worksheet6 = $workbook.WorkSheets.Item(6);$worksheet7 = $workbook.WorkSheets.Item(7);$worksheet8 = $workbook.WorkSheets.Item(8)
$worksheet9 = $workbook.WorkSheets.Item(9);$worksheet10 = $workbook.WorkSheets.Item(10);$worksheet11 = $workbook.WorkSheets.Item(11);$worksheet12 = $workbook.WorkSheets.Item(12)
$worksheet13 = $workbook.WorkSheets.Item(13);$worksheet14 = $workbook.WorkSheets.Item(14);$worksheet15 = $workbook.WorkSheets.Item(15)

# Name the worksheets
$worksheet1.Name = "General";$worksheet2.Name = "System";$worksheet3.Name = "Processor";$worksheet4.Name = "Memory";$worksheet5.Name = "Disk";$worksheet6.Name = "Network";$worksheet7.Name = "Installed Programs"
$worksheet8.Name = "Shares";$worksheet9.Name = "Services";$worksheet10.Name = "Scheduled Tasks";$worksheet11.Name = "Printers";$worksheet12.Name = "Process";$worksheet13.Name = "Local Users"
$worksheet14.Name = "ODBC Configured";$worksheet15.Name = "ODBC Drivers Installed"

# Create Heading for General worksheet
$worksheet1.Cells.Item(1,1) = "Device_Name";$worksheet1.Cells.Item(1,2) = "Domain";$worksheet1.Cells.Item(1,3) = "Role";$worksheet1.Cells.Item(1,4) = "HW_Make";$worksheet1.Cells.Item(1,5) = "HW_Model"
$worksheet1.Cells.Item(1,6) = "HW_Type";$worksheet1.Cells.Item(1,7) = "CPU_Count";$worksheet1.Cells.Item(1,8) = "Memory_GB";$worksheet1.Cells.Item(1,9) = "Operating_System";$worksheet1.Cells.Item(1,10) = "SP_Level"

# Create Heading for System worksheet
$worksheet2.Cells.Item(1,1) = "Device_Name";$worksheet2.Cells.Item(1,2) = "BIOS_Name";$worksheet2.Cells.Item(1,3) = "BIOS_Version";$worksheet2.Cells.Item(1,4) = "HW_Serial_#";$worksheet2.Cells.Item(1,5) = "Time_Zone"
$worksheet2.Cells.Item(1,6) = "WMI_Version";$worksheet2.Cells.Item(1,7) = "Virtual_Memory_Name";$worksheet2.Cells.Item(1,8) = "Virtual_Memory_CurrentUsage_MB";$worksheet2.Cells.Item(1,9) = "Virtual_Memory_PeakUsage_MB"
$worksheet2.Cells.Item(1,10) = "Virtual_Memory_AllocatedBaseSize_MB"

# Create Heading for Processor worksheet
$worksheet3.Cells.Item(1,1) = "Device_Name";$worksheet3.Cells.Item(1,2) = "Processor(s)";$worksheet3.Cells.Item(1,3) = "Name";$worksheet3.Cells.Item(1,4) = "Type";$worksheet3.Cells.Item(1,5) = "Family";$worksheet3.Cells.Item(1,6) = "Speed_MHz"
$worksheet3.Cells.Item(1,7) = "Cache_Size_MB";$worksheet3.Cells.Item(1,8) = "Interface";$worksheet3.Cells.Item(1,9) = "#_of_Sockets"

# Create Heading for Memory worksheet
$worksheet4.Cells.Item(1,1) = "Device_Name";$worksheet4.Cells.Item(1,2) = "Label";$worksheet4.Cells.Item(1,3) = "Capacity_GB";$worksheet4.Cells.Item(1,4) = "Form"
$worksheet4.Cells.Item(1,5) = "Type"

# Create Heading for Disk worksheet
$worksheet5.Cells.Item(1,1) = "Device_Name";$worksheet5.Cells.Item(1,2) = "Disk_Type";$worksheet5.Cells.Item(1,3) = "Drive_Letter";$worksheet5.Cells.Item(1,4) = "Capacity_GB";$worksheet5.Cells.Item(1,5) = "Free_Space_GB"

# Create Heading for Network worksheet
$worksheet6.Cells.Item(1,1) = "Device_Name";$worksheet6.Cells.Item(1,2) = "Network_Card";$worksheet6.Cells.Item(1,3) = "DHCP_Enabled";$worksheet6.Cells.Item(1,4) = "IP_Address";$worksheet6.Cells.Item(1,5) = "Subnet_Mask"
$worksheet6.Cells.Item(1,6) = "Default_Gateway";$worksheet6.Cells.Item(1,7) = "DNS_Servers";$worksheet6.Cells.Item(1,8) = "DNS_Reg";$worksheet6.Cells.Item(1,9) = "Primary_WINS";$worksheet6.Cells.Item(1,10) = "Secondary_WINS"
$worksheet6.Cells.Item(1,11) = "WINS_Lookup"

# Create Heading for Installed Programs worksheet
$worksheet7.Cells.Item(1,1) = "Device_Name";$worksheet7.Cells.Item(1,2) = "Display_Name";$worksheet7.Cells.Item(1,3) = "Display_Version";$worksheet7.Cells.Item(1,4) = "Install_Location";$worksheet7.Cells.Item(1,5) = "Publisher"

# Create Heading for Share Rights worksheet
$worksheet8.Cells.Item(1,1) = "Device_Name";$worksheet8.Cells.Item(1,2) = "Share_Name";$worksheet8.Cells.Item(1,3) = "Principal";$worksheet8.Cells.Item(1,4) = "Rights";$worksheet8.Cells.Item(1,5) = "AceFlags"
$worksheet8.Cells.Item(1,6) = "AceType"

# Create Heading for Services worksheet
$worksheet9.Cells.Item(1,1) = "Device_Name";$worksheet9.Cells.Item(1,2) = "Display_Name";$worksheet9.Cells.Item(1,3) = "Name";$worksheet9.Cells.Item(1,4) = "Start_Name";$worksheet9.Cells.Item(1,5) = "Start_Mode"
$worksheet9.Cells.Item(1,6) = "Path_Name";$worksheet9.Cells.Item(1,7) = "Description"

# Create Heading for Scheduled Tasks worksheet
$worksheet10.Cells.Item(1,1) = "Device_Name";$worksheet10.Cells.Item(1,2) = "Name";$worksheet10.Cells.Item(1,3) = "RunAs";$worksheet10.Cells.Item(1,4) = "Action";$worksheet10.Cells.Item(1,5) = "Name"
$worksheet10.Cells.Item(1,6) = "NextRunTime";$worksheet10.Cells.Item(1,7) = "LastRunTime";

# Create Heading for Printers worksheet
$worksheet11.Cells.Item(1,1) = "Device_Name";$worksheet11.Cells.Item(1,2) = "Location";$worksheet11.Cells.Item(1,3) = "Name";$worksheet11.Cells.Item(1,4) = "PrinterState";$worksheet11.Cells.Item(1,5) = "PrinterStatus"
$worksheet11.Cells.Item(1,6) = "ShareName";$worksheet11.Cells.Item(1,7) = "SystemName"

# Create Heading for Process worksheet
$worksheet12.Cells.Item(1,1) = "Device_Name";$worksheet12.Cells.Item(1,2) = "Name";$worksheet12.Cells.Item(1,3) = "Path";$worksheet12.Cells.Item(1,4) = "SessionId"

# Create Heading for Process worksheet
$worksheet13.Cells.Item(1,1) = "Device_Name";$worksheet13.Cells.Item(1,2) = "Group";$worksheet13.Cells.Item(1,3) = "User"

# Create Heading for ODBC Configured worksheet
$worksheet14.Cells.Item(1,1) = "Device_Name";$worksheet14.Cells.Item(1,2) = "dsn";$worksheet14.Cells.Item(1,3) = "Server";$worksheet14.Cells.Item(1,4) = "Port";$worksheet14.Cells.Item(1,5) = "DatabaseFile"
$worksheet14.Cells.Item(1,6) = "DatabaseName";$worksheet14.Cells.Item(1,7) = "UID";$worksheet14.Cells.Item(1,8) = "PWD";$worksheet14.Cells.Item(1,9) = "Start";$worksheet14.Cells.Item(1,10) = "LastUser"
$worksheet14.Cells.Item(1,11) = "Database";$worksheet14.Cells.Item(1,12) = "DefaultLibraries";$worksheet14.Cells.Item(1,13) = "DefaultPackage";$worksheet14.Cells.Item(1,14) = "DefaultPkgLibrary"
$worksheet14.Cells.Item(1,15) = "System";$worksheet14.Cells.Item(1,16) = "Driver";$worksheet14.Cells.Item(1,17) = "Description"

# Create Heading for ODBC Installed worksheet
$worksheet15.Cells.Item(1,1) = "Device_Name";$worksheet15.Cells.Item(1,2) = "Driver";$worksheet15.Cells.Item(1,3) = "DriverODBCVer";$worksheet15.Cells.Item(1,4) = "FileExtns";$worksheet15.Cells.Item(1,5) = "Setup"

$colSheets = ($worksheet1, $worksheet2, $worksheet3, $worksheet4, $worksheet5, $worksheet6, $worksheet7, $worksheet8, $worksheet9, $worksheet10, $worksheet11, $worksheet12, $worksheet13, $worksheet14, $worksheet15)
foreach ($colorItem in $colSheets) {
    $rowNumber = 2;$rowNumberCPU = 2;$rowNumberMem = 2;$rowNumberDisk = 2;$rowNumberNet = 2;$rowNumberPI = 2;$rowNumberSR = 2;$rowNumberSrv = 2;$rowNumberSch = 2;$rowNumberPrt = 2
    $rowNumberPrc = 2;$rowNumberLug = 2;$rowNumberODBCC = 2;$rowNumberODBCI = 2
    $workbookColor = $colorItem.UsedRange
    $workbookColor.Interior.ColorIndex = 34
    $workbookColor.Font.ColorIndex = 11
    $workbookColor.Font.Bold = $True
}

# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# Get infos with ouput in Excel
# ________________________________________________________________________
foreach ($strComputer in $colComputers){
    $programsInstalled = ""
    $strComputer = $strComputer.ServerName
    $columnNumber = 2
    # Start inventory                                                                  	                                      
    # Populate General Sheet(1) with information  
    Write-Progress -Activity "Getting general information ($strComputer)" -status "Running..." -id 1                 
    $items = gwmi Win32_ComputerSystem -Comp $strComputer | Select-Object Domain, DomainRole, Manufacturer, Model, SystemType, NumberOfProcessors, TotalPhysicalMemory 
    $systemType = $items.SystemType
    Write-Progress -Activity "Formating the output - general ($strComputer)" -status "Running..." -id 1 
    $columnNumber = FormatOutput $items $worksheet1 $rowNumber $columnNumber $strComputer                 
    Write-Progress -Activity "Getting systems information ($strComputer)" -status "Running..." -id 1
    $items = gwmi Win32_OperatingSystem -Comp $strComputer | Select-Object Caption, csdversion        
    #outputInfosServer $items, $worksheet1, $rowNumber        
    Write-Progress -Activity "Formating the output - system ($strComputer)" -status "Running..." -id 1    
    FormatOutput $items $worksheet1 $rowNumber $columnNumber $strComputer | Out-Null                  
    # Populate Systems Sheet
    $columnNumber = 2        
    #$worksheet2.Cells.Item($rowNumber, 1) = $strComputer
    Write-Progress -Activity "Getting systems information ($strComputer)" -status "Running..." -id 1     
    $items = gwmi Win32_BIOS -Comp $strComputer | Select-Object Name, SMBIOSbiosVersion, SerialNumber
    $columnNumber = FormatOutput $items $worksheet2 $rowNumber $columnNumber $strComputer                     	
    $items = gwmi Win32_TimeZone -Comp $strComputer | Select-Object Caption
    $columnNumber = FormatOutput $items $worksheet2 $rowNumber $columnNumber $strComputer       	      
    $items = gwmi Win32_WmiSetting -Comp $strComputer | Select-Object BuildVersion    
    $columnNumber = FormatOutput $items $worksheet2 $rowNumber $columnNumber $strComputer                	      
    $items = gwmi Win32_PageFileUsage -Comp $strComputer | Select-Object Name, CurrentUsage, PeakUsage, AllocatedBaseSize    
    FormatOutput $items $worksheet2 $rowNumber $columnNumber $strComputer | Out-Null                    		        
    # Populate Processor Sheet	                    
    Write-Progress -Activity "Getting processor information ($strComputer)" -status "Running..." -id 1     
    $items = gwmi Win32_Processor -Comp $strComputer | Select-Object DeviceID, Name, Description, family, currentClockSpeed, l2cacheSize, UpgradeMethod, SocketDesignation
    Write-Progress -Activity "Formating the output - processor ($strComputer)" -status "Running..." -id 1         			
    foreach ($item in $items) {  
        $columnNumber = 2
        FormatOutput $item $worksheet3 $rowNumberCPU $columnNumber $strComputer | Out-Null   
        $rowNumberCPU = $rowNumberCPU + 1 
    }   		        
    # Populate Memory Sheet    
    Write-Progress -Activity "Getting memory information ($strComputer)" -status "Running..." -id 1
    $items = gwmi Win32_PhysicalMemory -Comp $strComputer | Select-Object DeviceLocator, Capacity, FormFactor, TypeDetail
    Write-Progress -Activity "Formating the output - memory ($strComputer)" -status "Running..." -id 1         				
    #$memItems2 = gwmi Win32_PhysicalMemoryArray -Comp $strComputer            	  		    			
	foreach ($item in $items) { 
        $columnNumber = 2             
        FormatOutput $item $worksheet4 $rowNumberMem $columnNumber $strComputer | Out-Null
        $rowNumberMem = $rowNumberMem + 1
	}        	    	                				    			
    # Populate Disk Sheet    
    Write-Progress -Activity "Getting disks information ($strComputer)" -status "Running..." -id 1             
    $items = gwmi Win32_LogicalDisk -Comp $strComputer | Select-Object DriveType, DeviceID, Size, FreeSpace
    Write-Progress -Activity "Formating the output - disk ($strComputer)" -status "Running..." -id 1 
    foreach ($item in $items) {  
        $columnNumber = 2   
        FormatOutput $item $worksheet5 $rowNumberDisk $columnNumber $strComputer | Out-Null                   
        $rowNumberDisk = $rowNumberDisk + 1 
    }
    # Create a graph for the disks
    $worksheet5.Range("A1").Select | Out-Null
    $worksheet5.UsedRange.Columns.AutoFit() |Out-Null             
    $chart=$worksheet5.Shapes.AddChart().Chart
    $chart.chartType=$xlChart::xlBarClustered             
    $chart.HasTitle = $true 
    $chart.ChartTitle.Text = "Disk graph"             
    # Populate Network worksheet    
    Write-Progress -Activity "Getting network information ($strComputer)" -status "Running..." -id 1 
    $items = gwmi Win32_NetworkAdapterConfiguration -Comp $strComputer | Where{$_.IPEnabled -eq "True"} | Select-Object Caption, DHCPEnabled, IPAddress, IPSubnet, DefaultIPGateway, DNSServerSearchOrder, FullDNSRegistrationEnabled, WINSPrimaryServer, WINSSecondaryServer, WINSEnableLMHostsLookup
    Write-Progress -Activity "Formating the output - network ($strComputer)" -status "Running..." -id 1         			
    foreach ($item in $items) {  
        $columnNumber = 2
        FormatOutput $item $worksheet6 $rowNumberNet $columnNumber $strComputer | Out-Null        
        $rowNumberNet = $rowNumberNet + 1
    }                  
    Write-Progress -Activity "Getting programs installed information ($strComputer)" -status "Running..." -id 1       
    # Populate Installed Programs worksheet           
    $arrayprogramsInstalled = listProgramsInstalled "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"        
    $arrayprogramsInstalled2 = listProgramsInstalled "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"            
    $items = $arrayprogramsInstalled + $arrayprogramsInstalled2            
    Write-Progress -Activity "Formating the output - Installed Programs ($strComputer)" -status "Running..." -id 1         			
    foreach ($item in $items) {  
        $columnNumber = 2
        FormatOutput $item $worksheet7 $rowNumberPI $columnNumber $strComputer | Out-Null          
        $rowNumberPI = $rowNumberPI + 1
    }                  
    # Populate Shares worksheet 
    Write-Progress -Activity "Getting shares information ($strComputer)" -status "Running..." -id 1 
    if ($shares = Get-WmiObject Win32_Share -ComputerName $strComputer) {        
        $items = @() 
		$shares | Foreach {$items += Get-NtfsRights $_.Name $_.Path $_.__Server}
	}
	else {$shares = "Failed to get share information from {0}." -f $($_.ToUpper())}            
    Write-Progress -Activity "Formating the output - Shares ($strComputer)" -status "Running..." -id 1
    foreach ($item in $items) {  
        $columnNumber = 2
        FormatOutput $item $worksheet8 $rowNumberSR $columnNumber $strComputer | Out-Null          
        $rowNumberSR = $rowNumberSR + 1
    }                               
    # Populate Services worksheet      
    Write-Progress -Activity "Getting services information ($strComputer)" -status "Running..." -id 1 	
    $items = Get-WmiObject win32_service -Comp $strComputer | Select-Object DisplayName, Name, StartName, StartMode, PathName, Description                 
    Write-Progress -Activity "Formating the output - Services ($strComputer)" -status "Running..." -id 1
    foreach ($item in $items) {  
        $columnNumber = 2
        FormatOutput $item $worksheet9 $rowNumberSrv $columnNumber $strComputer | Out-Null         
        $rowNumberSrv = $rowNumberSrv + 1
    }                      
    # Populate Scheduled Tasks worksheet       
    Write-Progress -Activity "Getting tasks information ($strComputer)" -status "Running..." -id 1     
    $items = @()        
    try { $schedule = new-object -comobject "Schedule.Service" ; $schedule.Connect($strComputer) }
    catch [System.Management.Automation.PSArgumentException] { throw $_ }          
    $items += getTasks("\")
    # Close com
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($schedule) | Out-Null
    Remove-Variable schedule        
    Write-Progress -Activity "Formating the output - Scheduled Tasks ($strComputer)" -status "Running..." -id 1     
    foreach ($item in $items) {  
        $columnNumber = 2
        FormatOutput $item $worksheet10 $rowNumberSch $columnNumber $strComputer | Out-Null        
        $rowNumberSch = $rowNumberSch + 1
    }              
    # Populate Printers worksheet     
    Write-Progress -Activity "Getting printers information ($strComputer)" -status "Running..." -id 1
    $items = gwmi Win32_Printer -Comp $strComputer | Select-Object Location, Name, PrinterState, PrinterStatus, ShareName, SystemName           
    Write-Progress -Activity "Formating the output - Printers ($strComputer)" -status "Running..." -id 1  
    foreach ($item in $items) {  
        $columnNumber = 2
        FormatOutput $item $worksheet11 $rowNumberPrt $columnNumber $strComputer | Out-Null           
        $rowNumberPrt = $rowNumberPrt + 1
    }                 
    # Populate Process worksheet       
    Write-Progress -Activity "Getting process information ($strComputer)" -status "Running..." -id 1     
    $items = gwmi win32_process -ComputerName $strComputer | select-object Name, Path, SessionId 
    Write-Progress -Activity "Formating the output - Process ($strComputer)" -status "Running..." -id 1  
    foreach ($item in $items) {  
        $columnNumber = 2
        FormatOutput $item $worksheet12 $rowNumberPrc $columnNumber $strComputer | Out-Null
        $rowNumberPrc = $rowNumberPrc + 1
    }           
    # Populate Local Users worksheet       
    Write-Progress -Activity "Getting local users information ($strComputer)" -status "Running..." -id 1                 
    $items = getLocalUsersInGroup  
    Write-Progress -Activity "Formating the output - Local Users ($strComputer)" -status "Running..." -id 1   
    foreach ($item in $items) {  
        $columnNumber = 2
        FormatOutput $item $worksheet13 $rowNumberLug $columnNumber $strComputer | Out-Null        
        $rowNumberLug = $rowNumberLug + 1
    }                                                     
    # Populate ODBC Configured worksheet 
    Write-Progress -Activity "Getting ODBC connections Configured ($strComputer)" -status "Running..." -id 1   
    if($systemType -eq "x86-based PC") {
        $odbcConfigured = "SOFTWARE\\odbc\\odbc.ini"
        $odbcDriversInstalled = "SOFTWARE\\odbc\\odbcinst.ini"
    }
    else {
        $odbcConfigured = "SOFTWARE\\wow6432Node\\odbc\\odbc.ini"
        $odbcDriversInstalled = "SOFTWARE\\wow6432Node\\odbc\\odbcinst.ini"
    }     
    Write-Progress -Activity "Formating the output - ODBC connections Configured ($strComputer)" -status "Running..." -id 1 
    $items = listODBCConfigured $odbcConfigured        
    foreach ($item in $items) {  
        $columnNumber = 2
        FormatOutput $item $worksheet14 $rowNumberODBCC $columnNumber $strComputer | Out-Null  
        $rowNumberODBCC = $rowNumberODBCC + 1
    }                   
    # Populate ODBC Drivers Installed worksheet               
    Write-Progress -Activity "Getting ODBC Drivers Installed ($strComputer)" -status "Running..." -id 1 
    $items = listODBCInstalled $odbcDriversInstalled   
    Write-Progress -Activity "Formating the output - ODBC Drivers Installed ($strComputer)" -status "Running..." -id 1 
    foreach ($item in $items) {  
        $columnNumber = 2
        FormatOutput $item $worksheet15 $rowNumberODBCI $columnNumber $strComputer | Out-Null 
        $rowNumberODBCI = $rowNumberODBCI + 1
    }    
    Write-Progress -Activity "Formating the output for ($strComputer) ended" -status "Running..." -id 1                           		
    $rowNumber = $rowNumber + 1;$rowNumberCPU = $rowNumberCPU + 1;$rowNumberMem = $rowNumberMem + 1;$rowNumberDisk = $rowNumberDisk + 1;$rowNumberNet = $rowNumberNet + 1;$rowNumberPI  = $rowNumberPI + 1
    $rowNumberSR = $rowNumberSR + 1;$rowNumberSrv = $rowNumberSrv + 1;$rowNumberSch = $rowNumberSch + 1;$rowNumberPrt = $rowNumberPrt + 1;$rowNumberPrc = $rowNumberPrc + 1;$rowNumberLug = $rowNumberLug + 1
    $rowNumberODBCC = $rowNumberODBCC + 1;$rowNumberODBCI = $rowNumberODBCI + 1                
}

Write-Progress -Activity "Autofit column" -status "Running..." -id 1  
# Auto Fit worksheets
foreach ($colorItem in $colSheets) {
    $workbook = $colorItem.UsedRange													
    $workbook.EntireColumn.AutoFit() | Out-Null
}

Write-Progress -Activity "Getting and formating finished" -status "Ended" -id 1  

Write-Progress -Activity "Opening Excel" -status "Ended" -id 1  

$excelDocument.visible = $True

# ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
# That's all folks !
# ________________________________________________________________________
