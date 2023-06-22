# Define the time period for the report
$startTime = Get-Date "2023-06-01 00:00:00"  # Replace with your desired start time
$endTime = Get-Date "2023-06-30 23:59:59"    # Replace with your desired end time

# Function to calculate peak and average utilization
function CalculateUtilization([System.Collections.ArrayList]$utilizationList) {
    $peakUtilization = $utilizationList | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
    $averageUtilization = $utilizationList | Measure-Object -Average | Select-Object -ExpandProperty Average

    [PSCustomObject]@{
        PeakUtilization = $peakUtilization
        AverageUtilization = $averageUtilization
    }
}

# Function to format percentage values
function FormatPercentage($value) {
    "{0:N2}%" -f $value
}

# Get CPU utilization
try {
    $cpuUtilizationList = Get-WmiObject -Class Win32_PerfFormattedData_PerfOS_Processor -ErrorAction Stop |
        Where-Object { $_.Name -notlike "_Total" } |
        ForEach-Object { $_.PercentProcessorTime }

    $cpuReport = CalculateUtilization $cpuUtilizationList
}
catch {
    Write-Host "Failed to retrieve CPU utilization. Error: $_"
    exit 1
}

# Get memory (RAM) utilization
try {
    $memoryUtilizationList = Get-WmiObject -Class Win32_PerfFormattedData_PerfOS_Memory -ErrorAction Stop |
        ForEach-Object { $_.PercentCommittedBytesInUse }

    $memoryReport = CalculateUtilization $memoryUtilizationList
}
catch {
    Write-Host "Failed to retrieve memory utilization. Error: $_"
    exit 1
}

# Get disk utilization
try {
    $diskUtilizationList = Get-WmiObject -Class Win32_PerfFormattedData_PerfDisk_LogicalDisk -ErrorAction Stop |
        Where-Object { $_.Name -notlike "_Total" } |
        ForEach-Object { $_.PercentDiskTime }

    $diskReport = CalculateUtilization $diskUtilizationList
}
catch {
    Write-Host "Failed to retrieve disk utilization. Error: $_"
    exit 1
}

# Display the reports
Write-Host "===== CPU Utilization Report ====="
Write-Host "Peak Utilization: $(FormatPercentage $cpuReport.PeakUtilization)"
Write-Host "Average Utilization: $(FormatPercentage $cpuReport.AverageUtilization)"
Write-Host

Write-Host "===== Memory Utilization Report ====="
Write-Host "Peak Utilization: $(FormatPercentage $memoryReport.PeakUtilization)"
Write-Host "Average Utilization: $(FormatPercentage $memoryReport.AverageUtilization)"
Write-Host

Write-Host "===== Disk Utilization Report ====="
Write-Host

try {
    $disks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop

    foreach ($disk in $disks) {
        Write-Host "Drive: $($disk.DeviceID)"
        
        $diskUtilizationList = Get-WmiObject -Class Win32_PerfFormattedData_PerfDisk_LogicalDisk -Filter "Name='$($disk.DeviceID)'" -ErrorAction Stop |
            ForEach-Object { $_.PercentDiskTime }

        $diskReport = CalculateUtilization $diskUtilizationList

        Write-Host "Peak Utilization: $(FormatPercentage $diskReport.PeakUtilization)"
        Write-Host "Average Utilization: $(FormatPercentage $diskReport.AverageUtilization)"
        Write-Host
    }
}
catch {
    Write-Host "Failed to retrieve disk information. Error: $_"
    exit 1
}
