# SQL License analysis script
# Joshua Woleben
# Written 3/27/19

# Set up Excel parameters
$excel_file = "C:\Temp\SQL_Report.xlsx"
$sheet1_name = "Core and License Summary"
$sheet2_name = "Comprehensive Server List"
$sheet3_name = "Unresponsive Servers"
$sheet4_name = "Servers without an SQL License"

# $ErrorActionPreference = "SilentlyContinue"

# Load list of SQL servers to check
$sql_server_list = (Get-Content -Path "C:\Temp\sql_server_list.txt")

# Set up initial variables
$total_cores_express = 0
$total_cores_enterprise = 0
$total_cores_standard = 0
$version_list = @{}
$core_count_list = @{}
$license_list = @{}
$dead_servers = @()
$no_license_servers = @()
$version_count = @{}

# Check each server
ForEach ($server in $sql_server_list) {

# Reset all loop variables
    $instance = ""
    $license = ""
    $core_count = 0
    $license_object = ""
    $version_object = ""
    $core_count_object = ""

# Check server for response to ping, if none, record as dead server
if (Test-Connection -ComputerName $server -ErrorAction SilentlyContinue) {

#    $jobs += Start-Job -NoNewScope -ScriptBlock {
    Write-Output "Checking server $server ..."

    # Get processor count
    $core_count_object = (Get-WmiObject -Class Win32_processor -ComputerName $server)

    # If core count property exists, record it, otherwise count the number of CPUs returned
    if (Get-Member -InputObject $core_count_object -name "NumberOfCores" -MemberType Properties) {
        $core_count_number = ($core_count_object | Select -ExpandProperty NumberOfCores)
        if ($core_count_number.GetType() -eq [System.UInt32]) {
            $core_count = $core_count_number
        }
        else {
           $core_count = $core_count_number.GetValue(0)
        }
    }
    else {
        $core_count = $core_count_object.Count
    }
    Write-Verbose "$core_count cores found!"


    # Determine WMI namespace for current computer

    $comp_mgmt_nsp = (Get-WmiObject -ComputerName $server -Namespace "root\microsoft\sqlserver" -Class __NAMESPACE -ErrorAction SilentlyContinue |
        Where-Object {$_.Name -like "ComputerManagement*"} |
        Select-Object Name |
        Sort-Object Name -Descending |
        Select-Object -First 1).Name

        Write-Verbose "Computer management namespace: $comp_mgmt_nsp"
        $comp_mgmt_nsp = "root\microsoft\sqlserver\" + $comp_mgmt_nsp

        # Get SQL Server license type using WMI
        $license_object = Get-WmiObject -ComputerName $server -Namespace $comp_mgmt_nsp -Class "SqlServiceAdvancedProperty" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.ServiceName -like "MSSQL*" -and
                $_.PropertyName -eq "SKUNAME"
            } |
            Select-Object @{Name = "ComputerName"; Expression = { $server }},
            @{Name = "PropertyValue"; Expression = {
                 $_.PropertyStrValue}  
            } 
            if (($license_object | Get-Unique).Count -gt 1) {
                $license_object = ($license_object | Where-Object { $_.PropertyValue -match "Standard" -or $_.PropertyValue -match "Enterprise" }) 
            }
    
        # Get SQL version using WMI        
        $version_object = Get-WmiObject -ComputerName $server -Namespace $comp_mgmt_nsp -Class "SqlServiceAdvancedProperty" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.ServiceName -like "MSSQL*" -and
                $_.PropertyName -eq "VERSION"
            } -ErrorAction SilentlyContinue |
            Select-Object @{Name = "ComputerName"; Expression = { $server }},
            @{Name = "SQLVersion"; Expression = {
                 $_.PropertyStrValue}      
            } | Select-Object -First 1

    
    # Check to see if version exists, if so, set it properly
    if (Get-Member -InputObject $version_object -name "SQLVersion" -MemberType Properties) {
         $version = ($version_object | Select -ExpandProperty SQLVersion).ToString()
    }
    else {
        $version = "None"
    }

    # Record server version
    $version_list[$server] = $version

    # Record number of servers on version
    if ($version_count[$version] -eq $null) {
        $version_count[$version] = 1
    }
    else {
        $version_count[$version]++
    }

    # Verbose output
    Write-Verbose "SQL Server Version: $version_object"
    Write-Verbose "License found: $license_object"

    # Set license type
    $license = ( $license_object | Select-Object -First 1 | Select -ExpandProperty PropertyValue).ToString()
    
    Write-Verbose "License Value: $license"
    
    # Record server license type
    $license_list[$server] = $license

    # Record core count
    $core_count_list[$server] = $core_count

    # Update core count for current license type
    if ($license -match "Express") {
         $total_cores_express += $core_count
    }
    elseif ($license -match "Standard") {
        $total_cores_standard += $core_count
    }
    elseif ($license -match "Enterprise") {
        $total_cores_enterprise += $core_count
    }
    else {
        $no_license_servers += $server
    }
 #   } # end start-job
}
    else {
        Write-Output "$server not responding, skipping..."
        $dead_servers += $server
    }
    Wait-Job -Job $jobs -Timeout 15
}

# Loop is done, now create excel spreadsheet
$excel_object = New-Object -ComObject Excel.Application

# Turn off visibility
$excel_object.Visible = $false

# Open Excel file, create workbook
$excel_workbook = $excel_object.Workbooks.Add()


# Create the worksheets
$worksheet1 = $excel_workbook.Worksheets.Item(1)
$worksheet2 = $excel_workbook.Worksheets.Add()
$worksheet3 = $excel_workbook.Worksheets.Add()
$worksheet4 = $excel_workbook.Worksheets.Add()

$worksheet1.Name = $sheet1_name
$worksheet2.Name = $sheet2_name
$worksheet3.Name = $sheet3_name
$worksheet4.Name = $sheet4_name


# Generate summary page
$worksheet1.Cells.Item(1,1) = "License Type"
$worksheet1.Cells.Item(1,2) = "Core Count"
$worksheet1.Cells.Item(1,1).Font.Size = 14
$worksheet1.Cells.Item(1,1).Font.Bold = $true
$worksheet1.Cells.Item(1,2).Font.Size = 14
$worksheet1.Cells.Item(1,2).Font.Bold = $true
$worksheet1.Cells.Item(2,1) = "Express"
$worksheet1.Cells.Item(3,1) = "Standard"
$worksheet1.Cells.Item(4,1) = "Enterprise"
$worksheet1.Cells.Item(2,2) = $total_cores_express
$worksheet1.Cells.Item(3,2) = $total_cores_standard
$worksheet1.Cells.Item(4,2) = $total_cores_enterprise

$worksheet1.Cells.Item(1,4) = "Version Name"
$worksheet1.Cells.Item(1,5) = "Count"
$worksheet1.Cells.Item(1,4).Font.Size = 14
$worksheet1.Cells.Item(1,4).Font.Bold = $true
$worksheet1.Cells.Item(1,5).Font.Size = 14
$worksheet1.Cells.Item(1,5).Font.Bold = $true

$row = 2
ForEach ($key in $version_count.Keys) {
    $version_name = $key
    $count = $version_count[$key]
    $worksheet1.Cells.Item($row,4) = $version_name
    $worksheet1.Cells.Item($row,5) = $count
    $row ++
}
$w1_count = $row

# Generate Comprehensive Server List
$worksheet2.Cells.Item(1,1) = "Server Name"
$worksheet2.Cells.Item(1,2) = "Core Count"
$worksheet2.Cells.Item(1,3) = "License Type"
$worksheet2.Cells.Item(1,4) = "SQL Server Version"
for ($i = 1; $i -lt 5; $i += 1) {
    $worksheet2.Cells.Item(1,$i).Font.Size = 14
    $worksheet2.Cells.Item(1,$i).Font.Bold = $true
} 

$row = 2
ForEach ($key in $version_list.Keys) {
    $version = $version_list[$key]
    $cores = $core_count_list[$key]
    $server_name = $key
    $license_name = $license_list[$key]

    $worksheet2.Cells.Item($row,1) = $server_name
    $worksheet2.Cells.Item($row,2) = $cores
    $worksheet2.Cells.Item($row,3) = $license_name
    $worksheet2.Cells.Item($row,4) = $version
    $row ++
}
$w2_count = $row
# Generate Dead Server List
$worksheet3.Cells.Item(1,1) = "Unresponsive server hostnames"
$worksheet3.Cells.Item(1,1).Font.Size = 14
$worksheet3.Cells.Item(1,1).Font.Bold = $true

$row = 2
ForEach ($name in $dead_servers) {
    $worksheet3.Cells.Item($row,1) = $name
    $row ++
}
$w3_count = $row

# Generate SQL License list of servers without one
$worksheet4.Cells.Item(1,1) = "Hostnames with no SQL License Information"
$worksheet4.Cells.Item(1,1).Font.Size = 14
$worksheet4.Cells.Item(1,1).Font.Bold = $true

$row = 2
ForEach ($name in $no_license_servers) {
    $worksheet4.Cells.Item($row,1) = $name
    $row ++
}
$w4_count = $row

# Format the worksheets
$range_1 = $worksheet1.Range("A1:E$w1_count")
$range_1.EntireColumn.AutoFit()
$range_2 = $worksheet2.Range("A2:D$w2_count")
$range_2.Sort($range_2, 1)
$range_2.EntireColumn.AutoFit()
$range_3 = $worksheet3.Range("A2:A$w3_count")
$range_3.Sort($range_3, 1)
$range_3.EntireColumn.AutoFit()
$range_4 = $worksheet4.Range("A2:A$w4_count")
$range_4.Sort($range_4,1)
$range_4.EntireColumn.AutoFit()


# Save Excel File
$excel_object.DisplayAlerts = 'False'
$excel_workbook.SaveAs($excel_file)
$excel_workbook.Close
$excel_object.DiplayAlerts = 'False'
$excel_object.Quit()

Write-Output "Main Report: "
Write-Output "Server Name, Core Count, License Type, SQL Server Version"
ForEach ($key in $version_list.Keys) {
    $version = $version_list[$key]
    $cores = $core_count_list[$key]
    $license_name = $license_list[$key]
    Write-Output ($key + ", " + $cores + ", " + $license_name + ", " + $version)
}
Write-Output "`n`n`n"
Write-Output "Servers that did not respond: "
Write-Output $dead_servers

Write-Output "`n`n"
Write-Output "Servers without an SQL Server License: "
Write-Output $no_license_servers

Write-Output "`n`n`n"
Write-Output "Total Express cores: $total_cores_express"
Write-Output "Total Standard cores: $total_cores_standard"
Write-Output "Total Enterprise cores: $total_cores_enterprise"

