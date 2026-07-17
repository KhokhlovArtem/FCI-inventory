#requires -Version 5.1
#requires -Modules FailoverClusters

<#
.SYNOPSIS
    Collects Windows Failover Cluster inventory for SQL Server FCI and Hyper-V roles.
.NOTES
    Requires Windows PowerShell 5.1, FailoverClusters, administrator privileges,
    cluster read access, and WinRM access to owner nodes. The script creates only
    timestamped CSV reports and log files. Remove those generated files to revert.
    Run with -WhatIf before the first production run to preview file system changes.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Cluster = '.',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = (Join-Path -Path $PWD -ChildPath 'reports'),

    [Parameter()]
    [string[]]$RoleName,

    [Parameter()]
    [System.Management.Automation.PSCredential]$SqlCredential,

    [Parameter()]
    [System.Management.Automation.PSCredential]$NodeCredential,

    [Parameter()]
    [switch]$SkipSqlQuery,

    [Parameter()]
    [switch]$EncryptSqlConnection,

    [Parameter()]
    [switch]$TrustServerCertificate,

    [Parameter()]
    [ValidateRange(1, 300)]
    [int]$SqlConnectionTimeoutSeconds = 15
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Logging setup

$scriptName = 'FCI-inventory'
$logDirectory = Join-Path -Path $env:USERPROFILE -ChildPath 'powershell-script-logs'
if (-not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
    if ($PSCmdlet.ShouldProcess($logDirectory, 'Create log directory')) {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }
}
$logFile = $null
if (Test-Path -LiteralPath $logDirectory -PathType Container) {
    $logFile = Join-Path -Path $logDirectory -ChildPath (
        '{0}_{1:yyyy-MM-dd_HH-mm-ss}_operation.log' -f $scriptName, (Get-Date)
    )
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO',
        [AllowNull()][string]$Path = $logFile
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] $Level : $Message"
    Write-Host $logEntry
    if ($Path -and -not $WhatIfPreference) {
        Add-Content -LiteralPath $Path -Value $logEntry -Encoding UTF8
    }
}

trap {
    Write-Log -Message 'The inventory operation failed. Review the secure error channel for details.' -Level 'ERROR'
    exit 1
}

Write-Log -Message 'Inventory started.' -Level 'INFO'
if ($logFile) {
    Write-Log -Message "Log directory: $logDirectory" -Level 'INFO'
}

# Prerequisite validation

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Log -Message 'Administrator privileges are required.' -Level 'ERROR'
    exit 1
}
Write-Log -Message 'Administrator privileges confirmed.' -Level 'SUCCESS'

try {
    Import-Module FailoverClusters -ErrorAction Stop
    Write-Log -Message 'FailoverClusters module loaded.' -Level 'SUCCESS'
}
catch {
    Write-Log -Message 'Failed to load the FailoverClusters module.' -Level 'ERROR'
    exit 1
}

if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
    try {
        if ($PSCmdlet.ShouldProcess($OutputDirectory, 'Create report directory')) {
            New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
            Write-Log -Message "Created report directory: $OutputDirectory" -Level 'INFO'
        }
    }
    catch {
        Write-Log -Message 'Failed to create the report directory.' -Level 'ERROR'
        exit 1
    }
}

# Helper functions

function Get-ClusterPrivateValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Resource,
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $parameter = Get-ClusterParameter -InputObject $Resource -Name $Name -ErrorAction Stop
        if ($null -ne $parameter) {
            return $parameter.Value
        }
    }
    catch {
        return $null
    }

    return $null
}

function Invoke-ClusterNodeCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [Parameter()][object[]]$ArgumentList = @()
    )

    $localNames = @(
        $env:COMPUTERNAME,
        [System.Net.Dns]::GetHostName()
    ) | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() }

    $shortComputerName = ($ComputerName -split '\.')[0].ToLowerInvariant()
    if ($localNames -contains $shortComputerName) {
        return & $ScriptBlock @ArgumentList
    }

    $invokeParameters = @{
        ComputerName = $ComputerName
        ScriptBlock  = $ScriptBlock
        ArgumentList = $ArgumentList
        ErrorAction  = 'Stop'
    }
    if ($NodeCredential) {
        $invokeParameters.Credential = $NodeCredential
    }

    Invoke-Command @invokeParameters
}

function Get-NodePhysicalMemoryGB {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    $scriptBlock = {
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)
    }

    Invoke-ClusterNodeCommand -ComputerName $ComputerName -ScriptBlock $scriptBlock
}

function Get-ClusterDiskInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$DiskResource,
        [Parameter(Mandatory)][string]$OwnerNode
    )

    $diskGuid = Get-ClusterPrivateValue -Resource $DiskResource -Name 'DiskIdGuid'
    $diskSignature = Get-ClusterPrivateValue -Resource $DiskResource -Name 'DiskSignature'
    $diskPath = Get-ClusterPrivateValue -Resource $DiskResource -Name 'DiskPath'

    $scriptBlock = {
        param($ExpectedGuid, $ExpectedSignature, $ExpectedPath)

        $normalizeGuid = {
            param($Value)
            if ([string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
            return ([string]$Value).Trim().Trim('{', '}').ToLowerInvariant()
        }

        $expectedGuidNormalized = & $normalizeGuid $ExpectedGuid
        $allDisks = @(Get-Disk -ErrorAction Stop)
        $disk = $null

        if ($expectedGuidNormalized) {
            $disk = $allDisks | Where-Object {
                (& $normalizeGuid $_.Guid) -eq $expectedGuidNormalized
            } | Select-Object -First 1
        }

        if (($null -eq $disk) -and ($null -ne $ExpectedSignature) -and
            ([uint64]$ExpectedSignature -ne 0)) {
            $disk = $allDisks | Where-Object {
                ($null -ne $_.Signature) -and
                ([uint64]$_.Signature -eq [uint64]$ExpectedSignature)
            } | Select-Object -First 1
        }

        if (($null -eq $disk) -and -not [string]::IsNullOrWhiteSpace([string]$ExpectedPath)) {
            $normalizedPath = ([string]$ExpectedPath).TrimEnd('\').ToLowerInvariant()
            $matchingVolume = Get-Volume -ErrorAction Stop | Where-Object {
                (-not [string]::IsNullOrWhiteSpace([string]$_.Path)) -and
                ($_.Path.TrimEnd('\').ToLowerInvariant() -eq $normalizedPath)
            } | Select-Object -First 1

            if ($matchingVolume) {
                $partition = Get-Partition -Volume $matchingVolume -ErrorAction Stop |
                    Select-Object -First 1
                if ($partition) {
                    $disk = $allDisks | Where-Object Number -eq $partition.DiskNumber |
                        Select-Object -First 1
                }
            }
        }

        if ($null -eq $disk) {
            throw 'Failed to match the cluster resource to a disk.'
        }

        $partitions = @(Get-Partition -DiskNumber $disk.Number -ErrorAction Stop)
        $volumes = foreach ($partition in $partitions) {
            Get-Volume -Partition $partition -ErrorAction Stop
        }
        $volumes = @($volumes | Where-Object { $null -ne $_ })

        $mountPoints = @()
        foreach ($partition in $partitions) {
            $mountPoints += @($partition.AccessPaths | Where-Object {
                -not [string]::IsNullOrWhiteSpace([string]$_)
            })
        }

        $volumeSize = ($volumes | Measure-Object -Property Size -Sum).Sum
        $freeSpace = ($volumes | Measure-Object -Property SizeRemaining -Sum).Sum

        [pscustomobject]@{
            DiskNumber       = $disk.Number
            DiskGuid         = [string]$disk.Guid
            PartitionStyle   = [string]$disk.PartitionStyle
            BusType          = [string]$disk.BusType
            MountPoints      = (@($mountPoints | Sort-Object -Unique) -join ', ')
            FileSystemLabels = (@($volumes.FileSystemLabel | Where-Object { $_ } |
                Sort-Object -Unique) -join ', ')
            FileSystems      = (@($volumes.FileSystem | Where-Object { $_ } |
                Sort-Object -Unique) -join ', ')
            DiskSizeGB       = [math]::Round($disk.Size / 1GB, 2)
            VolumeSizeGB     = if ($null -ne $volumeSize) {
                [math]::Round($volumeSize / 1GB, 2)
            } else { $null }
            FreeSpaceGB      = if ($null -ne $freeSpace) {
                [math]::Round($freeSpace / 1GB, 2)
            } else { $null }
            FreePercent      = if ($volumeSize -gt 0) {
                [math]::Round(($freeSpace / $volumeSize) * 100, 2)
            } else { $null }
        }
    }

    Invoke-ClusterNodeCommand -ComputerName $OwnerNode -ScriptBlock $scriptBlock `
        -ArgumentList @($diskGuid, $diskSignature, $diskPath)
}

function Get-SqlMemoryInfo {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DataSource)

    $result = [ordered]@{
        SqlServerName      = $null
        SqlInstanceName    = $null
        SqlProductVersion  = $null
        MinMemoryMB        = $null
        MaxMemoryMB        = $null
        CurrentSqlMemoryMB = $null
        Note               = $null
    }

    $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $builder['Data Source'] = $DataSource
    $builder['Initial Catalog'] = 'master'
    $builder['Connect Timeout'] = $SqlConnectionTimeoutSeconds
    $builder['Application Name'] = 'FailoverClusterInstanceInventory'
    $builder['Encrypt'] = [bool]$EncryptSqlConnection
    $builder['TrustServerCertificate'] = [bool]$TrustServerCertificate

    if ($SqlCredential) {
        $networkCredential = $SqlCredential.GetNetworkCredential()
        $builder['Integrated Security'] = $false
        $builder['User ID'] = $networkCredential.UserName
        $builder['Password'] = $networkCredential.Password
    }
    else {
        $builder['Integrated Security'] = $true
    }

    $connection = New-Object System.Data.SqlClient.SqlConnection $builder.ConnectionString
    try {
        $connection.Open()

        $command = $connection.CreateCommand()
        $command.CommandTimeout = 30
        $command.CommandText = @'
SELECT
    CONVERT(nvarchar(256), SERVERPROPERTY('ServerName')) AS SqlServerName,
    CONVERT(nvarchar(128), SERVERPROPERTY('InstanceName')) AS SqlInstanceName,
    CONVERT(nvarchar(128), SERVERPROPERTY('ProductVersion')) AS ProductVersion,
    MAX(CASE WHEN name = 'min server memory (MB)' THEN value_in_use END) AS MinMemoryMB,
    MAX(CASE WHEN name = 'max server memory (MB)' THEN value_in_use END) AS MaxMemoryMB
FROM sys.configurations
WHERE name IN ('min server memory (MB)', 'max server memory (MB)');
'@
        $reader = $command.ExecuteReader()
        try {
            if ($reader.Read()) {
                $result.SqlServerName = if ($reader.IsDBNull(0)) { $null } else { $reader.GetString(0) }
                $result.SqlInstanceName = if ($reader.IsDBNull(1)) { 'MSSQLSERVER' } else { $reader.GetString(1) }
                $result.SqlProductVersion = if ($reader.IsDBNull(2)) { $null } else { $reader.GetString(2) }
                $result.MinMemoryMB = if ($reader.IsDBNull(3)) { $null } else { [int64]$reader.GetValue(3) }
                $result.MaxMemoryMB = if ($reader.IsDBNull(4)) { $null } else { [int64]$reader.GetValue(4) }
            }
        }
        finally {
            $reader.Close()
            $command.Dispose()
        }

        try {
            $memoryCommand = $connection.CreateCommand()
            $memoryCommand.CommandTimeout = 30
            $memoryCommand.CommandText = @'
SELECT CONVERT(bigint, physical_memory_in_use_kb / 1024)
FROM sys.dm_os_process_memory;
'@
            $currentMemory = $memoryCommand.ExecuteScalar()
            if (($null -ne $currentMemory) -and ($currentMemory -ne [DBNull]::Value)) {
                $result.CurrentSqlMemoryMB = [int64]$currentMemory
            }
            $memoryCommand.Dispose()
        }
        catch {
            $result.Note = 'Current SQL memory usage is unavailable.'
        }
    }
    finally {
        if ($connection.State -ne [System.Data.ConnectionState]::Closed) {
            $connection.Close()
        }
        $connection.Dispose()
    }

    [pscustomobject]$result
}

function Get-ClusterVmMemoryInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][guid]$VmId
    )

    $scriptBlock = {
        param([guid]$ExpectedVmId)
        Import-Module Hyper-V -ErrorAction Stop
        $vm = Get-VM -Id $ExpectedVmId -ErrorAction Stop

        [pscustomobject]@{
            VmName               = $vm.Name
            DynamicMemoryEnabled = $vm.DynamicMemoryEnabled
            StartupMemoryMB      = [math]::Round($vm.MemoryStartup / 1MB, 0)
            AssignedMemoryMB     = [math]::Round($vm.MemoryAssigned / 1MB, 0)
            MinimumMemoryMB      = [math]::Round($vm.MemoryMinimum / 1MB, 0)
            MaximumMemoryMB      = [math]::Round($vm.MemoryMaximum / 1MB, 0)
        }
    }

    Invoke-ClusterNodeCommand -ComputerName $ComputerName -ScriptBlock $scriptBlock `
        -ArgumentList @($VmId)
}

function Get-ClusterVMVhdInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][guid]$VmId
    )

    $scriptBlock = {
        param([guid]$ExpectedVmId)
        Import-Module Hyper-V -ErrorAction Stop
        $vm = Get-VM -Id $ExpectedVmId -ErrorAction Stop
        $drives = @(Get-VMHardDiskDrive -VM $vm -ErrorAction Stop)
        $results = New-Object System.Collections.Generic.List[object]

        foreach ($drive in $drives) {
            $vhd = $null
            try {
                $vhd = Get-VHD -Path $drive.Path -ErrorAction Stop
            }
            catch {
                $vhd = $null
            }

            $results.Add([pscustomobject]@{
                ControllerType     = [string]$drive.ControllerType
                ControllerNumber   = $drive.ControllerNumber
                ControllerLocation = $drive.ControllerLocation
                Path               = $drive.Path
                VhdFormat          = if ($vhd) { [string]$vhd.VhdFormat } else { $null }
                VhdType            = if ($vhd) { [string]$vhd.VhdType } else { $null }
                VhdSizeGB          = if ($vhd) { [math]::Round($vhd.Size / 1GB, 2) } else { $null }
                VhdFileSizeGB      = if ($vhd) { [math]::Round($vhd.FileSize / 1GB, 2) } else { $null }
            })
        }

        if ($results.Count -eq 0) {
            $results.Add([pscustomobject]@{
                ControllerType     = $null
                ControllerNumber   = $null
                ControllerLocation = $null
                Path               = $null
                VhdFormat          = $null
                VhdType            = $null
                VhdSizeGB          = $null
                VhdFileSizeGB      = $null
            })
        }

        [object[]]$results
    }

    Invoke-ClusterNodeCommand -ComputerName $ComputerName -ScriptBlock $scriptBlock `
        -ArgumentList @($VmId)
}

function Test-RoleNameMatch {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter()][string[]]$Patterns
    )

    if (-not $Patterns -or $Patterns.Count -eq 0) { return $true }
    foreach ($pattern in $Patterns) {
        if ($Name -like $pattern) { return $true }
    }
    return $false
}

function Export-InventoryReport {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
    param(
        [Parameter()]
        [object[]]$Rows,
        [Parameter(Mandatory)][string]$ReportType,
        [Parameter(Mandatory)][string]$OutputDir
    )

    if ($Rows.Count -eq 0) {
        Write-Log -Message "No data available for the $ReportType report. Skipping." -Level 'WARN'
        return
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $fileName = 'FCI-inv-{0}_{1}.csv' -f $ReportType, $timestamp
    $filePath = Join-Path -Path $OutputDir -ChildPath $fileName

    if ($PSCmdlet.ShouldProcess($filePath, 'Create CSV report')) {
        $Rows | Export-Csv -LiteralPath $filePath -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Saved $ReportType report: $filePath (rows: $($Rows.Count))." -Level 'SUCCESS'
    }
}

# Cluster connection

Write-Log -Message "Connecting to cluster '$Cluster'." -Level 'INFO'

$clusterObject = if ($Cluster -eq '.') {
    Get-Cluster -ErrorAction Stop
}
else {
    Get-Cluster -Name $Cluster -ErrorAction Stop
}
$clusterName = [string]$clusterObject.Name
Write-Log -Message "Connected to cluster: $clusterName." -Level 'SUCCESS'

# Role enumeration

$allGroups = @(Get-ClusterGroup -Cluster $clusterName -ErrorAction Stop)
$roles = @($allGroups | Where-Object {
    $isCore = ($_.PSObject.Properties.Name -contains 'IsCoreGroup') -and $_.IsCoreGroup
    $groupType = if ($_.PSObject.Properties.Name -contains 'GroupType') {
        [string]$_.GroupType
    } else { '' }
    $isKnownInfrastructureGroup = ($_.Name -in @('Cluster Group', 'Available Storage')) -or
        ($groupType -in @('Cluster', 'AvailableStorage'))
    (-not $isCore) -and (-not $isKnownInfrastructureGroup) -and
        (Test-RoleNameMatch -Name ([string]$_.Name) -Patterns $RoleName)
})

if ($roles.Count -eq 0) {
    Write-Log -Message 'No cluster roles matched the requested conditions.' -Level 'ERROR'
    exit 1
}

Write-Log -Message "Roles found: $($roles.Count)." -Level 'SUCCESS'

# Main role processing

$nodeMemoryCache = @{}
$reportRows = New-Object System.Collections.Generic.List[object]

foreach ($role in $roles) {
    Write-Log -Message "Processing role '$($role.Name)'." -Level 'INFO'

    $ownerNode = [string]$role.OwnerNode
    $resources = @($role | Get-ClusterResource -ErrorAction Stop)
    $diskResources = @($resources | Where-Object { [string]$_.ResourceType -eq 'Physical Disk' })
    $sqlResource = $resources | Where-Object { [string]$_.ResourceType -eq 'SQL Server' } |
        Select-Object -First 1
    $vmResource = $resources | Where-Object {
        [string]$_.ResourceType -in @('Virtual Machine', 'Virtual Machine Configuration')
    } | Select-Object -First 1

    $instanceType = 'Other'
    $instanceName = [string]$role.Name
    $connectionName = $null
    $memorySource = 'The cluster does not store allocated memory for this role type.'
    $minMemoryMB = $null
    $maxMemoryMB = $null
    $startupMemoryMB = $null
    $assignedMemoryMB = $null
    $dynamicMemoryEnabled = $null
    $currentMemoryMB = $null
    $productVersion = $null
    $roleNotes = New-Object System.Collections.Generic.List[string]

    if (-not $nodeMemoryCache.ContainsKey($ownerNode)) {
        try {
            $nodeMemoryCache[$ownerNode] = Get-NodePhysicalMemoryGB -ComputerName $ownerNode
        }
        catch {
            $nodeMemoryCache[$ownerNode] = $null
            $roleNotes.Add('Owner node physical memory is unavailable.')
            Write-Log -Message "Owner node physical memory is unavailable for '$ownerNode'." -Level 'WARN'
        }
    }
    $ownerNodePhysicalMemoryGB = $nodeMemoryCache[$ownerNode]

    # SQL Server FCI processing

    if ($sqlResource) {
        Write-Log -Message "Role '$($role.Name)' identified as SQL FCI." -Level 'INFO'

        $instanceType = 'SQL FCI'
        $memorySource = 'SQL Server: min/max server memory (MB)'

        $clusterInstanceName = Get-ClusterPrivateValue -Resource $sqlResource -Name 'InstanceName'
        if ([string]::IsNullOrWhiteSpace([string]$clusterInstanceName)) {
            if ([string]$sqlResource.Name -match '\((?<Instance>[^)]+)\)') {
                $clusterInstanceName = $Matches.Instance
            }
            else {
                $clusterInstanceName = 'MSSQLSERVER'
            }
        }
        $instanceName = [string]$clusterInstanceName

        $networkNames = foreach ($networkResource in @($resources | Where-Object {
            [string]$_.ResourceType -eq 'Network Name'
        })) {
            Get-ClusterPrivateValue -Resource $networkResource -Name 'DnsName'
        }
        $virtualServerName = @($networkNames | Where-Object { $_ } | Select-Object -First 1)[0]

        if ($virtualServerName) {
            if ($instanceName -eq 'MSSQLSERVER') {
                $connectionName = [string]$virtualServerName
            }
            else {
                $connectionName = '{0}\{1}' -f $virtualServerName, $instanceName
            }
        }
        else {
            $roleNotes.Add('Network Name DNS value was not found; SQL query skipped.')
        }

        if ((-not $SkipSqlQuery) -and $connectionName) {
            try {
                $sqlInfo = Get-SqlMemoryInfo -DataSource $connectionName
                $minMemoryMB = $sqlInfo.MinMemoryMB
                $maxMemoryMB = $sqlInfo.MaxMemoryMB
                $currentMemoryMB = $sqlInfo.CurrentSqlMemoryMB
                $productVersion = $sqlInfo.SqlProductVersion
                if ($sqlInfo.SqlInstanceName) { $instanceName = $sqlInfo.SqlInstanceName }
                if ($sqlInfo.Note) { $roleNotes.Add($sqlInfo.Note) }
                if ($maxMemoryMB -eq 2147483647) {
                    $roleNotes.Add('Max server memory is unlimited (SQL Server default value).')
                }
            }
            catch {
                $roleNotes.Add("SQL connection to '$connectionName' failed.")
                Write-Log -Message "SQL connection failed for '$connectionName'." -Level 'WARN'
            }
        }
        elseif ($SkipSqlQuery) {
            $roleNotes.Add('SQL query was disabled by SkipSqlQuery.')
        }
    }
    elseif ($vmResource) {
        Write-Log -Message "Role '$($role.Name)' identified as a Hyper-V VM." -Level 'INFO'

        $instanceType = 'Hyper-V VM'
        $memorySource = 'Hyper-V VM configuration'
        $vmIdValue = Get-ClusterPrivateValue -Resource $vmResource -Name 'VmId'

        if ($vmIdValue) {
            try {
                $vmInfo = Get-ClusterVmMemoryInfo -ComputerName $ownerNode -VmId ([guid]$vmIdValue)
                $instanceName = $vmInfo.VmName
                $minMemoryMB = $vmInfo.MinimumMemoryMB
                $maxMemoryMB = $vmInfo.MaximumMemoryMB
                $startupMemoryMB = $vmInfo.StartupMemoryMB
                $assignedMemoryMB = $vmInfo.AssignedMemoryMB
                $dynamicMemoryEnabled = $vmInfo.DynamicMemoryEnabled
                Write-Log -Message "VM '$instanceName' memory information collected." -Level 'SUCCESS'
            }
            catch {
                $roleNotes.Add('VM memory information is unavailable.')
                Write-Log -Message "VM memory information is unavailable for '$($role.Name)'." -Level 'WARN'
            }
        }
        else {
            $roleNotes.Add('VM cluster resource VmId was not found.')
        }
    }

    # Physical disk and VHD processing

    $diskResults = New-Object System.Collections.Generic.List[object]

    # Physical Disk resources from cluster
    $pdEntries = if ($diskResources.Count -gt 0) { $diskResources } else { , [object[]]@($null) }
    foreach ($pdResource in $pdEntries) {
        $pdInfo = $null
        $rowNotes = New-Object System.Collections.Generic.List[string]
        foreach ($note in $roleNotes) { $rowNotes.Add($note) }

        if ($pdResource) {
            try {
                $pdInfo = Get-ClusterDiskInfo -DiskResource $pdResource -OwnerNode $ownerNode
            }
            catch {
                $rowNotes.Add('Physical disk information is unavailable.')
                Write-Log -Message "Physical disk information is unavailable for '$($pdResource.Name)'." -Level 'WARN'
            }
        }

        $diskResults.Add([pscustomobject]@{
            DiskType   = 'PhysicalDisk'
            PdInfo     = $pdInfo
            PdResource = $pdResource
            VhdInfo    = $null
            Notes      = $rowNotes
        })
    }

    # VHD info for VM roles
    if ($vmResource -and (-not [string]::IsNullOrWhiteSpace([string]$vmIdValue))) {
        try {
            $vmVhdList = Get-ClusterVMVhdInfo -ComputerName $ownerNode -VmId ([guid]$vmIdValue)
            foreach ($vhd in $vmVhdList) {
                $rowNotes = New-Object System.Collections.Generic.List[string]
                foreach ($note in $roleNotes) { $rowNotes.Add($note) }
                $diskResults.Add([pscustomobject]@{
                    DiskType   = 'VHD'
                    PdInfo     = $null
                    PdResource = $null
                    VhdInfo    = $vhd
                    Notes      = $rowNotes
                })
            }
        }
        catch {
            $roleNotes.Add('VHD information is unavailable.')
            Write-Log -Message "VHD information is unavailable for VM '$($role.Name)'." -Level 'WARN'
        }
    }

    $successfulPdResults = @($diskResults | Where-Object { $null -ne $_.PdInfo })
    $roleTotalDiskSizeGB = if ($successfulPdResults.Count -gt 0) {
        [math]::Round(($successfulPdResults.PdInfo |
            Measure-Object -Property DiskSizeGB -Sum).Sum, 2)
    } else { $null }
    $roleTotalFreeSpaceGB = if ($successfulPdResults.Count -gt 0) {
        [math]::Round(($successfulPdResults.PdInfo |
            Measure-Object -Property FreeSpaceGB -Sum).Sum, 2)
    } else { $null }
    $roleDiskDataComplete = ($diskResources.Count -eq $successfulPdResults.Count)

    foreach ($diskResult in $diskResults) {
        $pdInfo = $diskResult.PdInfo
        $pdResource = $diskResult.PdResource
        $vhdInfo = $diskResult.VhdInfo
        $rowNotes = $diskResult.Notes

        $reportRows.Add([pscustomobject][ordered]@{
            CollectedAt               = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            ClusterName               = $clusterName
            RoleName                  = [string]$role.Name
            RoleState                 = [string]$role.State
            OwnerNode                 = $ownerNode
            OwnerNodePhysicalMemoryGB = $ownerNodePhysicalMemoryGB
            InstanceType              = $instanceType
            InstanceName              = $instanceName
            ConnectionName            = $connectionName
            ProductVersion            = $productVersion
            MemorySource              = $memorySource
            MinMemoryMB               = $minMemoryMB
            MaxMemoryMB               = $maxMemoryMB
            StartupMemoryMB           = $startupMemoryMB
            AssignedMemoryMB          = $assignedMemoryMB
            CurrentSqlMemoryMB        = $currentMemoryMB
            DynamicMemoryEnabled      = $dynamicMemoryEnabled
            DiskResourceCount         = $diskResources.Count
            RoleTotalDiskSizeGB       = $roleTotalDiskSizeGB
            RoleTotalFreeSpaceGB      = $roleTotalFreeSpaceGB
            RoleDiskDataComplete      = $roleDiskDataComplete
            DiskResourceName          = if ($pdResource) { [string]$pdResource.Name } else { $null }
            DiskResourceState         = if ($pdResource) { [string]$pdResource.State } else { $null }
            DiskNumber                = if ($pdInfo)  { $pdInfo.DiskNumber } else { $null }
            DiskGuid                  = if ($pdInfo)  { $pdInfo.DiskGuid   } else { $null }
            PartitionStyle            = if ($pdInfo)  { $pdInfo.PartitionStyle } else { $null }
            BusType                   = if ($pdInfo)  { $pdInfo.BusType    } else { $null }
            MountPoints               = if ($pdInfo)  { $pdInfo.MountPoints  } else { $null }
            FileSystemLabels          = if ($pdInfo)  { $pdInfo.FileSystemLabels } else { $null }
            FileSystems               = if ($pdInfo)  { $pdInfo.FileSystems    } else { $null }
            DiskSizeGB                = if ($pdInfo)  { $pdInfo.DiskSizeGB  } else { $null }
            VolumeSizeGB              = if ($pdInfo)  { $pdInfo.VolumeSizeGB } else { $null }
            FreeSpaceGB               = if ($pdInfo)  { $pdInfo.FreeSpaceGB } else { $null }
            FreePercent               = if ($pdInfo)  { $pdInfo.FreePercent } else { $null }
            VhdPath                   = if ($vhdInfo) { $vhdInfo.Path      } else { $null }
            VhdFormat                 = if ($vhdInfo) { $vhdInfo.VhdFormat } else { $null }
            VhdType                   = if ($vhdInfo) { $vhdInfo.VhdType   } else { $null }
            VhdSizeGB                 = if ($vhdInfo) { $vhdInfo.VhdSizeGB } else { $null }
            VhdFileSizeGB             = if ($vhdInfo) { $vhdInfo.VhdFileSizeGB } else { $null }
            VhdControllerType         = if ($vhdInfo) { $vhdInfo.ControllerType   } else { $null }
            VhdControllerNumber       = if ($vhdInfo) { $vhdInfo.ControllerNumber } else { $null }
            VhdControllerLocation     = if ($vhdInfo) { $vhdInfo.ControllerLocation } else { $null }
            Notes                     = ($rowNotes -join ' | ')
        })
    }
}

# Report export

Write-Log -Message 'Creating reports.' -Level 'INFO'

$sqlRows = @($reportRows | Where-Object { $_.InstanceType -eq 'SQL FCI' })
$vmRows = @($reportRows | Where-Object { $_.InstanceType -eq 'Hyper-V VM' })
$otherRows = @($reportRows | Where-Object { $_.InstanceType -eq 'Other' })

Write-Log -Message "SQL FCI roles: $($sqlRows.Count)." -Level 'INFO'
Write-Log -Message "Hyper-V VM roles: $($vmRows.Count)." -Level 'INFO'
Write-Log -Message "Other roles: $($otherRows.Count)." -Level 'INFO'

Export-InventoryReport -Rows $sqlRows -ReportType 'SQL' -OutputDir $OutputDirectory
Export-InventoryReport -Rows $vmRows -ReportType 'VM' -OutputDir $OutputDirectory

# Completion

Write-Log -Message 'Inventory completed successfully.' -Level 'SUCCESS'
if ($logFile) {
    Write-Log -Message "Log file: $logFile" -Level 'INFO'
}
