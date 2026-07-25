#requires -Version 5.1

# ==============================================================================
# SCRIPT NAME: MSSQL-inventory.ps1
#
# Purpose:
#   - Inventories database files of all service-hosted local SQL Server instances.
#   - Exports every visible user and system database file to a CSV report.
#   - Adds CSV rows for stopped or unavailable SQL Server instances.
#
# Requirements:
#   - Windows PowerShell 5.1.
#   - Permission to connect to every local SQL Server instance.
#   - VIEW ANY DATABASE for a complete database catalog inventory.
#   - Windows authentication is used by default.
#
# Optional SQL authentication:
#   - Set MSSQL_INVENTORY_SQL_USER and MSSQL_INVENTORY_SQL_PASSWORD in the
#     process environment before launching the script.
#   - The script does not read .env files directly.
# ==============================================================================

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = 'reports',

    [Parameter()]
    [ValidateRange(1, 300)]
    [int]$SqlConnectionTimeoutSeconds = 15,

    [Parameter()]
    [switch]$EncryptSqlConnection,

    [Parameter()]
    [switch]$TrustServerCertificate
)

$ErrorActionPreference = 'Stop'
$scriptRoot = $PSScriptRoot

if (-not [System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory = Join-Path -Path $scriptRoot -ChildPath $OutputDirectory
}

# ==================================================================
# LOGGING
# ==================================================================

$scriptName = 'MSSQL-inventory'
$logDirectory = Join-Path -Path $scriptRoot -ChildPath 'logs'
if (-not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
}
if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
    New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}

$runId = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$logFile = Join-Path -Path $logDirectory -ChildPath (
    '{0}_{1}.log' -f $scriptName, $runId
)

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] $Level : $Message"
    Write-Host $logEntry
    Add-Content -LiteralPath $logFile -Value $logEntry -Encoding UTF8
}

function Get-SafeErrorDetail {
    param([Parameter(Mandatory)][System.Exception]$Exception)

    $message = $Exception.Message -replace '(?i)(password|pwd)\s*=\s*[^;\s]+', '$1=***'
    if ($Exception -is [System.Data.SqlClient.SqlException]) {
        return ('SqlException Number={0}, State={1}, Class={2}: {3}' -f $Exception.Number, $Exception.State, $Exception.Class, $message)
    }

    return '{0}: {1}' -f $Exception.GetType().Name, $message
}

function ConvertTo-CsvSafeValue {
    param([Parameter()]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ($text -match '^[=+\-@\t\r]') {
        return "'$text"
    }

    return $text
}

# ==================================================================
# CONFIGURATION
# ==================================================================

$sqlUser = [Environment]::GetEnvironmentVariable('MSSQL_INVENTORY_SQL_USER', 'Process')
$sqlPassword = [Environment]::GetEnvironmentVariable('MSSQL_INVENTORY_SQL_PASSWORD', 'Process')
$hasSqlUser = -not [string]::IsNullOrWhiteSpace($sqlUser)
$hasSqlPassword = -not [string]::IsNullOrWhiteSpace($sqlPassword)

if ($hasSqlUser -xor $hasSqlPassword) {
    Write-Log -Message 'Both SQL authentication environment variables must be set together.' -Level 'ERROR'
    exit 1
}

$useSqlAuthentication = $hasSqlUser -and $hasSqlPassword
Write-Log -Message 'Script started.' -Level 'INFO'
Write-Log -Message "Script root: $scriptRoot" -Level 'INFO'
Write-Log -Message "Output directory: $OutputDirectory" -Level 'INFO'
Write-Log -Message ('Authentication: ' + $(if ($useSqlAuthentication) { 'SQL authentication from environment' } else { 'Windows authentication' })) -Level 'INFO'

# ==================================================================
# SQL HELPERS
# ==================================================================

function New-SqlConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DataSource,
        [Parameter(Mandatory)][string]$InitialCatalog
    )

    $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $builder['Data Source'] = $DataSource
    $builder['Initial Catalog'] = $InitialCatalog
    $builder['Connect Timeout'] = $SqlConnectionTimeoutSeconds
    $builder['Application Name'] = 'MSSQLDatabaseFileInventory'
    $builder['Encrypt'] = [bool]$EncryptSqlConnection
    $builder['TrustServerCertificate'] = [bool]$TrustServerCertificate

    if ($useSqlAuthentication) {
        $builder['Integrated Security'] = $false
        $builder['User ID'] = $sqlUser
        $builder['Password'] = $sqlPassword
    }
    else {
        $builder['Integrated Security'] = $true
    }

    return [System.Data.SqlClient.SqlConnection]::new($builder.ConnectionString)
}

function Invoke-SqlQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DataSource,
        [Parameter(Mandatory)][string]$InitialCatalog,
        [Parameter(Mandatory)][string]$Query
    )

    $connection = $null
    $command = $null
    $reader = $null
    $rows = New-Object System.Collections.Generic.List[object]

    try {
        $connection = New-SqlConnection -DataSource $DataSource -InitialCatalog $InitialCatalog
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandTimeout = 60
        $command.CommandText = $Query
        $reader = $command.ExecuteReader()

        while ($reader.Read()) {
            $row = [ordered]@{}
            for ($column = 0; $column -lt $reader.FieldCount; $column++) {
                $row[$reader.GetName($column)] = if ($reader.IsDBNull($column)) {
                    $null
                }
                else {
                    $reader.GetValue($column)
                }
            }
            $rows.Add([pscustomobject]$row)
        }
    }
    finally {
        if ($reader) { $reader.Dispose() }
        if ($command) { $command.Dispose() }
        if ($connection) { $connection.Dispose() }
    }

    return $rows.ToArray()
}

# ==================================================================
# REPORT HELPERS
# ==================================================================

$csvColumns = @(
    'CollectedAt', 'ComputerName', 'ConnectionTarget', 'SqlServerName',
    'SqlMachineName', 'SqlInstanceName', 'SqlIsClustered', 'ServiceName',
    'ServiceStatus', 'InstanceAvailability', 'InventoryComplete', 'DatabaseId',
    'DatabaseName', 'DatabaseState', 'LogicalFileName', 'FileType', 'FilePath',
    'FileAllocatedSizeMB', 'FileAllocatedSizeGB', 'FileUsedSpaceMB',
    'FileUsedSpaceGB', 'UsedSpaceScope', 'DatabaseLogUsedSpaceMB',
    'DatabaseLogUsedSpaceGB', 'DatabaseLogUsedSpaceScope', 'Notes'
)

function Add-InventoryRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Rows,
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter()]$DatabaseFile,
        [Parameter()]$LogUsage,
        [Parameter(Mandatory)][string]$InventoryComplete,
        [Parameter()][string]$Notes
    )

    $fileAllocatedSizeMB = if ($DatabaseFile -and $null -ne $DatabaseFile.FileAllocatedSizeMB) {
        [math]::Round([double]$DatabaseFile.FileAllocatedSizeMB, 2)
    }
    else { $null }
    $fileUsedSpaceMB = if ($DatabaseFile -and $null -ne $DatabaseFile.FileUsedSpaceMB) {
        [math]::Round([double]$DatabaseFile.FileUsedSpaceMB, 2)
    }
    else { $null }
    $databaseLogUsedSpaceMB = if ($LogUsage -and $null -ne $LogUsage.DatabaseLogUsedSpaceMB) {
        [math]::Round([double]$LogUsage.DatabaseLogUsedSpaceMB, 2)
    }
    else { $null }

    $fileType = if ($DatabaseFile) { [string]$DatabaseFile.FileType } else { $null }
    $usedSpaceScope = if ($fileType -eq 'ROWS') { 'File' } else { 'Unavailable' }
    $databaseLogUsedSpaceScope = if ($fileType -eq 'LOG' -and $null -ne $databaseLogUsedSpaceMB) {
        'DatabaseAggregate'
    }
    elseif ($fileType -eq 'LOG') {
        'Unavailable'
    }
    else {
        $null
    }

    $Rows.Add([pscustomobject][ordered]@{
        CollectedAt                 = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        ComputerName                = ConvertTo-CsvSafeValue $Context.ComputerName
        ConnectionTarget            = ConvertTo-CsvSafeValue $Context.ConnectionTarget
        SqlServerName               = ConvertTo-CsvSafeValue $Context.SqlServerName
        SqlMachineName              = ConvertTo-CsvSafeValue $Context.SqlMachineName
        SqlInstanceName             = ConvertTo-CsvSafeValue $Context.SqlInstanceName
        SqlIsClustered              = $Context.SqlIsClustered
        ServiceName                 = ConvertTo-CsvSafeValue $Context.ServiceName
        ServiceStatus               = ConvertTo-CsvSafeValue $Context.ServiceStatus
        InstanceAvailability        = $Context.InstanceAvailability
        InventoryComplete           = $InventoryComplete
        DatabaseId                  = if ($DatabaseFile) { $DatabaseFile.DatabaseId } else { $null }
        DatabaseName                = if ($DatabaseFile) { ConvertTo-CsvSafeValue $DatabaseFile.DatabaseName } else { $null }
        DatabaseState               = if ($DatabaseFile) { ConvertTo-CsvSafeValue $DatabaseFile.DatabaseState } else { $null }
        LogicalFileName             = if ($DatabaseFile) { ConvertTo-CsvSafeValue $DatabaseFile.LogicalFileName } else { $null }
        FileType                    = $fileType
        FilePath                    = if ($DatabaseFile) { ConvertTo-CsvSafeValue $DatabaseFile.FilePath } else { $null }
        FileAllocatedSizeMB         = $fileAllocatedSizeMB
        FileAllocatedSizeGB         = if ($null -ne $fileAllocatedSizeMB) { [math]::Round($fileAllocatedSizeMB / 1024, 2) } else { $null }
        FileUsedSpaceMB             = $fileUsedSpaceMB
        FileUsedSpaceGB             = if ($null -ne $fileUsedSpaceMB) { [math]::Round($fileUsedSpaceMB / 1024, 2) } else { $null }
        UsedSpaceScope              = $usedSpaceScope
        DatabaseLogUsedSpaceMB      = $databaseLogUsedSpaceMB
        DatabaseLogUsedSpaceGB      = if ($null -ne $databaseLogUsedSpaceMB) { [math]::Round($databaseLogUsedSpaceMB / 1024, 2) } else { $null }
        DatabaseLogUsedSpaceScope   = $databaseLogUsedSpaceScope
        Notes                       = ConvertTo-CsvSafeValue $Notes
    })
}

# ==================================================================
# LOCAL SQL SERVER INSTANCE DISCOVERY
# ==================================================================

$computerName = $env:COMPUTERNAME
$instanceServices = @(Get-Service -ErrorAction Stop | Where-Object {
    $_.Name -eq 'MSSQLSERVER' -or $_.Name -like 'MSSQL$*'
})
$reportRows = New-Object System.Collections.Generic.List[object]

if ($instanceServices.Count -eq 0) {
    Write-Log -Message 'No local service-hosted SQL Server instances were found.' -Level 'WARN'
    $noInstanceContext = @{
        ComputerName = $computerName; ConnectionTarget = $null; SqlServerName = $null
        SqlMachineName = $null; SqlInstanceName = $null; SqlIsClustered = $null
        ServiceName = $null; ServiceStatus = $null; InstanceAvailability = 'Unavailable'
    }
    Add-InventoryRow -Rows $reportRows -Context $noInstanceContext -InventoryComplete 'False' `
        -Notes 'No local MSSQLSERVER or MSSQL$ instance service was found.'
}
else {
    Write-Log -Message "Local SQL Server services found: $($instanceServices.Count)." -Level 'SUCCESS'
}

# ==================================================================
# SQL SERVER INVENTORY
# ==================================================================

$identityQuery = @'
SELECT
    CONVERT(nvarchar(256), SERVERPROPERTY('ServerName')) AS SqlServerName,
    CONVERT(nvarchar(256), SERVERPROPERTY('MachineName')) AS SqlMachineName,
    CONVERT(nvarchar(128), SERVERPROPERTY('InstanceName')) AS SqlInstanceName,
    CONVERT(int, SERVERPROPERTY('IsClustered')) AS SqlIsClustered,
    CONVERT(int, HAS_PERMS_BY_NAME(NULL, NULL, 'VIEW ANY DATABASE')) AS HasViewAnyDatabase;
'@

$masterFileQuery = @'
SELECT
    d.database_id AS DatabaseId,
    d.name AS DatabaseName,
    d.state_desc AS DatabaseState,
    mf.name AS LogicalFileName,
    mf.type_desc AS FileType,
    mf.physical_name AS FilePath,
    CONVERT(decimal(19, 2), mf.size * 8.0 / 1024.0) AS FileAllocatedSizeMB
FROM sys.databases AS d
INNER JOIN sys.master_files AS mf ON mf.database_id = d.database_id
ORDER BY d.database_id, mf.file_id;
'@

$databaseFileQuery = @'
SELECT
    DB_ID() AS DatabaseId,
    DB_NAME() AS DatabaseName,
    CONVERT(nvarchar(60), DATABASEPROPERTYEX(DB_NAME(), 'Status')) AS DatabaseState,
    file_id AS FileId,
    name AS LogicalFileName,
    type_desc AS FileType,
    physical_name AS FilePath,
    CONVERT(decimal(19, 2), size * 8.0 / 1024.0) AS FileAllocatedSizeMB,
    CASE WHEN type = 0 THEN
        CONVERT(decimal(19, 2), FILEPROPERTY(name, 'SpaceUsed') * 8.0 / 1024.0)
    ELSE NULL END AS FileUsedSpaceMB
FROM sys.database_files
ORDER BY file_id;
'@

$logUsageQuery = @'
SELECT
    CONVERT(decimal(19, 2), used_log_space_in_bytes / 1048576.0) AS DatabaseLogUsedSpaceMB
FROM sys.dm_db_log_space_usage;
'@

foreach ($service in $instanceServices) {
    $expectedInstanceName = if ($service.Name -eq 'MSSQLSERVER') {
        'DEFAULT'
    }
    else {
        $service.Name.Substring('MSSQL$'.Length)
    }
    $connectionTarget = if ($expectedInstanceName -eq 'DEFAULT') {
        '.'
    }
    else {
        '.\{0}' -f $expectedInstanceName
    }
    $context = @{
        ComputerName = $computerName; ConnectionTarget = $connectionTarget; SqlServerName = $null
        SqlMachineName = $null; SqlInstanceName = $null; SqlIsClustered = $null
        ServiceName = $service.Name; ServiceStatus = [string]$service.Status
        InstanceAvailability = 'Unavailable'
    }

    if ($service.Status -ne 'Running') {
        Write-Log -Message "Instance service '$($service.Name)' is $($service.Status)." -Level 'WARN'
        Add-InventoryRow -Rows $reportRows -Context $context -InventoryComplete 'False' `
            -Notes 'SQL Server service is stopped or unavailable.'
        continue
    }

    Write-Log -Message "Inventorying local target '$connectionTarget'." -Level 'INFO'
    try {
        $identity = @(Invoke-SqlQuery -DataSource $connectionTarget -InitialCatalog 'master' -Query $identityQuery)[0]
        if ($null -eq $identity) {
            throw 'SQL Server identity query returned no rows.'
        }

        $context.SqlServerName = $identity.SqlServerName
        $context.SqlMachineName = $identity.SqlMachineName
        $context.SqlInstanceName = if ([string]::IsNullOrWhiteSpace([string]$identity.SqlInstanceName)) {
            'DEFAULT'
        }
        else {
            [string]$identity.SqlInstanceName
        }
        $context.SqlIsClustered = $identity.SqlIsClustered
        $context.InstanceAvailability = 'Available'
        $catalogIsComplete = [int]$identity.HasViewAnyDatabase -eq 1
        $inventoryComplete = if ($catalogIsComplete) { 'True' } else { 'False' }
        $instanceNote = if ($catalogIsComplete) {
            $null
        }
        else {
            'VIEW ANY DATABASE is not granted; database inventory may be incomplete.'
        }

        $masterFiles = @(Invoke-SqlQuery -DataSource $connectionTarget -InitialCatalog 'master' -Query $masterFileQuery)
        if ($masterFiles.Count -eq 0) {
            Add-InventoryRow -Rows $reportRows -Context $context -InventoryComplete 'False' `
                -Notes 'Master query returned no visible database files.'
            Write-Log -Message "No visible database files on '$connectionTarget'." -Level 'WARN'
            continue
        }
    }
    catch {
        $detail = Get-SafeErrorDetail -Exception $_.Exception
        Write-Log -Message "Cannot inventory '$connectionTarget': $detail" -Level 'ERROR'
        Add-InventoryRow -Rows $reportRows -Context $context -InventoryComplete 'False' `
            -Notes 'Connection or master catalog query failed. Review execution log for details.'
        continue
    }

    foreach ($databaseGroup in ($masterFiles | Group-Object -Property DatabaseId)) {
        $databaseId = [int]$databaseGroup.Name
        $databaseName = [string]$databaseGroup.Group[0].DatabaseName
        $databaseState = [string]$databaseGroup.Group[0].DatabaseState
        $databaseNotes = @($instanceNote | Where-Object { $_ })

        if ($databaseState -ne 'ONLINE') {
            $databaseNotes += "Database is $databaseState; used-space metrics are unavailable."
            foreach ($databaseFile in $databaseGroup.Group) {
                Add-InventoryRow -Rows $reportRows -Context $context -DatabaseFile $databaseFile `
                    -InventoryComplete 'False' -Notes ($databaseNotes -join ' | ')
            }
            continue
        }

        try {
            $databaseFiles = @(Invoke-SqlQuery -DataSource $connectionTarget -InitialCatalog $databaseName -Query $databaseFileQuery)
        }
        catch {
            $detail = Get-SafeErrorDetail -Exception $_.Exception
            Write-Log -Message "Cannot inventory database '$databaseName' on '$connectionTarget': $detail" -Level 'WARN'
            $databaseNotes += 'Database query failed; used-space metrics are unavailable.'
            foreach ($databaseFile in $databaseGroup.Group) {
                Add-InventoryRow -Rows $reportRows -Context $context -DatabaseFile $databaseFile `
                    -InventoryComplete 'False' -Notes ($databaseNotes -join ' | ')
            }
            continue
        }

        $logUsage = $null
        try {
            $logUsage = @(Invoke-SqlQuery -DataSource $connectionTarget -InitialCatalog $databaseName -Query $logUsageQuery)[0]
        }
        catch {
            $detail = Get-SafeErrorDetail -Exception $_.Exception
            Write-Log -Message "Log usage is unavailable for '$databaseName' on '$connectionTarget': $detail" -Level 'WARN'
            $databaseNotes += 'Database log aggregate usage is unavailable.'
        }

        foreach ($databaseFile in $databaseFiles) {
            Add-InventoryRow -Rows $reportRows -Context $context -DatabaseFile $databaseFile `
                -LogUsage $logUsage -InventoryComplete $inventoryComplete -Notes ($databaseNotes -join ' | ')
        }
    }
}

# ==================================================================
# CSV EXPORT
# ==================================================================

$outputPath = Join-Path -Path $OutputDirectory -ChildPath (
    'MSSQL-inv-DBFiles_{0}.csv' -f $runId
)

$reportRows | Select-Object -Property $csvColumns | Export-Csv -LiteralPath $outputPath `
    -Delimiter ';' -NoTypeInformation -Encoding UTF8

if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
    Write-Log -Message 'CSV output file was not created.' -Level 'ERROR'
    exit 1
}

Write-Log -Message "Report saved: $outputPath (rows: $($reportRows.Count))." -Level 'SUCCESS'
Write-Log -Message "Log file: $logFile" -Level 'INFO'
