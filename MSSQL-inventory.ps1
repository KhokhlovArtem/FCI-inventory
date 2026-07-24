#requires -Version 5.1

# ==============================================================================
# SCRIPT NAME: MSSQL-inventory.ps1
#
# Purpose:
#   - Inventories database files of all local SQL Server instances.
#   - Exports one CSV row per database file, including system databases.
#   - Reports stopped or unavailable instances as dedicated CSV rows.
#
# Requirements:
#   - Windows PowerShell 5.1.
#   - Network access and permission to connect to each local SQL Server instance.
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
    [string]$OutputDirectory = (Join-Path -Path $PWD -ChildPath 'reports'),

    [Parameter()]
    [ValidateRange(1, 300)]
    [int]$SqlConnectionTimeoutSeconds = 15,

    [Parameter()]
    [switch]$EncryptSqlConnection,

    [Parameter()]
    [switch]$TrustServerCertificate
)

$ErrorActionPreference = 'Stop'

# ==================================================================
# LOGGING
# ==================================================================

$scriptName = 'MSSQL-inventory'
$logDirectory = Join-Path -Path $PWD -ChildPath 'logs'
if (-not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
}

$logFile = Join-Path -Path $logDirectory -ChildPath (
    '{0}_{1:yyyy-MM-dd_HH-mm-ss}.log' -f $scriptName, (Get-Date)
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

# ==================================================================
# CONFIGURATION AND PREREQUISITES
# ==================================================================

if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
    New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}

$sqlUser = [Environment]::GetEnvironmentVariable('MSSQL_INVENTORY_SQL_USER', 'Process')
$sqlPassword = [Environment]::GetEnvironmentVariable('MSSQL_INVENTORY_SQL_PASSWORD', 'Process')
$useSqlAuthentication = -not [string]::IsNullOrWhiteSpace($sqlUser) -and
    -not [string]::IsNullOrWhiteSpace($sqlPassword)

if ((-not $useSqlAuthentication) -and ($sqlUser -or $sqlPassword)) {
    Write-Log -Message 'Only one SQL authentication environment variable is set. Using Windows authentication.' -Level 'WARN'
}

Write-Log -Message 'Script started.' -Level 'INFO'
Write-Log -Message "Output directory: $OutputDirectory" -Level 'INFO'
Write-Log -Message ('Authentication: ' + $(if ($useSqlAuthentication) { 'SQL authentication from environment' } else { 'Windows authentication' })) -Level 'INFO'

# ==================================================================
# HELPER FUNCTIONS
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

    New-Object System.Data.SqlClient.SqlConnection $builder.ConnectionString
}

function Invoke-SqlQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DataSource,
        [Parameter(Mandatory)][string]$InitialCatalog,
        [Parameter(Mandatory)][string]$Query
    )

    $connection = New-SqlConnection -DataSource $DataSource -InitialCatalog $InitialCatalog
    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandTimeout = 60
        $command.CommandText = $Query
        $reader = $command.ExecuteReader()
        $rows = New-Object System.Collections.Generic.List[object]

        try {
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
            $reader.Close()
            $command.Dispose()
        }

        return @($rows)
    }
    finally {
        if ($connection.State -ne [System.Data.ConnectionState]::Closed) {
            $connection.Close()
        }
        $connection.Dispose()
    }
}

function Add-InventoryRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Rows,
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][string]$InstanceName,
        [Parameter(Mandatory)][string]$ServiceName,
        [Parameter(Mandatory)][string]$ServiceStatus,
        [Parameter(Mandatory)][string]$InstanceAvailability,
        [Parameter()]$DatabaseFile,
        [Parameter()][string]$DatabaseState,
        [Parameter()][string]$Notes
    )

    $fileSizeMB = if ($DatabaseFile -and $null -ne $DatabaseFile.FileSizeMB) {
        [math]::Round([double]$DatabaseFile.FileSizeMB, 2)
    }
    else { $null }
    $usedSpaceMB = if ($DatabaseFile -and $null -ne $DatabaseFile.UsedSpaceMB) {
        [math]::Round([double]$DatabaseFile.UsedSpaceMB, 2)
    }
    else { $null }

    $Rows.Add([pscustomobject][ordered]@{
        CollectedAt          = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        ServerName           = $ServerName
        InstanceName         = $InstanceName
        ServiceName          = $ServiceName
        ServiceStatus        = $ServiceStatus
        InstanceAvailability = $InstanceAvailability
        DatabaseName         = if ($DatabaseFile) { [string]$DatabaseFile.DatabaseName } else { $null }
        DatabaseState        = $DatabaseState
        LogicalFileName      = if ($DatabaseFile) { [string]$DatabaseFile.LogicalFileName } else { $null }
        FileType             = if ($DatabaseFile) { [string]$DatabaseFile.FileType } else { $null }
        FilePath             = if ($DatabaseFile) { [string]$DatabaseFile.FilePath } else { $null }
        FileSizeMB           = $fileSizeMB
        FileSizeGB           = if ($null -ne $fileSizeMB) { [math]::Round($fileSizeMB / 1024, 2) } else { $null }
        UsedSpaceMB          = $usedSpaceMB
        UsedSpaceGB          = if ($null -ne $usedSpaceMB) { [math]::Round($usedSpaceMB / 1024, 2) } else { $null }
        Notes                = $Notes
    })
}

# ==================================================================
# SQL SERVER INSTANCE DISCOVERY
# ==================================================================

$serverName = $env:COMPUTERNAME
$instanceServices = @(Get-Service -ErrorAction Stop | Where-Object {
    $_.Name -eq 'MSSQLSERVER' -or $_.Name -like 'MSSQL$*'
})

if ($instanceServices.Count -eq 0) {
    Write-Log -Message 'No local SQL Server services were found.' -Level 'WARN'
}
else {
    Write-Log -Message "SQL Server services found: $($instanceServices.Count)." -Level 'SUCCESS'
}

# ==================================================================
# DATABASE FILE INVENTORY
# ==================================================================

$reportRows = New-Object System.Collections.Generic.List[object]
$masterFileQuery = @'
SELECT
    d.name AS DatabaseName,
    d.state_desc AS DatabaseState,
    mf.name AS LogicalFileName,
    mf.type_desc AS FileType,
    mf.physical_name AS FilePath,
    CONVERT(decimal(19, 2), mf.size * 8.0 / 1024.0) AS FileSizeMB
FROM sys.databases AS d
INNER JOIN sys.master_files AS mf ON mf.database_id = d.database_id
ORDER BY d.name, mf.file_id;
'@

$databaseFileQuery = @'
SELECT
    DB_NAME() AS DatabaseName,
    name AS LogicalFileName,
    type_desc AS FileType,
    physical_name AS FilePath,
    CONVERT(decimal(19, 2), size * 8.0 / 1024.0) AS FileSizeMB,
    CONVERT(decimal(19, 2), FILEPROPERTY(name, 'SpaceUsed') * 8.0 / 1024.0) AS UsedSpaceMB
FROM sys.database_files
ORDER BY file_id;
'@

foreach ($service in $instanceServices) {
    $instanceName = if ($service.Name -eq 'MSSQLSERVER') {
        'MSSQLSERVER'
    }
    else {
        $service.Name.Substring('MSSQL$'.Length)
    }
    $dataSource = if ($instanceName -eq 'MSSQLSERVER') {
        $serverName
    }
    else {
        '{0}\{1}' -f $serverName, $instanceName
    }
    $serviceStatus = [string]$service.Status

    if ($service.Status -ne 'Running') {
        Write-Log -Message "Instance '$instanceName' is $serviceStatus." -Level 'WARN'
        Add-InventoryRow -Rows $reportRows -ServerName $serverName -InstanceName $instanceName `
            -ServiceName $service.Name -ServiceStatus $serviceStatus -InstanceAvailability 'Unavailable' `
            -Notes 'SQL Server service is stopped or unavailable.'
        continue
    }

    Write-Log -Message "Inventorying instance '$instanceName'." -Level 'INFO'
    try {
        $masterFiles = @(Invoke-SqlQuery -DataSource $dataSource -InitialCatalog 'master' -Query $masterFileQuery)
    }
    catch {
        Write-Log -Message "Cannot connect to instance '$instanceName'." -Level 'ERROR'
        Add-InventoryRow -Rows $reportRows -ServerName $serverName -InstanceName $instanceName `
            -ServiceName $service.Name -ServiceStatus $serviceStatus -InstanceAvailability 'Unavailable' `
            -Notes 'Connection to SQL Server failed. Review execution log for details.'
        continue
    }

    $masterFilesByDatabase = $masterFiles | Group-Object -Property DatabaseName
    foreach ($databaseGroup in $masterFilesByDatabase) {
        $databaseName = [string]$databaseGroup.Name
        $databaseState = [string]$databaseGroup.Group[0].DatabaseState

        if ($databaseState -ne 'ONLINE') {
            foreach ($databaseFile in $databaseGroup.Group) {
                Add-InventoryRow -Rows $reportRows -ServerName $serverName -InstanceName $instanceName `
                    -ServiceName $service.Name -ServiceStatus $serviceStatus -InstanceAvailability 'Available' `
                    -DatabaseFile $databaseFile -DatabaseState $databaseState `
                    -Notes "Database is $databaseState; used space is unavailable."
            }
            continue
        }

        try {
            $databaseFiles = @(Invoke-SqlQuery -DataSource $dataSource -InitialCatalog $databaseName -Query $databaseFileQuery)
            foreach ($databaseFile in $databaseFiles) {
                Add-InventoryRow -Rows $reportRows -ServerName $serverName -InstanceName $instanceName `
                    -ServiceName $service.Name -ServiceStatus $serviceStatus -InstanceAvailability 'Available' `
                    -DatabaseFile $databaseFile -DatabaseState $databaseState
            }
        }
        catch {
            Write-Log -Message "Cannot inventory database '$databaseName' on '$instanceName'." -Level 'WARN'
            foreach ($databaseFile in $databaseGroup.Group) {
                Add-InventoryRow -Rows $reportRows -ServerName $serverName -InstanceName $instanceName `
                    -ServiceName $service.Name -ServiceStatus $serviceStatus -InstanceAvailability 'Available' `
                    -DatabaseFile $databaseFile -DatabaseState $databaseState `
                    -Notes 'Database file size was collected from master; used space is unavailable.'
            }
        }
    }
}

# ==================================================================
# CSV EXPORT
# ==================================================================

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outputPath = Join-Path -Path $OutputDirectory -ChildPath (
    'MSSQL-inv-DBFiles_{0}.csv' -f $timestamp
)

$reportRows | Export-Csv -LiteralPath $outputPath -Delimiter ';' -NoTypeInformation -Encoding UTF8
Write-Log -Message "Report saved: $outputPath (rows: $($reportRows.Count))." -Level 'SUCCESS'
Write-Log -Message "Log file: $logFile" -Level 'INFO'
