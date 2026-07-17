# FCI-inventory

Inventory script for Windows Failover Cluster roles. Generates separate CSV reports for SQL Server FCI and Hyper-V VM roles with detailed disk and memory information.

## Reports

| File pattern | Content |
|---|---|
| `FCI-inv-SQL_<timestamp>.csv` | SQL Server Failover Cluster Instances — memory limits (min/max server memory), current memory usage, product version, disk inventory |
| `FCI-inv-VM_<timestamp>.csv` | Hyper-V Virtual Machines — startup/assigned/min/max memory, dynamic memory status, disk inventory |

## Requirements

- Windows PowerShell 5.1
- `FailoverClusters` module
- Administrator rights
- Cluster read permissions
- PowerShell Remoting (WinRM) to cluster owner nodes for disk and VM queries
- Network access to SQL FCI for min/max server memory queries

## Usage

```powershell
# Preview generated directories and reports before the first run
.\FCI-inventory.ps1 -WhatIf

# Basic inventory of the local cluster
.\FCI-inventory.ps1

# Specify a remote cluster and output directory
.\FCI-inventory.ps1 -Cluster CLUSTER01 -OutputDirectory C:\Reports

# Filter by role name (supports wildcards)
.\FCI-inventory.ps1 -RoleName 'SQL *','VM *'

# Skip SQL queries
.\FCI-inventory.ps1 -SkipSqlQuery

# Use SQL authentication
$cred = Get-Credential -Message 'SQL login'
.\FCI-inventory.ps1 -SqlCredential $cred

# Use alternate credentials for cluster node access
.\FCI-inventory.ps1 -NodeCredential (Get-Credential)
```

## Output columns

- `CollectedAt`, `ClusterName`, `RoleName`, `RoleState`, `OwnerNode`
- `OwnerNodePhysicalMemoryGB`, `InstanceType`, `InstanceName`, `ConnectionName`
- `ProductVersion`, `MemorySource`, `MinMemoryMB`, `MaxMemoryMB`
- `StartupMemoryMB`, `AssignedMemoryMB`, `CurrentSqlMemoryMB`, `DynamicMemoryEnabled`
- `DiskResourceCount`, `RoleTotalDiskSizeGB`, `RoleTotalFreeSpaceGB`, `RoleDiskDataComplete`
- Per-disk: `DiskResourceName`, `DiskNumber`, `DiskGuid`, `PartitionStyle`, `BusType`, `MountPoints`, `FileSystemLabels`, `FileSystems`, `DiskSizeGB`, `VolumeSizeGB`, `FreeSpaceGB`, `FreePercent`
- `Notes` — warnings and errors

## Logs

Logs are written to `$env:USERPROFILE\powershell-script-logs` as
`FCI-inventory_<timestamp>_operation.log`. Log and report messages include only
safe operational context; detailed exception data is not written to files.

## Safety

- Use `-WhatIf` before the first production run to preview report and log directory creation.
- The script only creates timestamped CSV reports and log files. Remove those files to revert.
- Supply credentials with `Get-Credential`; never place passwords in command arguments or logs.

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-Cluster` | `string` | `.` (local) | Cluster name |
| `-OutputDirectory` | `string` | `.\reports` | Output directory for CSV files |
| `-RoleName` | `string[]` | all roles | Role name filter (supports `*` and `?`) |
| `-SqlCredential` | `PSCredential` | Integrated Security | SQL authentication |
| `-NodeCredential` | `PSCredential` | current user | Credentials for remote node access |
| `-SkipSqlQuery` | `switch` | false | Skip SQL Server queries |
| `-EncryptSqlConnection` | `switch` | false | Encrypt SQL connection |
| `-TrustServerCertificate` | `switch` | false | Trust SQL Server certificate |
| `-SqlConnectionTimeoutSeconds` | `int` | 15 | SQL connection timeout |
