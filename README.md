# FCI-inventory

Скрипт инвентаризации ролей Windows Failover Cluster. Создаёт отдельные CSV-отчёты для SQL Server FCI и виртуальных машин Hyper-V с подробными сведениями о дисках и памяти.

## Отчёты

| Шаблон имени файла | Содержимое |
|---|---|
| `FCI-inv-SQL_<timestamp>.csv` | Экземпляры SQL Server Failover Cluster: ограничения памяти (min/max server memory), текущее потребление памяти, версия продукта, инвентаризация дисков |
| `FCI-inv-VM_<timestamp>.csv` | Виртуальные машины Hyper-V: стартовая, назначенная, минимальная и максимальная память, состояние динамической памяти, инвентаризация дисков |
| `MSSQL-inv-DBFiles_<timestamp>.csv` | Файлы всех баз данных локальных экземпляров SQL Server, включая системные БД |

## Требования

- Windows PowerShell 5.1
- `FailoverClusters` module
- Права администратора
- Права чтения кластера
- PowerShell Remoting (WinRM) к узлам-владельцам ролей для запросов дисков и виртуальных машин
- Сетевой доступ к SQL FCI для запросов значений min/max server memory

## Использование

```powershell
# Предварительный просмотр создаваемых каталогов и отчётов перед первым запуском
.\FCI-inventory.ps1 -WhatIf

# Инвентаризация локального кластера
.\FCI-inventory.ps1

# Указание удалённого кластера и каталога вывода
.\FCI-inventory.ps1 -Cluster CLUSTER01 -OutputDirectory C:\Reports

# Фильтрация по имени роли (поддерживаются шаблоны)
.\FCI-inventory.ps1 -RoleName 'SQL *','VM *'

# Отключение SQL-запросов
.\FCI-inventory.ps1 -SkipSqlQuery

# Использование аутентификации SQL Server
$cred = Get-Credential -Message 'Учётные данные SQL'
.\FCI-inventory.ps1 -SqlCredential $cred

# Использование альтернативных учётных данных для доступа к узлам кластера
.\FCI-inventory.ps1 -NodeCredential (Get-Credential)
```

### Инвентаризация файлов SQL Server

```powershell
# Все локальные экземпляры SQL Server
.\MSSQL-inventory.ps1

# Сохранение отчёта в другой каталог
.\MSSQL-inventory.ps1 -OutputDirectory C:\Reports

# Шифрование подключения к SQL Server
.\MSSQL-inventory.ps1 -EncryptSqlConnection -TrustServerCertificate
```

`MSSQL-inventory.ps1` находит службы `MSSQLSERVER` и `MSSQL$<InstanceName>`.
Запущенные экземпляры инвентаризируются, а для остановленных или недоступных
экземпляров создаётся отдельная строка CSV со статусом `Unavailable` и причиной
в поле `Notes`.

В отчёт попадают файлы пользовательских и системных БД: `master`, `model`,
`msdb` и `tempdb`. Одна строка соответствует одному файлу БД.

Поля отчёта `MSSQL-inv-DBFiles`:

- `InstanceName`, `ServiceName`, `ServiceStatus`, `InstanceAvailability`
- `DatabaseName`, `DatabaseState`, `LogicalFileName`, `FileType`, `FilePath`
- `FileSizeMB`, `FileSizeGB` - полный выделенный размер файла
- `UsedSpaceMB`, `UsedSpaceGB` - занятое пространство по данным SQL Server
- `Notes` - причина отсутствия данных, включая остановленный экземпляр или offline БД

По умолчанию используется Windows-аутентификация текущего пользователя. Для
SQL-аутентификации предусмотрены переменные окружения
`MSSQL_INVENTORY_SQL_USER` и `MSSQL_INVENTORY_SQL_PASSWORD`; их шаблон находится
в `.env.example`. Скрипт намеренно не читает `.env` напрямую: секреты должны
быть загружены в окружение внешним безопасным механизмом.

## Столбцы отчёта

- `CollectedAt`, `ClusterName`, `RoleName`, `RoleState`, `OwnerNode`
- `OwnerNodePhysicalMemoryGB`, `InstanceType`, `InstanceName`, `ConnectionName`
- `ProductVersion`, `MemorySource`, `MinMemoryMB`, `MaxMemoryMB`
- `StartupMemoryMB`, `AssignedMemoryMB`, `CurrentSqlMemoryMB`, `DynamicMemoryEnabled`
- `DiskResourceCount`, `RoleTotalDiskSizeGB`, `RoleTotalFreeSpaceGB`, `RoleDiskDataComplete`
- Для каждого диска: `DiskResourceName`, `DiskNumber`, `DiskGuid`, `PartitionStyle`, `BusType`, `MountPoints`, `FileSystemLabels`, `FileSystems`, `DiskSizeGB`, `VolumeSizeGB`, `FreeSpaceGB`, `FreePercent`
- `Notes` — предупреждения и ошибки

## Журналы

Журналы сохраняются в `$env:USERPROFILE\powershell-script-logs` с именем
`FCI-inventory_<timestamp>_operation.log`. Журнал и отчёт содержат только
безопасный операционный контекст; подробности исключений не записываются в файлы.

## Безопасность

- Перед первым рабочим запуском используйте `-WhatIf`, чтобы просмотреть создание каталогов журналов и отчётов.
- Скрипт создаёт только CSV-отчёты и журналы с отметкой времени. Для отката удалите эти файлы.
- Передавайте учётные данные через `Get-Credential`; не указывайте пароли в аргументах команд или журналах.

## Параметры

| Параметр | Тип | Значение по умолчанию | Описание |
|---|---|---|---|
| `-Cluster` | `string` | `.` (локальный) | Имя кластера |
| `-OutputDirectory` | `string` | `.\reports` | Каталог для CSV-файлов |
| `-RoleName` | `string[]` | все роли | Фильтр имён ролей (поддерживает `*` и `?`) |
| `-SqlCredential` | `PSCredential` | Integrated Security | Аутентификация SQL Server |
| `-NodeCredential` | `PSCredential` | текущий пользователь | Учётные данные для удалённого доступа к узлам |
| `-SkipSqlQuery` | `switch` | false | Отключить запросы к SQL Server |
| `-EncryptSqlConnection` | `switch` | false | Шифровать подключение к SQL Server |
| `-TrustServerCertificate` | `switch` | false | Доверять сертификату SQL Server |
| `-SqlConnectionTimeoutSeconds` | `int` | 15 | Время ожидания подключения к SQL Server |
