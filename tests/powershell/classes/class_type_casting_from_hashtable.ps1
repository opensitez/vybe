# vybe-test: powershell/classes/class_type_casting_from_hashtable
# A hashtable can be cast directly to a user-defined PowerShell class type to instantiate it
class DatabaseConfig {
    [string]$Host
    [int]$Port
    [bool]$UseSsl
}

$hash = @{
    Host   = "db.internal"
    Port   = 5432
    UseSsl = $true
}

$dbConfig = [DatabaseConfig]$hash

if ($dbConfig.GetType().Name -ne "DatabaseConfig") {
    Write-Host "FAIL: expected type DatabaseConfig, got $($dbConfig.GetType().Name)"
    exit 1
}

if ($dbConfig.Host -ne "db.internal" -or $dbConfig.Port -ne 5432 -or (-not $dbConfig.UseSsl)) {
    Write-Host "FAIL: property values not mapped properly from hashtable"
    exit 1
}

Write-Host "PASS"
exit 0
