# vybe-test: powershell/hashtables/hashtable_nested
$config = @{
    database = @{
        host = "localhost"
        port = 5432
    }
    cache = @{
        ttl = 300
    }
}
if ($config.database.host -ne "localhost") { Write-Host "FAIL: host";    exit 1 }
if ($config.database.port -ne 5432)        { Write-Host "FAIL: port";    exit 1 }
if ($config.cache.ttl     -ne 300)         { Write-Host "FAIL: ttl";     exit 1 }
Write-Host "PASS"
exit 0
