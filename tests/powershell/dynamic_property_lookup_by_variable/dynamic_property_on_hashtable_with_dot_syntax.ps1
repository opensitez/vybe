# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_on_hashtable_with_dot_syntax
$ht = @{ host = "127.0.0.1"; port = 8080 }
$pHost = "host"
$pPort = "port"
if ($ht.$pHost -ne "127.0.0.1" -or $ht.$pPort -ne 8080) {
    Write-Host "FAIL: Dynamic dot property lookup on hashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
