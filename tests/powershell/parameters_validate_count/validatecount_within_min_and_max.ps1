# vybe-test: powershell/parameters_validate_count/validatecount_within_min_and_max
function Select-Servers {
    param([ValidateCount(1, 3)][string[]]$Servers)
    return $Servers.Length
}
$r1 = Select-Servers -Servers "srv1"
$r2 = Select-Servers -Servers "srv1", "srv2", "srv3"
if ($r1 -ne 1 -or $r2 -ne 3) {
    Write-Host "FAIL: ValidateCount within bounds failed"
    exit 1
}
Write-Host "PASS"
exit 0
