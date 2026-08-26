# vybe-test: powershell/parameters_validate_count/validatecount_with_splatted_hashtable
function Set-Targets {
    param([ValidateCount(1, 4)][string[]]$Targets)
    return "TargetsCount:$($Targets.Length)"
}
$params = @{ Targets = @("t1", "t2") }
$res = Set-Targets @params
if ($res -ne "TargetsCount:2") {
    Write-Host "FAIL: ValidateCount splatted hashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
