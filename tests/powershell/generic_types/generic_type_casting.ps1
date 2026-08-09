# vybe-test: powershell/generic_types/generic_type_casting
$raw = @("X", "Y")
$typed = [System.Collections.Generic.List[string]]$raw
if (-not ($typed -is [System.Collections.Generic.List[string]])) {
    Write-Host "FAIL: cast to List[string] failed"
    exit 1
}
if ($typed[1] -ne "Y") {
    Write-Host "FAIL: typed[1] expected Y, got $($typed[1])"
    exit 1
}
Write-Host "PASS"
exit 0
