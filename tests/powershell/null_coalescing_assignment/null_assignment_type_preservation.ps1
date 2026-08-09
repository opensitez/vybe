# vybe-test: powershell/null_coalescing_assignment/null_assignment_type_preservation
$val = $null
$val ??= [datetime]::Now
if (-not ($val -is [datetime])) {
    Write-Host "FAIL: ??= [datetime] assignment lost type"
    exit 1
}
Write-Host "PASS"
exit 0
