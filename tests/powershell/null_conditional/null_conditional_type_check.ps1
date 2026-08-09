# vybe-test: powershell/null_conditional/null_conditional_type_check
$date = [datetime]::Now
$res = ${date}?.AddDays(1)
if (-not ($res -is [datetime])) {
    Write-Host "FAIL: non-null conditional method return is not [datetime]"
    exit 1
}
Write-Host "PASS"
exit 0
