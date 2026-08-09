# vybe-test: powershell/out_variables/out_variable_empty_pipeline
@() | Where-Object { $_ } -OutVariable emptyCap | Out-Null
if ($emptyCap.Count -ne 0) {
    Write-Host "FAIL: empty pipeline OutVariable expected Count 0, got $($emptyCap.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
