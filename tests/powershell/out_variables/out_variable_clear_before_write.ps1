# vybe-test: powershell/out_variables/out_variable_clear_before_write
$cap = @(100, 200, 300)
"New" | ForEach-Object { $_ } -OutVariable cap | Out-Null
if ($cap.Count -ne 1 -or $cap[0] -ne "New") {
    Write-Host "FAIL: non-append OutVariable expected overwrite with 'New'"
    exit 1
}
Write-Host "PASS"
exit 0
