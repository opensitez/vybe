# vybe-test: powershell/out_variables/out_variable_append
$buf = @(1)
1..2 | ForEach-Object { $_ * 10 } -OutVariable +buf | Out-Null
if ($buf.Count -ne 3 -or $buf[1] -ne 10 -or $buf[2] -ne 20) {
    Write-Host "FAIL: +OutVariable append expected Count 3 (1, 10, 20), got $($buf -join ', ')"
    exit 1
}
Write-Host "PASS"
exit 0
