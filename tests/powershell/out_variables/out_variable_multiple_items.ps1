# vybe-test: powershell/out_variables/out_variable_multiple_items
1..5 | Where-Object { $_ % 2 -eq 0 } -OutVariable evens | Out-Null
if ($evens.Count -ne 2 -or $evens[0] -ne 2 -or $evens[1] -ne 4) {
    Write-Host "FAIL: OutVariable multiple items expected 2, 4"
    exit 1
}
Write-Host "PASS"
exit 0
