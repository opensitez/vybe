# vybe-test: powershell/out_variables/out_variable_scope_caller
function Invoke-Sub {
    1..2 | ForEach-Object { $_ } -OutVariable script:subOut | Out-Null
}
Invoke-Sub
if ($script:subOut.Count -ne 2) {
    Write-Host "FAIL: script-scoped OutVariable capture expected 2 items"
    exit 1
}
Write-Host "PASS"
exit 0
