# vybe-test: powershell/readonly_variables/readonly_variable_array_type
New-Variable -Name "RO_ARRAY" -Value @(1, 2, 3) -Option ReadOnly
if ($RO_ARRAY.Length -ne 3 -or $RO_ARRAY[0] -ne 1) {
    Write-Host "FAIL: ReadOnly array variable expected Length 3, item 1"
    exit 1
}
Write-Host "PASS"
exit 0
