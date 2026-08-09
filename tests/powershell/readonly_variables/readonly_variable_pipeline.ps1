# vybe-test: powershell/readonly_variables/readonly_variable_pipeline
New-Variable -Name "RO_PIPE" -Value @("A", "B") -Option ReadOnly
$res = $RO_PIPE | ForEach-Object { $_.ToLower() }
if ($res[0] -ne "a" -or $res[1] -ne "b") {
    Write-Host "FAIL: ReadOnly variable in pipeline expected 'a', 'b'"
    exit 1
}
Write-Host "PASS"
exit 0
