# vybe-test: powershell/readonly_variables/readonly_variable_option_flag
New-Variable -Name "OPT_RO" -Value "FLAG" -Option ReadOnly
$varObj = Get-Variable -Name "OPT_RO"
if (-not ($varObj.Options -band [System.Management.Automation.ScopedItemOptions]::ReadOnly)) {
    Write-Host "FAIL: ScopedItemOptions ReadOnly flag missing"
    exit 1
}
Write-Host "PASS"
exit 0
