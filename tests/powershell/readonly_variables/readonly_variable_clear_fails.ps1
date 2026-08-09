# vybe-test: powershell/readonly_variables/readonly_variable_clear_fails
New-Variable -Name "UNCLEAR_RO" -Value 555 -Option ReadOnly
try {
    Clear-Variable -Name "UNCLEAR_RO" -ErrorAction Stop
    Write-Host "FAIL: Clear-Variable on ReadOnly variable succeeded without -Force"
    exit 1
} catch {
    if ($UNCLEAR_RO -ne 555) {
        Write-Host "FAIL: ReadOnly variable cleared without -Force"
        exit 1
    }
}
Write-Host "PASS"
exit 0
