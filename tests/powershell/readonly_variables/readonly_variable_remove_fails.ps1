# vybe-test: powershell/readonly_variables/readonly_variable_remove_fails
New-Variable -Name "NO_REMOVE_RO" -Value "Keep" -Option ReadOnly
try {
    Remove-Variable -Name "NO_REMOVE_RO" -ErrorAction Stop
    Write-Host "FAIL: Remove-Variable on ReadOnly variable succeeded without -Force"
    exit 1
} catch {
    if ($NO_REMOVE_RO -ne "Keep") {
        Write-Host "FAIL: ReadOnly variable removed without -Force"
        exit 1
    }
}
Write-Host "PASS"
exit 0
