# vybe-test: powershell/readonly_variables/readonly_variable_set_variable_cmdlet_fails
New-Variable -Name "CMDLET_RO" -Value "Static" -Option ReadOnly
try {
    Set-Variable -Name "CMDLET_RO" -Value "Dynamic" -ErrorAction Stop
    Write-Host "FAIL: Set-Variable on ReadOnly variable succeeded without -Force"
    exit 1
} catch {
    if ($CMDLET_RO -ne "Static") {
        Write-Host "FAIL: Set-Variable mutated ReadOnly variable without -Force"
        exit 1
    }
}
Write-Host "PASS"
exit 0
