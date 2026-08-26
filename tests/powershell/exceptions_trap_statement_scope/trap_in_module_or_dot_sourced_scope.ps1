# vybe-test: powershell/exceptions_trap_statement_scope/trap_in_module_or_dot_sourced_scope
$trapped = $false
. {
    trap { $script:trapped = $true; continue }
    1 / 0
}
if (-not $trapped) {
    Write-Host "FAIL: Trap in dot-sourced block failed"
    exit 1
}
Write-Host "PASS"
exit 0
