# vybe-test: powershell/trap_statements/function_trap
$trapped = $false
function Test-FuncTrap {
    trap { $script:trapped = $true; continue }
    throw "TrapError"
}
Test-FuncTrap
if (-not $trapped) {
    Write-Host "FAIL: Function trap failed"
    exit 1
}
Write-Host "PASS"
exit 0
