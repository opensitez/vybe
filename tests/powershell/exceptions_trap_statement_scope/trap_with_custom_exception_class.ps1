# vybe-test: powershell/exceptions_trap_statement_scope/trap_with_custom_exception_class
class CustomTrapEx : System.Exception {
    CustomTrapEx([string]$m) : base($m) {}
}
$caught = $false
function Test-CustomTrap {
    trap [CustomTrapEx] {
        $script:caught = $true
        continue
    }
    throw [CustomTrapEx]::new("CustomTrap")
}
Test-CustomTrap
if (-not $caught) {
    Write-Host "FAIL: Trap with custom exception class failed"
    exit 1
}
Write-Host "PASS"
exit 0
