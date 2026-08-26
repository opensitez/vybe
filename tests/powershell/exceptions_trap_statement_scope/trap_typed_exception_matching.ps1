# vybe-test: powershell/exceptions_trap_statement_scope/trap_typed_exception_matching
$matched = ""
function Test-TypedTrap {
    trap [System.DivideByZeroException] {
        $script:matched = "DivideByZero"
        continue
    }
    trap [System.FormatException] {
        $script:matched = "Format"
        continue
    }
    1 / 0
}
Test-TypedTrap
if ($matched -ne "DivideByZero") {
    Write-Host "FAIL: Typed trap exception matching failed, got '$matched'"
    exit 1
}
Write-Host "PASS"
exit 0
