# vybe-test: powershell/exceptions_trap_statement_scope/trap_scope_inheritance_in_child_scope
$parentTrapped = $false
function Test-ParentTrap {
    trap {
        $script:parentTrapped = $true
        continue
    }
    & {
        # Child scope with no trap
        1 / 0
    }
}
Test-ParentTrap
if (-not $parentTrapped) {
    Write-Host "FAIL: Trap scope inheritance in child scope failed"
    exit 1
}
Write-Host "PASS"
exit 0
