# vybe-test: powershell/exceptions_trap_statement_scope/child_scope_trap_overrides_parent_trap
$childRan = $false
$parentRan = $false
function Test-ChildTrapOverride {
    trap { $script:parentRan = $true; continue }
    & {
        trap { $script:childRan = $true; continue }
        1 / 0
    }
}
Test-ChildTrapOverride
if (-not $childRan -or $parentRan) {
    Write-Host "FAIL: Child trap overriding parent trap failed"
    exit 1
}
Write-Host "PASS"
exit 0
