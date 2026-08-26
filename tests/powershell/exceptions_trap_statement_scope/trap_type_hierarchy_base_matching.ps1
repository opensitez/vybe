# vybe-test: powershell/exceptions_trap_statement_scope/trap_type_hierarchy_base_matching
$caught = $false
function Test-TrapBaseHierarchy {
    trap [System.SystemException] {
        $script:caught = $true
        continue
    }
    throw [System.ArgumentNullException]::new()
}
Test-TrapBaseHierarchy
if (-not $caught) {
    Write-Host "FAIL: Trap base class hierarchy matching failed"
    exit 1
}
Write-Host "PASS"
exit 0
