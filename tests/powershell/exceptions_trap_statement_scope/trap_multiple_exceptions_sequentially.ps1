# vybe-test: powershell/exceptions_trap_statement_scope/trap_multiple_exceptions_sequentially
$log = [System.Collections.Generic.List[string]]::new()
function Test-MultiTrap {
    trap {
        $log.Add($_.Exception.Message)
        continue
    }
    throw "First"
    throw "Second"
}
Test-MultiTrap
if ($log.Count -ne 2 -or $log[0] -ne "First" -or $log[1] -ne "Second") {
    Write-Host "FAIL: Sequential trap handling failed, count=$($log.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
