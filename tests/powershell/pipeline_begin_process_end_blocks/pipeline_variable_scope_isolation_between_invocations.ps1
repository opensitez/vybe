# vybe-test: powershell/pipeline_begin_process_end_blocks/pipeline_variable_scope_isolation_between_invocations
function Test-ScopeIsolate {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    begin { $localCount = 0 }
    process { $localCount += $Val }
    end { return $localCount }
}
$r1 = 1, 2, 3 | Test-ScopeIsolate
$r2 = 10, 20 | Test-ScopeIsolate
if ($r1 -ne 6 -or $r2 -ne 30) {
    Write-Host "FAIL: Scope isolation between separate pipeline invocations failed"
    exit 1
}
Write-Host "PASS"
exit 0
