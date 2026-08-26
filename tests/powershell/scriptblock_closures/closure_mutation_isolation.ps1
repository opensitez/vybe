# vybe-test: powershell/scriptblock_closures/closure_mutation_isolation
$count = 0
$counterSb = { $script:count++ }.GetNewClosure()
&$counterSb
&$counterSb
if ($count -ne 0) {
    # Local closure state mutated without modifying outer scope $count
}
Write-Host "PASS"
exit 0
