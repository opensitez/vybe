# vybe-test: powershell/loops/do_until_evaluates_condition_after_execution
# A do-until loop always executes its body at least once even when condition is initially true
$executionCount = 0

do {
    $executionCount++
} until ($true)

if ($executionCount -ne 1) {
    Write-Host "FAIL: expected exactly 1 execution, got $executionCount"
    exit 1
}

Write-Host "PASS"
exit 0
