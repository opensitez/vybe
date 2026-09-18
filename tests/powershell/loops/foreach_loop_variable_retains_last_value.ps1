# vybe-test: powershell/loops/foreach_loop_variable_retains_last_value
# The foreach loop iteration variable retains the value of the final element after completion
$elements = @(10, 20, 30, 40)

foreach ($current in $elements) {
    # iterate through collection
}

if ($current -ne 40) {
    Write-Host "FAIL: expected loop variable to retain last value 40, got '$current'"
    exit 1
}

Write-Host "PASS"
exit 0
