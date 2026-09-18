# vybe-test: powershell/process_record_streaming/process_block_return_statement_exits_current_iteration_only
# An explicit return statement inside a process block exits only the current element's iteration
function FilterOddsViaReturn {
    process {
        if ($_ % 2 -eq 1) {
            return
        }
        $_
    }
}

$evenNumbers = @(1, 2, 3, 4, 5, 6) | FilterOddsViaReturn

if ($evenNumbers.Count -ne 3) {
    Write-Host "FAIL: expected 3 even numbers, got $($evenNumbers.Count)"
    exit 1
}

$expected = "2, 4, 6"
$actual = $evenNumbers -join ", "

if ($actual -ne $expected) {
    Write-Host "FAIL: expected '$expected', got '$actual'"
    exit 1
}

Write-Host "PASS"
exit 0
