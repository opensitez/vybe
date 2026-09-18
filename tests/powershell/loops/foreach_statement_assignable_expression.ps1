# vybe-test: powershell/loops/foreach_statement_assignable_expression
# A foreach statement can be assigned directly to a variable to collect yielded pipeline outputs
$squared = foreach ($num in 1..5) {
    $num * $num
}

if ($squared.Count -ne 5) {
    Write-Host "FAIL: expected 5 elements, got $($squared.Count)"
    exit 1
}

$expected = @(1, 4, 9, 16, 25)
for ($i = 0; $i -lt 5; $i++) {
    if ($squared[$i] -ne $expected[$i]) {
        Write-Host "FAIL: at index $i, expected $($expected[$i]), got $($squared[$i])"
        exit 1
    }
}

Write-Host "PASS"
exit 0
