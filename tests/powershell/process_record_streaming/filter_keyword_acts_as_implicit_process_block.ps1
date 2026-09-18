# vybe-test: powershell/process_record_streaming/filter_keyword_acts_as_implicit_process_block
# The filter keyword defines a command whose entire body implicitly runs as a process block for every input item
filter MultiplyByFour {
    $_ * 4
}

$results = @(2, 4, 6) | MultiplyByFour

if ($results.Count -ne 3) {
    Write-Host "FAIL: expected 3 filtered items, got $($results.Count)"
    exit 1
}

$expected = @(8, 16, 24)
for ($i = 0; $i -lt 3; $i++) {
    if ($results[$i] -ne $expected[$i]) {
        Write-Host "FAIL: at index ${i}, expected $($expected[$i]), got $($results[$i])"
        exit 1
    }
}

Write-Host "PASS"
exit 0
