# vybe-test: powershell/cmdlets/get_unique_cmdlet_pipeline
# Get-Unique filters adjacent duplicate values from a sorted pipeline stream
$sortedInput = @(1, 1, 2, 2, 2, 3, 4, 4, 5)

$uniqueResults = $sortedInput | Get-Unique

if ($uniqueResults.Count -ne 5) {
    Write-Host "FAIL: expected 5 unique elements, got $($uniqueResults.Count)"
    exit 1
}

$expected = @(1, 2, 3, 4, 5)
for ($i = 0; $i -lt 5; $i++) {
    if ($uniqueResults[$i] -ne $expected[$i]) {
        Write-Host "FAIL: mismatch at index ${i} - expected $($expected[$i]), got $($uniqueResults[$i])"
        exit 1
    }
}

Write-Host "PASS"
exit 0
