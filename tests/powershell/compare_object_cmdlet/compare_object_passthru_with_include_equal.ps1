# vybe-test: powershell/compare_object_cmdlet/compare_object_passthru_with_include_equal
# -PassThru and -IncludeEqual emits matching elements with original type and '==' SideIndicator
$diff = Compare-Object @(5, 10) @(10, 15) -IncludeEqual -PassThru

$eqItem = $diff | Where-Object { $_.SideIndicator -eq '==' }

if ($null -eq $eqItem) {
    Write-Host "FAIL: equal element missing in PassThru output"
    exit 1
}

if ($eqItem -ne 10) {
    Write-Host "FAIL: expected equal item 10, got: $eqItem"
    exit 1
}

Write-Host "PASS"
exit 0
