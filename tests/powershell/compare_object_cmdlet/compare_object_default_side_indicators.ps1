# vybe-test: powershell/compare_object_cmdlet/compare_object_default_side_indicators
# Compare-Object marks reference-only elements with '<=' and difference-only with '=>'
$diff = Compare-Object @(1, 2) @(2, 3)

$left  = $diff | Where-Object { $_.SideIndicator -eq '<=' }
$right = $diff | Where-Object { $_.SideIndicator -eq '=>' }

if ($left.InputObject -ne 1) {
    Write-Host "FAIL: expected reference element 1 with '<=', got: $($left.InputObject)"
    exit 1
}

if ($right.InputObject -ne 3) {
    Write-Host "FAIL: expected difference element 3 with '=>', got: $($right.InputObject)"
    exit 1
}

Write-Host "PASS"
exit 0
