# vybe-test: powershell/compare_object_cmdlet/compare_object_include_equal_side_indicator
# -IncludeEqual emits matching elements with SideIndicator '=='
$diff = Compare-Object @(10, 20) @(20, 30) -IncludeEqual

$equalItems = @($diff | Where-Object { $_.SideIndicator -eq '==' })

if ($equalItems.Count -ne 1) {
    Write-Host "FAIL: expected 1 equal item with '==', got $($equalItems.Count)"
    exit 1
}

if ($equalItems[0].InputObject -ne 20) {
    Write-Host "FAIL: expected equal item value 20, got: $($equalItems[0].InputObject)"
    exit 1
}

Write-Host "PASS"
exit 0
