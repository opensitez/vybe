# vybe-test: powershell/parameters_validate_count/validatecount_error_message_contains_counts
function Check-Counts {
    param([ValidateCount(3, 5)][int[]]$Nums)
    return $Nums
}
$msg = ""
try {
    $x = Check-Counts -Nums 1, 2 # count 2 < 3
} catch {
    $msg = $_.Exception.Message
}
if (-not ($msg.Contains("3") -and $msg.Contains("5") -and $msg.Contains("2"))) {
    Write-Host "FAIL: Error message should contain min, max and actual counts, got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
