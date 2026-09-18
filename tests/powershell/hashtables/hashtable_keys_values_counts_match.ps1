# vybe-test: powershell/hashtables/hashtable_keys_values_counts_match
# The Keys and Values collection counts always equal the hashtable's overall Count property
$table = @{
    K1 = 100
    K2 = 200
    K3 = 300
    K4 = 400
}

if ($table.Count -ne 4) {
    Write-Host "FAIL: expected Count 4, got $($table.Count)"
    exit 1
}

if ($table.Keys.Count -ne 4) {
    Write-Host "FAIL: expected Keys.Count 4, got $($table.Keys.Count)"
    exit 1
}

if ($table.Values.Count -ne 4) {
    Write-Host "FAIL: expected Values.Count 4, got $($table.Values.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
