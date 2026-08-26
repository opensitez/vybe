# vybe-test: powershell/string_normalization_forms/sort_normalized_strings
$arr = @("cafe`u{0301}", "banana", "apple") | ForEach-Object { $_.Normalize() }
$sorted = $arr | Sort-Object
if ($sorted[0] -ne "apple" -or $sorted[1] -ne "banana" -or $sorted[2] -ne "caf`u{00E9}") {
    Write-Host "FAIL: Normalized strings sorting order failed"
    exit 1
}
Write-Host "PASS"
exit 0
