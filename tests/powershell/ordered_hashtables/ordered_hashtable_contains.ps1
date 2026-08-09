# vybe-test: powershell/ordered_hashtables/ordered_hashtable_contains
$h = [ordered]@{ KeyOne = "val" }
if (-not $h.Contains("KeyOne")) {
    Write-Host "FAIL: Contains('KeyOne') expected true"
    exit 1
}
if ($h.Contains("MissingKey")) {
    Write-Host "FAIL: Contains('MissingKey') expected false"
    exit 1
}
Write-Host "PASS"
exit 0
