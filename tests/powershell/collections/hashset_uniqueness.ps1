# vybe-test: powershell/collections/hashset_uniqueness
$set = [System.Collections.Generic.HashSet[string]]::new()
$set.Add("a") | Out-Null
$set.Add("b") | Out-Null
$set.Add("a") | Out-Null  # duplicate
if ($set.Count -ne 2) { Write-Host "FAIL: expected 2 unique, got $($set.Count)"; exit 1 }
if (-not $set.Contains("b")) { Write-Host "FAIL: missing 'b'"; exit 1 }
$set.Remove("a") | Out-Null
if ($set.Contains("a")) { Write-Host "FAIL: 'a' should be removed"; exit 1 }
Write-Host "PASS"
exit 0
