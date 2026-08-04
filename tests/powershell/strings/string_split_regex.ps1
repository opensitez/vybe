# vybe-test: powershell/strings/string_split_regex
$csv = "one,,two,,,three"
$parts = $csv -split ",+"  # one or more commas
if ($parts.Count -ne 3)       { Write-Host "FAIL: count $($parts.Count)"; exit 1 }
if ($parts[0] -ne "one")   { Write-Host "FAIL: [0]"; exit 1 }
if ($parts[1] -ne "two")   { Write-Host "FAIL: [1]"; exit 1 }
if ($parts[2] -ne "three") { Write-Host "FAIL: [2]"; exit 1 }
Write-Host "PASS"
exit 0
