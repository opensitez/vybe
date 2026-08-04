# vybe-test: powershell/collections/generic_dictionary
$dict = [System.Collections.Generic.Dictionary[string,int]]::new()
$dict.Add("one", 1)
$dict.Add("two", 2)
$dict.Add("three", 3)
if ($dict.Count -ne 3) { Write-Host "FAIL: count"; exit 1 }
if ($dict["two"] -ne 2) { Write-Host "FAIL: lookup"; exit 1 }
if (-not $dict.ContainsKey("three")) { Write-Host "FAIL: ContainsKey"; exit 1 }
$dict.Remove("one")
if ($dict.ContainsKey("one")) { Write-Host "FAIL: should be removed"; exit 1 }
Write-Host "PASS"
exit 0
