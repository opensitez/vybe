# vybe-test: powershell/hashtables/hashtable_merge
$defaults = @{ color = "blue"; size = "M"; weight = 1.0 }
$overrides = @{ color = "red"; weight = 2.5 }
# Merge: overrides win
$merged = $defaults.Clone()
foreach ($key in $overrides.Keys) { $merged[$key] = $overrides[$key] }
if ($merged.color  -ne "red")  { Write-Host "FAIL: color";  exit 1 }
if ($merged.size   -ne "M")    { Write-Host "FAIL: size";   exit 1 }
if ($merged.weight -ne 2.5)    { Write-Host "FAIL: weight"; exit 1 }
Write-Host "PASS"
exit 0
