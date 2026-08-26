# vybe-test: powershell/type_hashcode_combiner/hashcode_struct_add_fluent_pattern
$hc = [System.HashCode]::new()
$hc.Add("Key")
$hc.Add(123)
$hash = $hc.ToHashCode()
if ($hash -eq 0 -and $hash -ne $hc.ToHashCode()) { Write-Host "FAIL: HashCode struct Add failed"; exit 1 }
Write-Host "PASS"; exit 0
