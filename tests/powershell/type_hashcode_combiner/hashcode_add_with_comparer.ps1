# vybe-test: powershell/type_hashcode_combiner/hashcode_add_with_comparer
$hc = [System.HashCode]::new()
$hc.Add("HELLO", [System.StringComparer]::OrdinalIgnoreCase)
$h1 = $hc.ToHashCode()
$hc2 = [System.HashCode]::new()
$hc2.Add("hello", [System.StringComparer]::OrdinalIgnoreCase)
$h2 = $hc2.ToHashCode()
if ($h1 -ne $h2) { Write-Host "FAIL: HashCode Add with StringComparer failed"; exit 1 }
Write-Host "PASS"; exit 0
