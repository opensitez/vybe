# vybe-test: powershell/type_hashcode_combiner/hashcode_add_multiple_types_sequentially
$hc = [System.HashCode]::new()
$hc.Add([int]10)
$hc.Add([double]20.5)
$hc.Add([string]"text")
$hc.Add([bool]$true)
$res = $hc.ToHashCode()
if ($res -eq 0 -and $res -ne $hc.ToHashCode()) { Write-Host "FAIL: Multi-type sequential Add failed"; exit 1 }
Write-Host "PASS"; exit 0
