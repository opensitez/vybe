# vybe-test: powershell/type_hashcode_combiner/hashcode_gethashcode_method_throws_notsupported
$hc = [System.HashCode]::new()
$caught = $false
try {
    $x = $hc.GetHashCode()
} catch [System.NotSupportedException] {
    $caught = $true
}
if (-not $caught) { Write-Host "FAIL: NotSupportedException expected on HashCode.GetHashCode"; exit 1 }
Write-Host "PASS"; exit 0
