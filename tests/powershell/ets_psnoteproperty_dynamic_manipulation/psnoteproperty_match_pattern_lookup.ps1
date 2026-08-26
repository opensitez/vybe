# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_match_pattern_lookup
$obj = [pscustomobject]@{ Prefix_A = 1; Prefix_B = 2; Other = 3 }
$matched = @($obj.PSObject.Properties.Match("Prefix_*"))
if ($matched.Count -ne 2) {
    Write-Host "FAIL: PSObject.Properties.Match pattern lookup failed, got $($matched.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
