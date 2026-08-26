# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_loop_over_multiple_properties
$props = @("First", "Last", "Role")
$user = [pscustomobject]@{ First = "Alice"; Last = "Smith"; Role = "Admin" }
$collected = [System.Collections.Generic.List[string]]::new()
foreach ($p in $props) {
    $collected.Add($user.$p)
}
if ($collected.Count -ne 3 -or $collected[0] -ne "Alice" -or $collected[2] -ne "Admin") {
    Write-Host "FAIL: Dynamic property loop extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
