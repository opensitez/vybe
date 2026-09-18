# vybe-test: powershell/compare_object_cmdlet/compare_object_multiple_properties_comparison
# Multiple properties can be passed to -Property to define composite equality
$item1 = [pscustomobject]@{ Org = "Dev"; Level = 3; Note = "Old" }
$item2 = [pscustomobject]@{ Org = "Dev"; Level = 4; Note = "New" }

$diff = Compare-Object @($item1) @($item2) -Property Org, Level

if ($diff.Count -ne 2) {
    Write-Host "FAIL: expected 2 differences across composite keys, got $($diff.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
