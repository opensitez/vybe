# vybe-test: powershell/intrinsic_collection_overloads/where_on_single_scalar_object
# In PowerShell v3+, non-collection scalar values have intrinsic .Where() methods added automatically
$number = 75

# Match succeeds: returns the scalar in a collection of 1
$matched = $number.Where({ $_ -gt 50 })
if ($matched.Count -ne 1 -or $matched[0] -ne 75) {
    Write-Host "FAIL: scalar .Where() matching expected count 1 and value 75, got count $($matched.Count)"
    exit 1
}

# Match fails: returns an empty collection of count 0
$unmatched = $number.Where({ $_ -gt 100 })
if ($unmatched.Count -ne 0) {
    Write-Host "FAIL: scalar .Where() non-matching expected count 0, got $($unmatched.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
