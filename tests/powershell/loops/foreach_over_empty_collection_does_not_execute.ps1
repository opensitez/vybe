# vybe-test: powershell/loops/foreach_over_empty_collection_does_not_execute
# A foreach loop over an empty collection completes with zero iterations and without error
$executed = $false
$emptyCollection = @()

foreach ($element in $emptyCollection) {
    $executed = $true
}

if ($executed) {
    Write-Host "FAIL: loop body was executed on an empty array"
    exit 1
}

Write-Host "PASS"
exit 0
