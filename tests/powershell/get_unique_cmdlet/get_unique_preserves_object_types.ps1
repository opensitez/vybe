# vybe-test: powershell/get_unique_cmdlet/get_unique_preserves_object_types
# Get-Unique preserves the original .NET types of stream objects
$mixed = @([int]10, [int]10, [double]20.5, [double]20.5)
$unique = @($mixed | Get-Unique)

if ($unique.Count -ne 2) {
    Write-Host "FAIL: expected 2 unique elements, got $($unique.Count)"
    exit 1
}

if ($unique[0] -isnot [int] -or $unique[1] -isnot [double]) {
    Write-Host "FAIL: output object types not preserved, got $($unique[0].GetType().FullName), $($unique[1].GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
