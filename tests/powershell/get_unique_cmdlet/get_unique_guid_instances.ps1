# vybe-test: powershell/get_unique_cmdlet/get_unique_guid_instances
# Consecutive identical GUID values are collapsed by value comparison
$g1 = [Guid]::Parse("11111111-1111-1111-1111-111111111111")
$g2 = [Guid]::Parse("11111111-1111-1111-1111-111111111111")
$g3 = [Guid]::Parse("22222222-2222-2222-2222-222222222222")

$unique = @($g1, $g2, $g3 | Get-Unique)

if ($unique.Count -ne 2) {
    Write-Host "FAIL: expected 2 unique GUIDs, got $($unique.Count)"
    exit 1
}

if ($unique[0] -ne $g1 -or $unique[1] -ne $g3) {
    Write-Host "FAIL: unique GUID values mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
