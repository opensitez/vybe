# vybe-test: powershell/collections_generic_hashset/guid_hashset_deduping
$g1 = [guid]::NewGuid()
$g2 = [guid]::NewGuid()
$set = [System.Collections.Generic.HashSet[guid]]::new()
$set.Add($g1); $set.Add($g2); $set.Add($g1)
if ($set.Count -ne 2) {
    Write-Host "FAIL: Guid HashSet deduping failed"
    exit 1
}
Write-Host "PASS"
exit 0
