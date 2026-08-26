# vybe-test: powershell/collections_sorted_dictionary/guid_keys_sorting
$g1 = [guid]::Parse("00000000-0000-0000-0000-000000000001")
$g2 = [guid]::Parse("00000000-0000-0000-0000-000000000002")
$sd = [System.Collections.Generic.SortedDictionary[guid, string]]::new()
$sd.Add($g2, "second"); $sd.Add($g1, "first")
$keys = @($sd.Keys)
if ($keys[0] -ne $g1 -or $keys[1] -ne $g2) {
    Write-Host "FAIL: Guid keys sorting failed"
    exit 1
}
Write-Host "PASS"
exit 0
