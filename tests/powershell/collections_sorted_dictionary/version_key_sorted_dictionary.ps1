# vybe-test: powershell/collections_sorted_dictionary/version_key_sorted_dictionary
$sd = [System.Collections.Generic.SortedDictionary[version, string]]::new()
$sd.Add([version]"2.0", "v2"); $sd.Add([version]"1.0", "v1"); $sd.Add([version]"1.5", "v1.5")
$keys = @($sd.Keys)
if ($keys[0] -ne [version]"1.0" -or $keys[1] -ne [version]"1.5" -or $keys[2] -ne [version]"2.0") {
    Write-Host "FAIL: Version keys sorting failed"
    exit 1
}
Write-Host "PASS"
exit 0
