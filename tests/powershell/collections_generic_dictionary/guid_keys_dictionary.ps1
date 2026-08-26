# vybe-test: powershell/collections_generic_dictionary/guid_keys_dictionary
$g1 = [guid]::NewGuid()
$g2 = [guid]::NewGuid()
$d = [System.Collections.Generic.Dictionary[guid, string]]::new()
$d.Add($g1, "first")
$d.Add($g2, "second")
if ($d[$g1] -ne "first" -or $d[$g2] -ne "second") {
    Write-Host "FAIL: Guid keys dictionary failed"
    exit 1
}
Write-Host "PASS"
exit 0
