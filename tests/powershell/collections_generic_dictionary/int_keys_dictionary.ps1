# vybe-test: powershell/collections_generic_dictionary/int_keys_dictionary
$d = [System.Collections.Generic.Dictionary[int, string]]::new()
$d.Add(101, "Server1")
$d.Add(102, "Server2")
if ($d[101] -ne "Server1" -or $d[102] -ne "Server2") {
    Write-Host "FAIL: Integer keys dictionary failed"
    exit 1
}
Write-Host "PASS"
exit 0
