# vybe-test: powershell/new_object_cmdlet/new_object_guid_tostring_matches_format
# New-Object System.Guid -ArgumentList <guid string> round-trips through ToString()
$guidStr = "12345678-1234-1234-1234-123456789abc"
$guid = New-Object System.Guid -ArgumentList $guidStr

if ($guid.ToString() -ne $guidStr) {
    Write-Host "FAIL: GUID round-trip mismatch: '$($guid.ToString())'"
    exit 1
}

Write-Host "PASS"
exit 0
