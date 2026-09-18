# vybe-test: powershell/new_object_cmdlet/new_object_version_components_accessible
# New-Object System.Version -ArgumentList "major.minor.build" exposes Major/Minor/Build properties
$ver = New-Object System.Version -ArgumentList "3", "14", "159"

if ($ver.Major -ne 3 -or $ver.Minor -ne 14 -or $ver.Build -ne 159) {
    Write-Host "FAIL: Version components mismatch: $($ver.ToString())"
    exit 1
}

Write-Host "PASS"
exit 0
