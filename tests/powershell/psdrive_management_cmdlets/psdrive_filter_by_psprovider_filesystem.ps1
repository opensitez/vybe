# vybe-test: powershell/psdrive_management_cmdlets/psdrive_filter_by_psprovider_filesystem
# Get-PSDrive -PSProvider FileSystem filters drive results to FileSystem drives
$fsDrives = @(Get-PSDrive -PSProvider FileSystem)

if ($fsDrives.Count -lt 1) {
    Write-Host "FAIL: expected at least 1 FileSystem drive, got: $($fsDrives.Count)"
    exit 1
}

foreach ($d in $fsDrives) {
    if ($d.Provider.Name -ne "FileSystem") {
        Write-Host "FAIL: non-FileSystem drive returned: '$($d.Name)' ($($d.Provider.Name))"
        exit 1
    }
}

Write-Host "PASS"
exit 0
