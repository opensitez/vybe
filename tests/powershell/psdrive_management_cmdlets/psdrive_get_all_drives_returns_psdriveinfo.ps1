# vybe-test: powershell/psdrive_management_cmdlets/psdrive_get_all_drives_returns_psdriveinfo
# Get-PSDrive returns instances of System.Management.Automation.PSDriveInfo representing mounted provider drives
$drives = Get-PSDrive

if ($drives.Count -lt 4) {
    Write-Host "FAIL: expected at least 4 default PSDrives, got: $($drives.Count)"
    exit 1
}

if (-not ($drives[0] -is [System.Management.Automation.PSDriveInfo])) {
    Write-Host "FAIL: expected PSDriveInfo type, got: $($drives[0].GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
