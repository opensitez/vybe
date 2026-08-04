# vybe-test: powershell/profile_loading/profile_folder_path.ps1
$folder = Split-Path $PROFILE
if (-not $folder) {
    Write-Host "FAIL: expected folder for profile"
    exit 1
}
Write-Host 'PASS'
exit 0
