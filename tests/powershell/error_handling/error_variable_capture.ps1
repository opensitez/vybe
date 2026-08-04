# vybe-test: powershell/error_handling/error_variable_capture
Get-Item "nonexistent_path_xyz_12345" -ErrorAction SilentlyContinue -ErrorVariable myErr
if ($myErr.Count -eq 0) {
    Write-Host "FAIL: error should be captured"
    exit 1
}
Write-Host "PASS"
exit 0
