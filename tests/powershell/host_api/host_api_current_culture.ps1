# vybe-test: powershell/host_api/host_api_current_culture
if (-not [System.Globalization.CultureInfo]::CurrentCulture) {
    Write-Host "FAIL: expected current culture"
    exit 1
}
Write-Host 'PASS'
exit 0
