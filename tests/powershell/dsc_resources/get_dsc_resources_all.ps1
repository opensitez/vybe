# vybe-test: powershell/dsc_resources/get_dsc_resources_all
$resources = Get-DscResource -ErrorAction SilentlyContinue
if ($resources.Count -lt 1) {
    Write-Host "FAIL: expected at least one DSC resource"
    exit 1
}
Write-Host 'PASS'
exit 0
