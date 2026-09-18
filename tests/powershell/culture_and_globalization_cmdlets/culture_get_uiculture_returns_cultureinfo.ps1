# vybe-test: powershell/culture_and_globalization_cmdlets/culture_get_uiculture_returns_cultureinfo
# Get-UICulture returns the UI display CultureInfo of the execution host
$uic = Get-UICulture

if ($null -eq $uic) {
    Write-Host "FAIL: Get-UICulture returned `$null"
    exit 1
}

if (-not ($uic -is [System.Globalization.CultureInfo])) {
    Write-Host "FAIL: expected CultureInfo type from Get-UICulture, got: $($uic.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
