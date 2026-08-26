# vybe-test: powershell/json_custom_object_roundtrip/custom_object_with_version_property
$orig = [pscustomobject]@{ AppVersion = [version]"3.2.1" }
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.AppVersion.Major -ne 3 -and $recovered.AppVersion -ne "3.2.1") {
    Write-Host "FAIL: Version property roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
