# vybe-test: powershell/json_custom_object_roundtrip/pscustomobject_with_unicode_characters
$orig = [pscustomobject]@{ City = "Z`u{00FC}rich"; Symbol = "`u{2764}" }
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.City -ne "Z`u{00FC}rich" -or $recovered.Symbol -ne "`u{2764}") {
    Write-Host "FAIL: Unicode characters roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
