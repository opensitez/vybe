# vybe-test: powershell/json_custom_object_roundtrip/pscustomobject_with_escaped_string_values
$orig = [pscustomobject]@{ Text = "Line 1`nLine 2`tTabbed `"Quotes`"" }
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Text -ne "Line 1`nLine 2`tTabbed `"Quotes`"") {
    Write-Host "FAIL: Escaped string values roundtrip failed, got '$($recovered.Text)'"
    exit 1
}
Write-Host "PASS"
exit 0
