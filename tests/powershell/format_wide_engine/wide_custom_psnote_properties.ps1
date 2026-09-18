# vybe-test: powershell/format_wide_engine/wide_custom_psnote_properties
$obj = [pscustomobject]@{ Static = "Init" }

# Attach dynamic NoteProperty
$obj | Add-Member -NotePropertyName "WideDynamicProp" -NotePropertyValue "DynamicWideVal"

$output = $obj | Format-Wide -Property WideDynamicProp | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "DynamicWideVal")) {
    Write-Host "FAIL: dynamic NoteProperty missing from wide output: $output"
    exit 1
}

Write-Host "PASS"
exit 0
