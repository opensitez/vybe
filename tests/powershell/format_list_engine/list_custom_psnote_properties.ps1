# vybe-test: powershell/format_list_engine/list_custom_psnote_properties
$obj = [pscustomobject]@{ StaticField = "StaticData" }

# Attaching dynamic NoteProperty via Add-Member
$obj | Add-Member -NotePropertyName "DynamicInfo" -NotePropertyValue "AttachedDynamically"

$output = $obj | Format-List | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "StaticField\s*:\s*StaticData" -and $output -match "DynamicInfo\s*:\s*AttachedDynamically")) {
    Write-Host "FAIL: dynamic NoteProperty missing from list output: $output"
    exit 1
}

Write-Host "PASS"
exit 0
