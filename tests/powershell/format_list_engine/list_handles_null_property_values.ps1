# vybe-test: powershell/format_list_engine/list_handles_null_property_values
$obj = [pscustomobject]@{
    DefinedProp = "PresentValue"
    EmptyProp   = $null
}

# $null property values must render the label with an empty value without throwing
$output = $obj | Format-List | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "DefinedProp\s*:\s*PresentValue")) {
    Write-Host "FAIL: DefinedProp missing: $output"
    exit 1
}

if (-not ($output -match "EmptyProp\s*:")) {
    Write-Host "FAIL: EmptyProp label missing: $output"
    exit 1
}

Write-Host "PASS"
exit 0
