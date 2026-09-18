# vybe-test: powershell/format_list_engine/list_multiline_string_value_preservation
$obj = [pscustomobject]@{
    Message = "LineAlpha`nLineBeta`nLineGamma"
}

# Multi-line strings in property values must preserve each line in list view
$output = $obj | Format-List -Property Message | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "LineAlpha" -and $output -match "LineBeta" -and $output -match "LineGamma")) {
    Write-Host "FAIL: multi-line string content missing from list view: $output"
    exit 1
}

Write-Host "PASS"
exit 0
