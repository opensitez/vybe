# vybe-test: powershell/format_wide_engine/wide_primitive_string_formatting
$strings = @("FirstItem", "SecondItem", "ThirdItem")

# Primitive strings format into wide columns without needing -Property
$output = $strings | Format-Wide -Column 3 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "FirstItem" -and $output -match "SecondItem" -and $output -match "ThirdItem")) {
    Write-Host "FAIL: primitive strings failed to format in wide columns: $output"
    exit 1
}

Write-Host "PASS"
exit 0
