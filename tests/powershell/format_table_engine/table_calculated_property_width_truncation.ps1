# vybe-test: powershell/format_table_engine/table_calculated_property_width_truncation
$data = @([pscustomobject]@{ Description = "Supercalifragilisticexpialidocious" })

# Width key in calculated property constrains column width, causing truncation with unicode ellipsis
$col = @{ Label = "Desc"; Expression = { $_.Description }; Width = 8 }
$output = $data | Format-Table $col | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# The column should truncate and include the unicode ellipsis char (\u2026)
$hasEllipsis = ($output -match "\u2026") -or ($output -match "\.\.\.")
if (-not $hasEllipsis) {
    Write-Host "FAIL: truncated column did not contain ellipsis, got: $output"
    exit 1
}

# The full string must NOT appear unbroken
if ($output -match "Supercalifragilisticexpialidocious") {
    Write-Host "FAIL: long string was not truncated despite Width=8 constraint"
    exit 1
}

Write-Host "PASS"
exit 0
