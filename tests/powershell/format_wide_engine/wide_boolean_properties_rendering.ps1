# vybe-test: powershell/format_wide_engine/wide_boolean_properties_rendering
$states = @(
    [pscustomobject]@{ IsActive = $true },
    [pscustomobject]@{ IsActive = $false }
)

# Boolean properties rendered in wide view
$output = $states | Format-Wide -Property IsActive -Column 2 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "True" -and $output -match "False")) {
    Write-Host "FAIL: boolean properties failed to render in wide view: $output"
    exit 1
}

Write-Host "PASS"
exit 0
