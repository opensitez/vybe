# vybe-test: powershell/format_wide_engine/wide_direct_input_object_parameter
$single = [pscustomobject]@{ Code = "ALPHA_77" }

# Format-Wide accepts -InputObject explicitly outside of pipeline streaming
$output = Format-Wide -InputObject $single -Property Code | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "ALPHA_77")) {
    Write-Host "FAIL: direct -InputObject failed in Format-Wide: $output"
    exit 1
}

Write-Host "PASS"
exit 0
