# vybe-test: powershell/format_list_engine/list_direct_input_object_parameter
$single = [pscustomobject]@{ City = "Kyoto"; Population = 1460000 }

# Format-List accepts -InputObject directly outside of the pipeline
$output = Format-List -InputObject $single -Property City, Population | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "City\s*:\s*Kyoto" -and $output -match "Population\s*:\s*1460000")) {
    Write-Host "FAIL: direct -InputObject failed in Format-List: $output"
    exit 1
}

Write-Host "PASS"
exit 0
