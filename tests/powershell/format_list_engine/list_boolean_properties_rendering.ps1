# vybe-test: powershell/format_list_engine/list_boolean_properties_rendering
$state = [pscustomobject]@{
    IsEnabled   = $true
    IsSuspended = $false
}

# Boolean properties format as True and False
$output = $state | Format-List | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "IsEnabled\s*:\s*True" -and $output -match "IsSuspended\s*:\s*False")) {
    Write-Host "FAIL: boolean properties failed to render: $output"
    exit 1
}

Write-Host "PASS"
exit 0
