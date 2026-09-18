# vybe-test: powershell/format_list_engine/list_numeric_properties_rendering
$data = [pscustomobject]@{
    IntegerVal = 1048576
    DoubleVal  = 3.14159
    Negative   = -99
}

# Numeric properties format as clean string representations in list view
$output = $data | Format-List | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "IntegerVal\s*:\s*1048576" -and 
          $output -match "DoubleVal\s*:\s*3\.14159" -and 
          $output -match "Negative\s*:\s*-99")) {
    Write-Host "FAIL: numeric properties failed to render: $output"
    exit 1
}

Write-Host "PASS"
exit 0
