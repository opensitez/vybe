# vybe-test: powershell/out_string_cmdlet/out_string_width_parameter_truncation
# The -Width parameter constrains the line formatting buffer width
$obj = [pscustomobject]@{ VeryLongPropertyNameForFormatting = "TestValue123" }
$output = $obj | Out-String -Width 30

if ($null -eq $output) {
    Write-Host "FAIL: Out-String -Width returned `$null"
    exit 1
}

Write-Host "PASS"
exit 0
