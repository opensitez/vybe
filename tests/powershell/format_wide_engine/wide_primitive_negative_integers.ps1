# vybe-test: powershell/format_wide_engine/wide_primitive_negative_integers
$negatives = @(-10, -20, -30, -40)

# Negative integers format with minus signs in wide columns
$output = $negatives | Format-Wide -Column 2 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

foreach ($n in $negatives) {
    if (-not ($output -match [regex]::Escape("$n"))) {
        Write-Host "FAIL: negative number $n missing from wide output: $output"
        exit 1
    }
}

Write-Host "PASS"
exit 0
