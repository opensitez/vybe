# vybe-test: powershell/format_table_engine/table_column_order_preservation
$obj = [pscustomobject]@{ Alpha = "A"; Beta = "B"; Gamma = "G" }

# Requesting properties in reverse order Gamma, Beta, Alpha
$output = $obj | Format-Table -Property Gamma, Beta, Alpha | Out-String

$gammaIndex = $output.IndexOf("Gamma")
$betaIndex  = $output.IndexOf("Beta")
$alphaIndex = $output.IndexOf("Alpha")

if ($gammaIndex -lt 0 -or $betaIndex -lt 0 -or $alphaIndex -lt 0) {
    Write-Host "FAIL: not all column headers appeared in output: $output"
    exit 1
}

# The column headers must appear in the exact requested order
if (-not ($gammaIndex -lt $betaIndex -and $betaIndex -lt $alphaIndex)) {
    Write-Host "FAIL: columns did not preserve requested order (Gamma < Beta < Alpha), got indices: Gamma=$gammaIndex, Beta=$betaIndex, Alpha=$alphaIndex"
    exit 1
}

Write-Host "PASS"
exit 0
