# vybe-test: powershell/out_string_cmdlet/out_string_large_width_prevents_wrapping
# A wide buffer width (e.g. -Width 300) prevents premature truncation and line wrapping of long strings
$longStr = "DataSequence_" + ("A" * 150)
$obj = [pscustomobject]@{ ExtensiveField = $longStr }
$output = $obj | Out-String -Width 300

if ($output -notmatch $longStr) {
    Write-Host "FAIL: wide buffer prematurely wrapped or truncated string"
    exit 1
}

Write-Host "PASS"
exit 0
