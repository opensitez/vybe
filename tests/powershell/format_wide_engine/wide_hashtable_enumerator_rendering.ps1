# vybe-test: powershell/format_wide_engine/wide_hashtable_enumerator_rendering
$hash = @{
    KeyAlpha = "Val1"
    KeyBeta  = "Val2"
}

# Hashtable enumerator formatted in wide view selecting Name
$output = $hash.GetEnumerator() | Format-Wide -Property Name -Column 2 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "KeyAlpha" -and $output -match "KeyBeta")) {
    Write-Host "FAIL: hashtable enumerator failed in wide view: $output"
    exit 1
}

Write-Host "PASS"
exit 0
