# vybe-test: powershell/split_path_cmdlet/split_path_extension_case_preservation
# -Extension preserves the original casing of the file extension
$resUpper = Split-Path "/images/banner.PNG" -Extension
$resMixed = Split-Path "/scripts/build.Ps1" -Extension

if ($resUpper -ne ".PNG") {
    Write-Host "FAIL: uppercase extension mismatch, expected '.PNG', got: '$resUpper'"
    exit 1
}

if ($resMixed -ne ".Ps1") {
    Write-Host "FAIL: mixed-case extension mismatch, expected '.Ps1', got: '$resMixed'"
    exit 1
}

Write-Host "PASS"
exit 0
