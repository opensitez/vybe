# vybe-test: powershell/get_unique_cmdlet/get_unique_single_element
# A single-element collection passes through Get-Unique without modification
$res = @("solo_element") | Get-Unique

if ($res -ne "solo_element") {
    Write-Host "FAIL: single element pass-through mismatch, got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
