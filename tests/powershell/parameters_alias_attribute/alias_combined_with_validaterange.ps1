# vybe-test: powershell/parameters_alias_attribute/alias_combined_with_validaterange
function Set-VolumeLevel {
    param(
        [Alias("Vol")]
        [ValidateRange(0, 100)]
        [int]$Volume
    )
    return $Volume
}
$res = Set-VolumeLevel -Vol 75
if ($res -ne 75) {
    Write-Host "FAIL: Alias combined with ValidateRange failed"
    exit 1
}
Write-Host "PASS"
exit 0
