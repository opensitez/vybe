# vybe-test: powershell/type_conversion/double_to_int_truncates
[double]$d = 9.9
[int]$i = [int]$d
# PowerShell rounds to nearest even (banker's rounding), but explicit cast truncates
$truncated = [Math]::Truncate($d)
if ($truncated -ne 9) {
    Write-Host "FAIL: Truncate(9.9) should be 9, got $truncated"
    exit 1
}
Write-Host "PASS"
exit 0
