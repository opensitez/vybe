# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/catch_block_returning_value
function Get-SafeNumber([string]$str) {
    try {
        return [int]::Parse($str)
    } catch [System.FormatException] {
        return -1
    }
}
$r1 = Get-SafeNumber "100"
$r2 = Get-SafeNumber "invalid"
if ($r1 -ne 100 -or $r2 -ne -1) {
    Write-Host "FAIL: Catch block returning value failed, r1=$r1, r2=$r2"
    exit 1
}
Write-Host "PASS"
exit 0
