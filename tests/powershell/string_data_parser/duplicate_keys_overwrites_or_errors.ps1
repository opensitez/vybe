# vybe-test: powershell/string_data_parser/duplicate_keys_overwrites_or_errors
$str = @"
dup = first
dup = second
"@
# In PowerShell ConvertFrom-StringData, duplicate key causes error or overwrite depending on delimiter mode
$caughtOrParsed = $false
try {
    $ht = ConvertFrom-StringData -StringData $str -ErrorAction Stop
    $caughtOrParsed = ($ht["dup"] -eq "second" -or $ht["dup"] -eq "first")
} catch {
    $caughtOrParsed = $true # throwing is also valid PowerShell behavior
}
if (-not $caughtOrParsed) {
    Write-Host "FAIL: Duplicate key handling failed"
    exit 1
}
Write-Host "PASS"
exit 0
