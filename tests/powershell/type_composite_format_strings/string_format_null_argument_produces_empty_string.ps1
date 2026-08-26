# vybe-test: powershell/type_composite_format_strings/string_format_null_argument_produces_empty_string
$res = [string]::Format("Val:{0}", [object]$null)
if ($res -ne "Val:") {
    Write-Host "FAIL: Null argument format failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
