# vybe-test: powershell/classes_custom_methods_overloading/overload_with_char_and_string
class CharOverload {
    [string]Echo([char]$c) { return "CHAR:$c" }
    [string]Echo([string]$s) { return "STR:$s" }
}
$co = [CharOverload]::new()
$r1 = $co.Echo([char]'A')
$r2 = $co.Echo("A")
if ($r1 -ne "CHAR:A" -or $r2 -ne "STR:A") {
    Write-Host "FAIL: Char vs String overload resolution failed"
    exit 1
}
Write-Host "PASS"
exit 0
