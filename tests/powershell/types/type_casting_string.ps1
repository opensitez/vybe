# vybe-test: powershell/types/type_casting_string
$x = [string]123
if ($x -ne "123") {
    Write-Host "FAIL: expected '123', got '$x'"
    exit 1
}
Write-Host "PASS"
exit 0
