# vybe-test: powershell/null_conditional/null_conditional_string_method
$text = "PowerShell"
$res = ${text}?.Substring(0, 5)
if ($res -ne "Power") {
    Write-Host "FAIL: null-conditional Substring expected 'Power', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
