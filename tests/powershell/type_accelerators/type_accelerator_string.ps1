# vybe-test: powershell/type_accelerators/type_accelerator_string
$s = [string]12345
if ($s -ne "12345") {
    Write-Host "FAIL: string expected '12345', got '$s'"
    exit 1
}
Write-Host "PASS"
exit 0
