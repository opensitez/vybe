# vybe-test: powershell/ast_parsing/ast_parse_script
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
