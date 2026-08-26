# vybe-test: powershell/dynamic_assembly_type_resolution/resolve_type_case_insensitively
$type1 = [type]"system.text.stringbuilder"
$type2 = [type]"SYSTEM.TEXT.STRINGBUILDER"
if ($type1 -ne [System.Text.StringBuilder] -or $type2 -ne [System.Text.StringBuilder]) {
    Write-Host "FAIL: Case-insensitive type resolution failed"
    exit 1
}
Write-Host "PASS"
exit 0
