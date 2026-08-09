# vybe-test: powershell/using_statements/using_namespace_case_insensitivity
using namespace system.text

$sb = [stringbuilder]::new("test")
if ($sb.ToString() -ne "test") {
    Write-Host "FAIL: case-insensitive using namespace system.text failed"
    exit 1
}
Write-Host "PASS"
exit 0
