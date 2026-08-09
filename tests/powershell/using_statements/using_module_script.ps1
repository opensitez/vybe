# vybe-test: powershell/using_statements/using_module_script
using namespace System.Collections

$al = [ArrayList]::new()
[void]$al.Add("A")
if ($al[0] -ne "A") {
    Write-Host "FAIL: using namespace System.Collections [ArrayList] failed"
    exit 1
}
Write-Host "PASS"
exit 0
