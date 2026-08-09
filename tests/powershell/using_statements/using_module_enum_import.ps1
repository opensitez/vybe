# vybe-test: powershell/using_statements/using_module_enum_import
using namespace System.Diagnostics

$pri = [ProcessPriorityClass]::Normal
if ($pri -ne [System.Diagnostics.ProcessPriorityClass]::Normal) {
    Write-Host "FAIL: ProcessPriorityClass enum expected Normal"
    exit 1
}
Write-Host "PASS"
exit 0
