# vybe-test: powershell/using_statements/using_module_class_import
using namespace System.Threading.Tasks

$t = [Task]::FromResult(100)
if ($t.Result -ne 100) {
    Write-Host "FAIL: Task.FromResult expected 100, got $($t.Result)"
    exit 1
}
Write-Host "PASS"
exit 0
