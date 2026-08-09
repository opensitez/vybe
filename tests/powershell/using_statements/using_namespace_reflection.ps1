# vybe-test: powershell/using_statements/using_namespace_reflection
using namespace System.Reflection

$binding = [BindingFlags]::Instance -bor [BindingFlags]::Public
if (-not $binding.HasFlag([BindingFlags]::Public)) {
    Write-Host "FAIL: BindingFlags Public flag missing"
    exit 1
}
Write-Host "PASS"
exit 0
