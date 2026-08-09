# vybe-test: powershell/using_statements/using_namespace_io
using namespace System.IO

$ms = [MemoryStream]::new()
if ($ms.CanWrite -ne $true) {
    Write-Host "FAIL: using namespace System.IO MemoryStream expected CanWrite=true"
    exit 1
}
Write-Host "PASS"
exit 0
