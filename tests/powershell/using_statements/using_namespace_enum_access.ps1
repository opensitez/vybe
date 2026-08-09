# vybe-test: powershell/using_statements/using_namespace_enum_access
using namespace System.IO

$mode = [FileAccess]::Read
if ($mode -ne [System.IO.FileAccess]::Read) {
    Write-Host "FAIL: using namespace enum access expected Read"
    exit 1
}
Write-Host "PASS"
exit 0
