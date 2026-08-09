# vybe-test: powershell/using_statements/using_namespace_multiple
using namespace System.Text
using namespace System.Collections.Generic

$sb = [StringBuilder]::new("Data")
$list = [List[int]]::new()
$list.Add(42)
if ($sb.ToString() -ne "Data" -or $list[0] -ne 42) {
    Write-Host "FAIL: multiple using namespace statements resolution failed"
    exit 1
}
Write-Host "PASS"
exit 0
