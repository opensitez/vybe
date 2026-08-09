# vybe-test: powershell/using_statements/using_namespace_duplicate_ignored
using namespace System.Text
using namespace System.Text

$sb = [StringBuilder]::new("NoDup")
if ($sb.ToString() -ne "NoDup") {
    Write-Host "FAIL: duplicate using namespace statements error"
    exit 1
}
Write-Host "PASS"
exit 0
