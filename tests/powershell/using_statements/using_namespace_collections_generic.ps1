# vybe-test: powershell/using_statements/using_namespace_collections_generic
using namespace System.Collections.Generic

$list = [List[string]]::new()
$list.Add("Item")
if ($list[0] -ne "Item") {
    Write-Host "FAIL: using namespace System.Collections.Generic [List[string]] failed"
    exit 1
}
Write-Host "PASS"
exit 0
