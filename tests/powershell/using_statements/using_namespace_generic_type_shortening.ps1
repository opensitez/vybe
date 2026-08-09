# vybe-test: powershell/using_statements/using_namespace_generic_type_shortening
using namespace System.Collections.Generic

$dict = [Dictionary[string, List[int]]]::new()
$dict["Evens"] = [List[int]]::new()
$dict["Evens"].Add(2)
if ($dict["Evens"][0] -ne 2) {
    Write-Host "FAIL: nested generic type shortening via using namespace failed"
    exit 1
}
Write-Host "PASS"
exit 0
