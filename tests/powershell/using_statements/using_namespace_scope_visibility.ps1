# vybe-test: powershell/using_statements/using_namespace_scope_visibility
using namespace System.Collections.Generic

function Test-Scope {
    $l = [List[int]]::new()
    $l.Add(9)
    return $l[0]
}
$res = Test-Scope
if ($res -ne 9) {
    Write-Host "FAIL: using namespace in function scope expected 9, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
