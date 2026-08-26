# vybe-test: powershell/dynamic_method_invocations_by_string/invoke_list_method_dynamically
$list = [System.Collections.Generic.List[int]]::new()
$addMethod = "Add"
$list.$addMethod(10)
$list.$addMethod(20)
if ($list.Count -ne 2 -or $list[0] -ne 10) {
    Write-Host "FAIL: Dynamic List method invocation failed"
    exit 1
}
Write-Host "PASS"
exit 0
