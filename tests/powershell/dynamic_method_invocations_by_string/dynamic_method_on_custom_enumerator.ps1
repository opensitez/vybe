# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_on_custom_enumerator
$list = [System.Collections.Generic.List[int]]::new([int[]]@(10, 20))
$enum = $list.GetEnumerator()
$mMove = "MoveNext"
$null = $enum.$mMove()
$val = $enum.Current
if ($val -ne 10) {
    Write-Host "FAIL: Dynamic method on enumerator failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
