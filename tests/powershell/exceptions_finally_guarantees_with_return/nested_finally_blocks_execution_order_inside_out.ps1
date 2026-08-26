# vybe-test: powershell/exceptions_finally_guarantees_with_return/nested_finally_blocks_execution_order_inside_out
$order = [System.Collections.Generic.List[string]]::new()
function Test-NestedFinally {
    try {
        try {
            return "FromInner"
        } finally {
            $order.Add("InnerFinally")
        }
    } finally {
        $order.Add("OuterFinally")
    }
}
$res = Test-NestedFinally
if ($res -ne "FromInner" -or $order.Count -ne 2 -or $order[0] -ne "InnerFinally" -or $order[1] -ne "OuterFinally") {
    Write-Host "FAIL: Nested finally execution order failed, got $($order -join '->')"
    exit 1
}
Write-Host "PASS"
exit 0
