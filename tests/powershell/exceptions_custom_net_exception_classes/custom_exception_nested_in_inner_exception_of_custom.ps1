# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_nested_in_inner_exception_of_custom
class LevelOneEx : System.Exception { LevelOneEx([string]$m) : base($m) {} }
class LevelTwoEx : System.Exception { LevelTwoEx([string]$m, [System.Exception]$i) : base($m, $i) {} }
$l1 = [LevelOneEx]::new("Inner")
$l2 = [LevelTwoEx]::new("Outer", $l1)
if ($l2.InnerException -isnot [LevelOneEx]) {
    Write-Host "FAIL: Nested custom exceptions failed"
    exit 1
}
Write-Host "PASS"
exit 0
