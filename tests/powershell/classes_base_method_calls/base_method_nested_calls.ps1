# vybe-test: powershell/classes_base_method_calls/base_method_nested_calls
class Level1 {
    [int]Value() { return 10 }
}
class Level2 : Level1 {
    [int]Value() { return ([Level1]$this).Value() + 20 }
}
class Level3 : Level2 {
    [int]Value() { return ([Level2]$this).Value() + 30 }
}
$l3 = [Level3]::new()
$res = $l3.Value()
if ($res -ne 60) {
    Write-Host "FAIL: Nested base method calls expected 60, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
