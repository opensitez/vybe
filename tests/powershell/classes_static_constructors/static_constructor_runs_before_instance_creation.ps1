# vybe-test: powershell/classes_static_constructors/static_constructor_runs_before_instance_creation
class OrderTest {
    static [int]$StaticCount = 0
    [int]$InstanceId
    static OrderTest() {
        [OrderTest]::StaticCount = 100
    }
    OrderTest() {
        $this.InstanceId = [OrderTest]::StaticCount + 1
    }
}
$obj = [OrderTest]::new()
if ($obj.InstanceId -ne 101) {
    Write-Host "FAIL: Static constructor order failed, got $($obj.InstanceId)"
    exit 1
}
Write-Host "PASS"
exit 0
