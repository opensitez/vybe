# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_10
class FinalClass_10 {
    [int]$Val = 10
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_10]::new()
if ($inst.Compute() -ne (10 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
