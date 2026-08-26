# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_7
class FinalClass_7 {
    [int]$Val = 7
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_7]::new()
if ($inst.Compute() -ne (7 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
