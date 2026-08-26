# vybe-test: powershell/classes_this_chaining_and_scope/this_reference_equality_with_self
class SelfRef {
    [bool]IsSelf([SelfRef]$other) {
        return $this -eq $other
    }
}
$s1 = [SelfRef]::new()
$s2 = [SelfRef]::new()
if (-not $s1.IsSelf($s1) -or $s1.IsSelf($s2)) {
    Write-Host "FAIL: `$this reference equality test failed"
    exit 1
}
Write-Host "PASS"
exit 0
