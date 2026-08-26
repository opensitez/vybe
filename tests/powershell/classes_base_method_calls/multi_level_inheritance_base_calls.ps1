# vybe-test: powershell/classes_base_method_calls/multi_level_inheritance_base_calls
class Grandparent {
    [string]Name() { return "Grandparent" }
}
class Parent : Grandparent {
    [string]Name() { return "Parent" }
}
class Child : Parent {
    [string]AllNames() {
        $gp = ([Grandparent]$this).Name()
        $p = ([Parent]$this).Name()
        return "$gp->$p->Child"
    }
}
$ch = [Child]::new()
$res = $ch.AllNames()
if ($res -ne "Grandparent->Parent->Child") {
    Write-Host "FAIL: Multi-level inheritance base calls failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
