# vybe-test: powershell/classes_polymorphism_and_casting/invalid_downcast_throws_invalid_cast_exception
class ParentA {}
class ChildA : ParentA {}
class ChildB : ParentA {}
$ca = [ChildA]::new()
[ParentA]$p = $ca
$caught = $false
try {
    $bad = [ChildB]$p
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception on invalid downcast"
    exit 1
}
Write-Host "PASS"
exit 0
