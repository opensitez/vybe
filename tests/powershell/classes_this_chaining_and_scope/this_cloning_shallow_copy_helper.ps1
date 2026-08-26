# vybe-test: powershell/classes_this_chaining_and_scope/this_cloning_shallow_copy_helper
class CloneablePoint {
    [int]$X; [int]$Y
    CloneablePoint([int]$x, [int]$y) { $this.X = $x; $this.Y = $y }
    [CloneablePoint]Copy() {
        return [CloneablePoint]::new($this.X, $this.Y)
    }
}
$cp1 = [CloneablePoint]::new(7, 8)
$cp2 = $cp1.Copy()
if ($cp2.X -ne 7 -or $cp2.Y -ne 8 -or $cp1 -eq $cp2) {
    Write-Host "FAIL: Copy method returning new instance via `$this failed"
    exit 1
}
Write-Host "PASS"
exit 0
