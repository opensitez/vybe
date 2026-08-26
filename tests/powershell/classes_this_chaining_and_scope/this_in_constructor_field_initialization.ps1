# vybe-test: powershell/classes_this_chaining_and_scope/this_in_constructor_field_initialization
class Coord {
    [int]$X; [int]$Y
    Coord([int]$x, [int]$y) {
        $this.X = $x
        $this.Y = $y
    }
}
$c = [Coord]::new(100, 200)
if ($c.X -ne 100 -or $c.Y -ne 200) {
    Write-Host "FAIL: Constructor field init via `$this failed"
    exit 1
}
Write-Host "PASS"
exit 0
