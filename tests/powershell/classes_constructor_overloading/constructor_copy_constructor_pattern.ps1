# vybe-test: powershell/classes_constructor_overloading/constructor_copy_constructor_pattern
class Point {
    [int]$X
    [int]$Y
    Point([int]$x, [int]$y) { $this.X = $x; $this.Y = $y }
    Point([Point]$other) { $this.X = $other.X; $this.Y = $other.Y }
}
$p1 = [Point]::new(10, 20)
$p2 = [Point]::new($p1)
if ($p2.X -ne 10 -or $p2.Y -ne 20) {
    Write-Host "FAIL: Copy constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
