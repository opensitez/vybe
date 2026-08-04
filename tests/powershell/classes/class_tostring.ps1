# vybe-test: powershell/classes/class_tostring
class Point {
    [int]$X
    [int]$Y
    Point([int]$x, [int]$y) { $this.X = $x; $this.Y = $y }
    [string]ToString() { return "($($this.X), $($this.Y))" }
}
$pt = [Point]::new(3, 4)
$str = $pt.ToString()
if ($str -ne "(3, 4)") {
    Write-Host "FAIL: expected '(3, 4)', got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
