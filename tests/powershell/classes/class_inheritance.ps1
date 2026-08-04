# vybe-test: powershell/classes/class_inheritance
class Shape {
    [double]$Area
    [string]Describe() { return "shape" }
}
class Circle : Shape {
    [double]$Radius
    Circle([double]$r) {
        $this.Radius = $r
        $this.Area = [Math]::PI * $r * $r
    }
    [string]Describe() { return "circle" }
}
$c = [Circle]::new(5)
if ($c.Describe() -ne "circle") { Write-Host "FAIL: describe"; exit 1 }
$expected = [Math]::Round([Math]::PI * 25, 4)
$actual   = [Math]::Round($c.Area, 4)
if ($actual -ne $expected) { Write-Host "FAIL: area $actual vs $expected"; exit 1 }
Write-Host "PASS"
exit 0
